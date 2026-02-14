from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable


ConditionOperator = str
StepHandler = Callable[[dict[str, Any]], dict[str, Any]]


@dataclass(frozen=True)
class ChainCondition:
    source_node_id: str
    metric: str
    operator: ConditionOperator
    threshold: float


@dataclass(frozen=True)
class ChainNode:
    node_id: str
    step: str
    depends_on: list[str]
    conditions: list[ChainCondition]


def _to_float(value: Any) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _evaluate_condition(condition: ChainCondition, node_outputs: dict[str, dict[str, Any]]) -> tuple[bool, str]:
    source_payload = node_outputs.get(condition.source_node_id, {})
    metric_value = _to_float(source_payload.get("metrics", {}).get(condition.metric))
    threshold = float(condition.threshold)
    operator = condition.operator

    passed = False
    if operator == ">":
        passed = metric_value > threshold
    elif operator == ">=":
        passed = metric_value >= threshold
    elif operator == "<":
        passed = metric_value < threshold
    elif operator == "<=":
        passed = metric_value <= threshold
    elif operator == "==":
        passed = metric_value == threshold

    reason = (
        f"Condition {condition.source_node_id}.{condition.metric} {operator} {threshold:.6f} "
        f"evaluated with value {metric_value:.6f}: {'pass' if passed else 'fail'}"
    )
    return passed, reason


def _load_chain_nodes(chain_config: dict[str, Any]) -> list[ChainNode]:
    nodes: list[ChainNode] = []
    for row in chain_config.get("nodes", []):
        conditions = [
            ChainCondition(
                source_node_id=str(item["source_node_id"]),
                metric=str(item["metric"]),
                operator=str(item["operator"]),
                threshold=float(item["threshold"]),
            )
            for item in row.get("conditions", [])
        ]
        nodes.append(
            ChainNode(
                node_id=str(row["node_id"]),
                step=str(row["step"]),
                depends_on=[str(item) for item in row.get("depends_on", [])],
                conditions=conditions,
            )
        )
    return nodes


def run_chain_from_manifest(*, project_manifest_path: Path, handlers: dict[str, StepHandler]) -> dict[str, Any]:
    manifest = json.loads(project_manifest_path.read_text(encoding="utf-8"))
    chain_config = dict(manifest.get("execution_chain", {}))
    chain_nodes = _load_chain_nodes(chain_config)

    node_outputs: dict[str, dict[str, Any]] = {}
    execution_trace: list[dict[str, Any]] = []

    for node in chain_nodes:
        if any(dep not in node_outputs for dep in node.depends_on):
            execution_trace.append(
                {
                    "node_id": node.node_id,
                    "step": node.step,
                    "status": "blocked",
                    "reason": "dependency_missing",
                }
            )
            continue

        condition_results = [_evaluate_condition(condition, node_outputs) for condition in node.conditions]
        failed_reasons = [reason for passed, reason in condition_results if not passed]
        if failed_reasons:
            execution_trace.append(
                {
                    "node_id": node.node_id,
                    "step": node.step,
                    "status": "skipped",
                    "reason": "conditions_failed",
                    "details": failed_reasons,
                }
            )
            continue

        handler = handlers.get(node.step)
        if handler is None:
            execution_trace.append(
                {
                    "node_id": node.node_id,
                    "step": node.step,
                    "status": "blocked",
                    "reason": "handler_missing",
                }
            )
            continue

        output = handler({"node_id": node.node_id, "step": node.step, "node_outputs": node_outputs})
        payload = dict(output if isinstance(output, dict) else {})
        node_outputs[node.node_id] = payload
        execution_trace.append(
            {
                "node_id": node.node_id,
                "step": node.step,
                "status": "completed",
                "metrics": dict(payload.get("metrics", {})),
            }
        )

    manifest["execution_chain"] = {
        **chain_config,
        "last_execution": {
            "trace": execution_trace,
            "node_outputs": node_outputs,
        },
    }
    project_manifest_path.write_text(json.dumps(manifest, indent=2, sort_keys=True), encoding="utf-8")
    return {"trace": execution_trace, "node_outputs": node_outputs}


def build_default_research_execution_chain(*, sharpe_threshold: float = 0.8, drawdown_limit: float = 0.25) -> dict[str, Any]:
    return {
        "schema_version": "1.0",
        "mode": "dag",
        "nodes": [
            {
                "node_id": "walk_forward",
                "step": "walk_forward",
                "depends_on": [],
                "conditions": [],
            },
            {
                "node_id": "optimization",
                "step": "optimization",
                "depends_on": ["walk_forward"],
                "conditions": [
                    {
                        "source_node_id": "walk_forward",
                        "metric": "sharpe",
                        "operator": ">",
                        "threshold": float(sharpe_threshold),
                    },
                    {
                        "source_node_id": "walk_forward",
                        "metric": "max_drawdown",
                        "operator": ">",
                        "threshold": -float(drawdown_limit),
                    },
                ],
            },
            {
                "node_id": "stress",
                "step": "stress",
                "depends_on": ["optimization"],
                "conditions": [],
            },
            {
                "node_id": "governance_evaluation",
                "step": "governance_evaluation",
                "depends_on": ["stress"],
                "conditions": [
                    {
                        "source_node_id": "stress",
                        "metric": "stress_pass_rate",
                        "operator": ">=",
                        "threshold": 0.8,
                    }
                ],
            },
        ],
    }

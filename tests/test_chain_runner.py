from __future__ import annotations

import json

from backtesting.chain_runner import build_default_research_execution_chain, run_chain_from_manifest


def test_run_chain_from_manifest_executes_all_nodes_when_conditions_pass(tmp_path):
    manifest_path = tmp_path / "research_project.json"
    manifest_path.write_text(
        json.dumps(
            {
                "hypothesis": {"hypothesis_id": "hyp_1"},
                "execution_chain": build_default_research_execution_chain(sharpe_threshold=0.8, drawdown_limit=0.25),
            }
        ),
        encoding="utf-8",
    )

    handlers = {
        "walk_forward": lambda _ctx: {"metrics": {"sharpe": 1.1, "max_drawdown": -0.10}},
        "optimization": lambda _ctx: {"metrics": {"best_score": 1.2}},
        "stress": lambda _ctx: {"metrics": {"stress_pass_rate": 0.9}},
        "governance_evaluation": lambda _ctx: {"metrics": {"promotion_ready": 1.0}},
    }

    result = run_chain_from_manifest(project_manifest_path=manifest_path, handlers=handlers)

    completed = [row for row in result["trace"] if row["status"] == "completed"]
    assert [row["node_id"] for row in completed] == [
        "walk_forward",
        "optimization",
        "stress",
        "governance_evaluation",
    ]

    updated_manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert "last_execution" in updated_manifest["execution_chain"]


def test_run_chain_from_manifest_skips_downstream_on_failed_condition(tmp_path):
    manifest_path = tmp_path / "research_project.json"
    manifest_path.write_text(
        json.dumps(
            {
                "hypothesis": {"hypothesis_id": "hyp_2"},
                "execution_chain": build_default_research_execution_chain(sharpe_threshold=1.0, drawdown_limit=0.15),
            }
        ),
        encoding="utf-8",
    )

    handlers = {
        "walk_forward": lambda _ctx: {"metrics": {"sharpe": 0.7, "max_drawdown": -0.30}},
        "optimization": lambda _ctx: {"metrics": {"best_score": 1.2}},
        "stress": lambda _ctx: {"metrics": {"stress_pass_rate": 0.95}},
        "governance_evaluation": lambda _ctx: {"metrics": {"promotion_ready": 1.0}},
    }

    result = run_chain_from_manifest(project_manifest_path=manifest_path, handlers=handlers)

    statuses = {row["node_id"]: row["status"] for row in result["trace"]}
    assert statuses["walk_forward"] == "completed"
    assert statuses["optimization"] == "skipped"
    assert statuses["stress"] == "blocked"
    assert statuses["governance_evaluation"] == "blocked"


def test_write_experiment_skeleton_includes_execution_chain(tmp_path):
    from ui.research_lab_page import ResearchLabPage
    from tests.test_research_lab_manifest import _build_context, _build_record

    page = object.__new__(ResearchLabPage)
    page._research_lab_dir = tmp_path / "research_lab"

    output_path = ResearchLabPage._write_experiment_skeleton(page, _build_record("hyp_chain"), _build_context())
    payload = json.loads(output_path.read_text(encoding="utf-8"))

    assert payload["execution_chain"]["mode"] == "dag"
    assert [node["node_id"] for node in payload["execution_chain"]["nodes"]] == [
        "walk_forward",
        "optimization",
        "stress",
        "governance_evaluation",
    ]

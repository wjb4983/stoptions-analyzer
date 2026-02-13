from __future__ import annotations

from dataclasses import dataclass
from typing import Mapping

import numpy as np


@dataclass(frozen=True)
class ExplainabilityArtifacts:
    top_drivers: list[dict[str, float | str]]
    uncertainty: dict[str, float]
    risk_constraints: dict[str, float]
    decision_trace: list[dict[str, float | str]]
    counterfactual: dict[str, object]
    red_flags: list[dict[str, object]]


def build_trade_explainability(
    *,
    trade_id: str,
    timestamp: str,
    symbol: str,
    requested_size: float,
    sized_trade: float,
    feature_values: Mapping[str, float],
    baseline_values: Mapping[str, float] | None = None,
    feature_uncertainty: Mapping[str, float] | None = None,
    risk_constraints: Mapping[str, float] | None = None,
) -> ExplainabilityArtifacts:
    baseline = {k: float(v) for k, v in (baseline_values or {}).items()}
    uncertainty_map = {k: max(1e-6, float(v)) for k, v in (feature_uncertainty or {}).items()}
    attributions: list[tuple[str, float, float, float, float]] = []
    for name, raw_val in feature_values.items():
        value = float(raw_val)
        base = float(baseline.get(name, 0.0))
        sigma = float(uncertainty_map.get(name, 1.0))
        z = (value - base) / sigma
        contribution = z * float(sized_trade)
        attributions.append((name, contribution, value, base, sigma))

    sorted_attr = sorted(attributions, key=lambda row: abs(float(row[1])), reverse=True)
    top_drivers = [
        {
            "feature": name,
            "direction": "positive" if contrib >= 0.0 else "negative",
            "attribution_score": float(contrib),
            "feature_value": float(value),
            "baseline_value": float(base),
            "feature_sigma": float(sigma),
        }
        for name, contrib, value, base, sigma in sorted_attr[:3]
    ]

    running_score = 0.0
    decision_trace: list[dict[str, float | str]] = []
    for step_idx, (name, contrib, value, _, _) in enumerate(sorted_attr[:5], start=1):
        running_score += float(contrib)
        decision_trace.append(
            {
                "step": str(step_idx),
                "trade_id": trade_id,
                "timestamp": timestamp,
                "symbol": symbol,
                "feature": name,
                "feature_value": float(value),
                "incremental_score": float(contrib),
                "cumulative_score": float(running_score),
            }
        )

    red_flags = detect_explanation_red_flags(
        sized_trade=sized_trade,
        attributions=sorted_attr,
        risk_constraints={k: float(v) for k, v in (risk_constraints or {}).items()},
    )
    uncertainty_payload = {
        "trade_size_confidence": float(np.exp(-abs(float(sized_trade - requested_size)))),
        "attribution_dispersion": float(np.std(np.asarray([row[1] for row in sorted_attr], dtype=float))) if sorted_attr else 0.0,
    }
    constraints_payload = {k: float(v) for k, v in (risk_constraints or {}).items()}
    return ExplainabilityArtifacts(
        top_drivers=top_drivers,
        uncertainty=uncertainty_payload,
        risk_constraints=constraints_payload,
        decision_trace=decision_trace,
        counterfactual=build_counterfactual_explanation(
            trade_id=trade_id,
            symbol=symbol,
            sized_trade=sized_trade,
            top_drivers=top_drivers,
            risk_constraints=constraints_payload,
        ),
        red_flags=red_flags,
    )


def build_counterfactual_explanation(
    *,
    trade_id: str,
    symbol: str,
    sized_trade: float,
    top_drivers: list[dict[str, float | str]],
    risk_constraints: dict[str, float],
) -> dict[str, object]:
    if not top_drivers:
        return {"trade_id": trade_id, "symbol": symbol, "recommended_trade_if_neutral": 0.0, "actions": []}
    target = 0.0
    delta = target - float(sized_trade)
    actions = []
    normalizer = max(1e-6, sum(abs(float(row["attribution_score"])) for row in top_drivers))
    for row in top_drivers:
        direction = -1.0 if str(row["direction"]) == "positive" else 1.0
        weight = abs(float(row["attribution_score"])) / normalizer
        actions.append(
            {
                "feature": str(row["feature"]),
                "required_shift": float(direction * abs(delta) * weight),
                "rationale": "Move contribution toward neutral sizing decision.",
            }
        )
    return {
        "trade_id": trade_id,
        "symbol": symbol,
        "recommended_trade_if_neutral": float(target),
        "current_trade": float(sized_trade),
        "delta_to_neutral": float(delta),
        "risk_constraints_considered": risk_constraints,
        "actions": actions,
    }


def detect_explanation_red_flags(
    *,
    sized_trade: float,
    attributions: list[tuple[str, float, float, float, float]],
    risk_constraints: dict[str, float],
) -> list[dict[str, object]]:
    flags: list[dict[str, object]] = []
    if not attributions:
        return flags
    scores = np.asarray([float(row[1]) for row in attributions], dtype=float)
    dominant = float(np.max(np.abs(scores)))
    total = float(np.sum(np.abs(scores)))
    if total > 0 and dominant / total > 0.85:
        flags.append({"flag": "single_feature_dominance", "severity": "high", "value": dominant / total})

    signed_sum = float(np.sum(scores))
    if sized_trade != 0.0 and signed_sum * float(sized_trade) < 0.0:
        flags.append({"flag": "non_intuitive_sign", "severity": "high", "value": signed_sum})

    constraint_cap = float(risk_constraints.get("max_abs_trade", np.inf))
    if abs(float(sized_trade)) >= constraint_cap > 0:
        flags.append({"flag": "constraint_boundary_hit", "severity": "medium", "value": constraint_cap})
    if np.std(scores) > 3.0 * (abs(np.mean(scores)) + 1e-6):
        flags.append({"flag": "unstable_attribution_dispersion", "severity": "medium", "value": float(np.std(scores))})
    return flags


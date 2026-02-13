from __future__ import annotations

from src.analysis.explainability import (
    build_counterfactual_explanation,
    build_trade_explainability,
    detect_explanation_red_flags,
)


def test_build_trade_explainability_contains_required_sections() -> None:
    artifact = build_trade_explainability(
        trade_id="t1",
        timestamp="2024-01-01T00:00:00",
        symbol="AAA",
        requested_size=0.3,
        sized_trade=0.25,
        feature_values={"momentum": 1.2, "volatility": -0.3, "spread": 0.1},
        baseline_values={"momentum": 0.0, "volatility": 0.0, "spread": 0.0},
        feature_uncertainty={"momentum": 0.4, "volatility": 0.2, "spread": 0.3},
        risk_constraints={"max_abs_trade": 0.25},
    )

    assert len(artifact.top_drivers) > 0
    assert "trade_size_confidence" in artifact.uncertainty
    assert "max_abs_trade" in artifact.risk_constraints
    assert len(artifact.decision_trace) > 0
    assert "actions" in artifact.counterfactual


def test_detect_explanation_red_flags_catches_non_intuitive_sign() -> None:
    flags = detect_explanation_red_flags(
        sized_trade=0.5,
        attributions=[("a", -1.0, 0.0, 0.0, 1.0), ("b", -0.5, 0.0, 0.0, 1.0)],
        risk_constraints={"max_abs_trade": 0.5},
    )
    names = {row["flag"] for row in flags}
    assert "non_intuitive_sign" in names
    assert "constraint_boundary_hit" in names


def test_build_counterfactual_explanation_has_neutral_target() -> None:
    payload = build_counterfactual_explanation(
        trade_id="t2",
        symbol="BBB",
        sized_trade=-0.7,
        top_drivers=[
            {"feature": "momentum", "direction": "negative", "attribution_score": -0.8},
            {"feature": "vol", "direction": "positive", "attribution_score": 0.2},
        ],
        risk_constraints={"max_abs_trade": 1.0},
    )
    assert payload["recommended_trade_if_neutral"] == 0.0
    assert len(payload["actions"]) == 2

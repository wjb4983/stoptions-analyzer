from __future__ import annotations

from models.deployment import (
    ModelSlots,
    PromotionGates,
    REASON_PROMOTION_GATES_PASSED,
    REASON_RISK_BREACH,
    REASON_ROLLBACK_TO_PRIOR_CHAMPION,
    REASON_SHADOW_UNDERPERFORMANCE,
)


def test_shadow_mode_promotion_when_all_gates_pass() -> None:
    slots = ModelSlots(champion="model_v1", challenger="model_v2")
    gates = PromotionGates(
        min_risk_adjusted_return_delta=0.05,
        max_drawdown=-0.15,
        max_turnover=3.0,
        min_stability_score=0.6,
    )

    decision = slots.evaluate_shadow_mode(
        champion_metrics={"risk_adjusted_return": 1.10},
        challenger_metrics={
            "risk_adjusted_return": 1.25,
            "max_drawdown": -0.09,
            "turnover_total": 1.2,
            "stability_score": 0.73,
        },
        gates=gates,
    )

    assert decision["promoted"] is True
    assert decision["reason_code"] == REASON_PROMOTION_GATES_PASSED
    assert slots.champion == "model_v2"
    assert slots.challenger == "model_v1"
    assert slots.audit_log[-1].reason_code == REASON_PROMOTION_GATES_PASSED


def test_shadow_mode_auto_rollback_on_risk_breach() -> None:
    slots = ModelSlots(champion="model_v1", challenger="model_v2")
    gates = PromotionGates(max_drawdown=-0.12, max_turnover=2.0)

    decision = slots.evaluate_shadow_mode(
        champion_metrics={"risk_adjusted_return": 0.8},
        challenger_metrics={
            "risk_adjusted_return": 1.5,
            "max_drawdown": -0.20,
            "turnover_total": 0.5,
            "stability_score": 0.9,
        },
        gates=gates,
    )

    assert decision["rolled_back"] is True
    assert decision["reason_code"] == REASON_RISK_BREACH
    assert slots.challenger is None
    assert slots.candidate == "model_v2"
    assert slots.audit_log[-1].reason_code == REASON_RISK_BREACH


def test_shadow_mode_auto_rollback_on_underperformance() -> None:
    slots = ModelSlots(champion="model_v1", challenger="model_v2")
    gates = PromotionGates(min_risk_adjusted_return_delta=0.10, min_stability_score=0.5)

    decision = slots.evaluate_shadow_mode(
        champion_metrics={"risk_adjusted_return": 1.2},
        challenger_metrics={
            "risk_adjusted_return": 1.25,
            "max_drawdown": -0.05,
            "turnover_total": 0.7,
            "stability_score": 0.8,
        },
        gates=gates,
    )

    assert decision["rolled_back"] is True
    assert decision["reason_code"] == REASON_SHADOW_UNDERPERFORMANCE
    assert slots.candidate == "model_v2"
    assert slots.audit_log[-1].reason_code == REASON_SHADOW_UNDERPERFORMANCE


def test_post_promotion_rollback_to_prior_champion() -> None:
    slots = ModelSlots(champion="model_v1", challenger="model_v2")
    gates = PromotionGates(min_risk_adjusted_return_delta=0.0, max_drawdown=-0.15, max_turnover=3.0, min_stability_score=0.55)

    promote = slots.evaluate_shadow_mode(
        champion_metrics={"risk_adjusted_return": 1.0},
        challenger_metrics={
            "risk_adjusted_return": 1.2,
            "max_drawdown": -0.10,
            "turnover_total": 1.0,
            "stability_score": 0.7,
        },
        gates=gates,
    )
    assert promote["promoted"] is True

    rolled_back = slots.monitor_challenger_and_rollback(
        champion_metrics={"risk_adjusted_return": 1.1},
        challenger_metrics={"risk_adjusted_return": 0.8, "max_drawdown": -0.18, "turnover_total": 3.5},
        gates=gates,
    )

    assert rolled_back is True
    assert slots.champion == "model_v1"
    assert slots.challenger == "model_v2"
    assert slots.audit_log[-1].reason_code == REASON_ROLLBACK_TO_PRIOR_CHAMPION

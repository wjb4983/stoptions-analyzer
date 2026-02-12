from __future__ import annotations

import numpy as np

from src.backtesting.strategies import GreekNeutralTargetRequest, build_greek_neutral_targets
from src.backtesting.vectorized import GreekLimitConfig, backtest_vectorized


def test_expiry_day_cash_settlement_and_roll_event() -> None:
    prices = np.array([[100.0], [103.0], [105.0], [107.0]])
    signals = np.array([[0.0], [1.0], [1.0], [0.0]])
    dates = np.array(["2024-01-10", "2024-01-11", "2024-01-12", "2024-01-13"], dtype=object)

    result = backtest_vectorized(
        prices,
        signals,
        dates=dates,
        contract_expiry_by_asset=["2024-01-12"],
        contract_strike_by_asset=[102.0],
        contract_option_type_by_asset=["call"],
        contract_multipliers_by_asset=[100.0],
        contract_settlement_style_by_asset=["cash"],
        enable_contract_roll=True,
        roll_days_before_expiry=1,
    )

    events = result.cost_breakdown["derivatives_events"]
    assert any(evt["event"] == "expiry_cash_settlement" for evt in events)
    assert any(evt["event"] == "contract_roll" for evt in events)
    assert result.cost_breakdown["lifecycle"][2] > 0.0


def test_assignment_exercise_resets_contract_position() -> None:
    prices = np.array([[100.0], [101.0], [99.0], [98.0]])
    signals = np.array([[0.0], [1.0], [1.0], [1.0]])
    dates = np.array(["2024-01-10", "2024-01-11", "2024-01-12", "2024-01-13"], dtype=object)

    result = backtest_vectorized(
        prices,
        signals,
        dates=dates,
        contract_expiry_by_asset=["2024-01-12"],
        contract_strike_by_asset=[100.0],
        contract_option_type_by_asset=["put"],
        contract_settlement_style_by_asset=["physical"],
    )

    assert result.positions[2] == 0.0
    assert any(evt["event"] == "assignment_exercise" for evt in result.cost_breakdown["derivatives_events"])


def test_greek_neutral_api_and_limits() -> None:
    base = np.array([[0.6, -0.4], [0.5, -0.5], [0.5, -0.5]])
    delta = np.array([[1.0, 1.0], [0.8, 1.2], [0.8, 1.2]])
    neutral = build_greek_neutral_targets(
        GreekNeutralTargetRequest(base_targets=base, delta=delta, neutralize=("delta",))
    )
    assert neutral.shape == base.shape
    assert np.allclose(np.sum(neutral * delta, axis=1), 0.0, atol=1e-8)

    prices = np.array([[100.0, 100.0], [101.0, 99.0], [102.0, 98.0]])
    result = backtest_vectorized(
        prices,
        neutral,
        greek_sensitivities={"delta": np.array([[1.0, 1.0], [1.0, 1.0], [1.0, 1.0]])},
        greek_limit_config=GreekLimitConfig(max_abs_delta=0.1),
    )
    assert "greek_diagnostics" in result.cost_breakdown

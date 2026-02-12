from __future__ import annotations

import numpy as np

from src.backtesting.execution import BpsSlippage, FixedCommission, ShortBorrowCost
from src.backtesting.signals.config import NoExitConfig, parse_entry_signal_config
from src.backtesting.signals.engine import build_standardized_targets
from src.backtesting.strategies.ensemble import weighted_voting
from src.backtesting.vectorized import backtest_vectorized


def test_vectorized_positions_shift_signals_forward() -> None:
    prices = np.array([100.0, 101.0, 102.0, 103.0])
    signals = np.array([0.0, 1.0, -1.0, 0.5])

    result = backtest_vectorized(prices, signals)

    expected_positions = np.array([0.0, 0.0, 1.0, -1.0])
    assert np.allclose(result.positions, expected_positions)


def test_vectorized_multi_asset_aggregates_returns_and_costs() -> None:
    prices = np.array(
        [
            [100.0, 50.0],
            [101.0, 49.0],
            [102.0, 48.0],
        ]
    )
    signals = np.array(
        [
            [1.0, -1.0],
            [1.0, -1.0],
            [0.0, 0.0],
        ]
    )

    result = backtest_vectorized(
        prices,
        signals,
        slippage_model=BpsSlippage(10.0),
        fee_model=FixedCommission(0.001),
        borrow_cost_model=ShortBorrowCost(0.10, periods_per_year=252.0),
        weights=np.array([0.6, 0.4]),
    )

    assert result.positions.shape == (3, 2)
    assert result.trades.shape == (3, 2)
    assert result.turnover.shape == (3,)
    assert result.daily_returns.shape == (3,)

    # Signal-to-position shift means costs start on the second bar.
    assert result.cost_breakdown["total"][1] > 0
    # Held short exposure incurs borrow drag.
    assert result.cost_breakdown["borrow"][1] > 0
    assert np.isclose(
        result.cost_breakdown["totals"]["total"],
        np.sum(result.cost_breakdown["total"]),
    )


def test_standardized_signal_generation_is_no_lookahead() -> None:
    base_prices = np.array([[100.0], [101.0], [102.0], [101.0], [100.0], [99.0]])
    shocked_prices = base_prices.copy()
    shocked_prices[-1, 0] = 130.0
    missing = np.zeros_like(base_prices, dtype=bool)

    entry_cfg = parse_entry_signal_config("ma_trend", {"ma_window": 3}, default_lookback_days=10, default_skip_days=1)

    base = build_standardized_targets(
        close_prices=base_prices,
        missing_mask=missing,
        entry_config=entry_cfg,
        exit_config=NoExitConfig(),
    )
    shocked = build_standardized_targets(
        close_prices=shocked_prices,
        missing_mask=missing,
        entry_config=entry_cfg,
        exit_config=NoExitConfig(),
    )

    assert np.allclose(base.values[:-1], shocked.values[:-1])
    assert np.allclose(base.confidence[:-1], shocked.confidence[:-1])
    assert np.array_equal(base.horizon_bars[:-1], shocked.horizon_bars[:-1])


def test_ensemble_combiner_is_point_in_time() -> None:
    signals = np.array(
        [
            [[1.0, -1.0]],
            [[-1.0, -1.0]],
            [[1.0, 1.0]],
        ]
    )
    combined = weighted_voting(signals, np.array([0.7, 0.3]))
    changed = signals.copy()
    changed[-1, 0, :] = np.array([-1.0, -1.0])
    combined_changed = weighted_voting(changed, np.array([0.7, 0.3]))

    assert np.allclose(combined[:-1], combined_changed[:-1])

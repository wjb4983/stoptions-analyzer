from __future__ import annotations

import numpy as np

from src.backtesting.execution import BpsSlippage, FixedCommission, ShortBorrowCost
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

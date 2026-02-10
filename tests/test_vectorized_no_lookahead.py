from __future__ import annotations

import numpy as np

from src.backtesting.vectorized import backtest_vectorized


def test_vectorized_positions_shift_signals_forward() -> None:
    prices = np.array([100.0, 101.0, 102.0, 103.0])
    signals = np.array([0.0, 1.0, -1.0, 0.5])

    result = backtest_vectorized(prices, signals)

    expected_positions = np.array([0.0, 0.0, 1.0, -1.0])
    assert np.allclose(result.positions, expected_positions)

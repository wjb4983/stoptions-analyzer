from __future__ import annotations

import numpy as np

from src.backtesting.execution import PartialFillModel
from src.backtesting.vectorized import backtest_vectorized


def test_partial_fill_model_carries_residual_orders_forward() -> None:
    model = PartialFillModel(max_participation_per_bar=0.5)
    requested = np.array([[0.8], [0.0], [0.0]], dtype=float)
    available = np.ones_like(requested)

    filled, residual, fills = model.run(requested, available, latency_bars=0)

    np.testing.assert_allclose(filled[:, 0], np.array([0.5, 0.3, 0.0]))
    np.testing.assert_allclose(residual[:, 0], np.array([0.3, 0.0, 0.0]))
    assert len(fills) == requested.shape[0] * requested.shape[1]


def test_partial_fill_model_respects_latency_bars() -> None:
    model = PartialFillModel(max_participation_per_bar=1.0)
    requested = np.array([[1.0], [0.0], [0.0]], dtype=float)
    available = np.ones_like(requested)

    filled, residual, _fills = model.run(requested, available, latency_bars=1)

    np.testing.assert_allclose(filled[:, 0], np.array([0.0, 1.0, 0.0]))
    np.testing.assert_allclose(residual[:, 0], np.array([0.0, 0.0, 0.0]))


def test_vectorized_backtest_exposes_fill_artifacts() -> None:
    prices = np.array([100.0, 101.0, 102.0, 103.0], dtype=float)
    signals = np.array([0.0, 0.8, 0.8, 0.8], dtype=float)
    result = backtest_vectorized(
        prices,
        signals,
        volumes=np.ones_like(prices),
        available_bar_volume=np.ones_like(prices),
        max_participation_per_bar=0.25,
        latency_bars=1,
    )

    assert result.fills
    total_filled = sum(float(row["filled_size"]) for row in result.fills)
    assert total_filled > 0
    assert any(float(row["residual_size"]) > 0 for row in result.fills)

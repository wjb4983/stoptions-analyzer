from __future__ import annotations

import numpy as np

from src.backtesting.event_driven import EventDrivenBacktester, Order
from src.backtesting.execution import LatencyQueueDriftSlippage, PartialFillModel
from src.backtesting.vectorized import backtest_vectorized, replay_from_event_logs


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


def test_order_lifecycle_persists_across_bars_until_filled() -> None:
    prices = np.array([100.0, 100.0, 100.0, 100.0, 100.0, 100.0], dtype=float)
    signals = np.array([0.0, 1.0, 1.0, 1.0, 1.0, 1.0], dtype=float)
    result = backtest_vectorized(
        prices,
        signals,
        available_bar_volume=np.array([1.0, 1.0, 1.0, 1.0, 1.0, 1.0]),
        max_participation_per_bar=0.25,
        time_in_force="gtc",
        urgency="high",
    )

    fills = [row for row in result.execution_events if row["event_type"] == "fill"]
    assert len(fills) >= 2
    assert any(row["state"] == "partial" for row in fills)
    assert fills[-1]["state"] in {"partial", "filled"}
    assert all(row["time_in_force"] == "gtc" for row in result.execution_events)
    assert all(row["urgency"] == "high" for row in result.execution_events)


def test_order_lifecycle_cancel_replace_and_replay_determinism() -> None:
    prices = np.array([100.0, 100.0, 100.0, 100.0, 100.0], dtype=float)
    signals = np.array([0.0, 1.0, -1.0, -1.0, -1.0], dtype=float)
    result = backtest_vectorized(
        prices,
        signals,
        available_bar_volume=np.array([1.0, 1.0, 1.0, 1.0, 1.0]),
        max_participation_per_bar=0.25,
    )

    event_types = [row["event_type"] for row in result.execution_events]
    assert "cancel" in event_types
    assert "submit" in event_types
    replayed = replay_from_event_logs(result.execution_events)
    assert replayed == replay_from_event_logs(list(reversed(result.execution_events)))


def test_vectorized_latency_drift_slippage_uses_latency_context() -> None:
    prices = np.array([100.0, 100.0, 100.0, 100.0, 100.0, 100.0], dtype=float)
    signals = np.array([0.0, 1.0, 1.0, 1.0, 1.0, 1.0], dtype=float)
    model = LatencyQueueDriftSlippage(drift_bps_per_bar=4.0, queue_drift_bps=2.0)

    fast = backtest_vectorized(
        prices,
        signals,
        slippage_model=model,
        available_bar_volume=np.ones_like(prices),
        max_participation_per_bar=1.0,
        latency_bars=0,
        latency_ms=0,
    )
    slow = backtest_vectorized(
        prices,
        signals,
        slippage_model=model,
        available_bar_volume=np.ones_like(prices),
        max_participation_per_bar=1.0,
        latency_bars=2,
        latency_ms=0,
    )

    assert slow.cost_breakdown["totals"]["slippage"] > fast.cost_breakdown["totals"]["slippage"]


def test_event_driven_backtester_accepts_latency_aware_slippage_context() -> None:
    class OneShot:
        def __init__(self) -> None:
            self.sent = False

        def on_bar(self, bar, portfolio):
            if self.sent:
                return []
            self.sent = True
            return [Order(symbol="A", quantity=1.0, side="buy")]

    model = LatencyQueueDriftSlippage(drift_bps_per_bar=2.0, queue_drift_bps=3.0)
    bars = [
        {"open": 100.0, "timestamp": "t0"},
        {"open": 100.0, "timestamp": "t1", "latency_bars": 2, "queue_rank_proxy": 1.0, "latency_ms": 0},
    ]
    result = EventDrivenBacktester(initial_cash=1_000.0, slippage_model=model).run(bars, OneShot())
    assert result.fills
    assert result.fills[0].price > 100.0

from __future__ import annotations

import numpy as np

from src.backtesting.event_driven import EventDrivenBacktester, Order
from src.backtesting.vectorized import backtest_vectorized
from tests.fixtures_datasets import (
    deterministic_event_bars,
    deterministic_event_target_positions,
    deterministic_flat_prices_case,
    deterministic_missing_bar_case,
    deterministic_multi_asset_case,
    deterministic_single_asset_case,
    deterministic_zero_volume_case,
)


class TargetPositionStrategy:
    """Emit orders so position equals target signal after next-open fill."""

    def __init__(self, symbol: str, target_positions: np.ndarray) -> None:
        self.symbol = symbol
        self.target_positions = target_positions
        self.bar_idx = 0

    def on_bar(self, bar: dict[str, float | str], portfolio) -> list[Order]:
        target = float(self.target_positions[self.bar_idx])
        current = portfolio.positions.get(self.symbol).quantity if self.symbol in portfolio.positions else 0.0
        delta = target - current
        self.bar_idx += 1
        if delta > 0:
            return [Order(symbol=self.symbol, quantity=delta, side="buy", timestamp=str(bar.get("timestamp")))]
        if delta < 0:
            return [Order(symbol=self.symbol, quantity=abs(delta), side="sell", timestamp=str(bar.get("timestamp")))]
        return []


def test_signal_timing_has_no_lookahead_leakage() -> None:
    prices = np.array([100.0, 50.0, 200.0], dtype=float)
    signals = np.array([0.0, 1.0, 1.0], dtype=float)

    result = backtest_vectorized(prices, signals, initial_equity=1.0)

    # If there were leakage, the big jump on bar 2 would be captured immediately.
    assert np.isclose(result.returns[1], 0.0)
    assert np.isclose(result.positions[1], 0.0)
    assert np.isclose(result.positions[2], 1.0)


def test_trade_and_position_consistency_invariants() -> None:
    prices, signals, weights = deterministic_multi_asset_case()
    result = backtest_vectorized(prices, signals, weights=weights)

    expected_positions = np.zeros_like(signals)
    expected_positions[1:] = signals[:-1]
    expected_trades = expected_positions.copy()
    expected_trades[1:] = expected_positions[1:] - expected_positions[:-1]

    assert np.allclose(result.positions, expected_positions)
    assert np.allclose(result.trades, expected_trades)
    assert np.allclose(result.turnover, np.sum(np.abs(expected_trades) * weights, axis=1))


def test_pnl_reconciles_to_equity_and_returns() -> None:
    prices, signals = deterministic_single_asset_case()
    initial_equity = 1000.0

    result = backtest_vectorized(prices, signals, initial_equity=initial_equity)

    pnl_from_returns = np.zeros_like(result.pnl)
    pnl_from_returns[1:] = result.equity_curve[:-1] * result.returns[1:]

    assert np.allclose(result.pnl, pnl_from_returns)
    assert np.allclose(result.equity_curve[1:] - result.equity_curve[:-1], result.pnl[1:])


def test_missing_bars_preserve_index_and_compute() -> None:
    prices, signals, timestamps = deterministic_missing_bar_case()

    result = backtest_vectorized(prices, signals)

    # Gap in timestamps does not break calculations when bars are missing.
    assert timestamps[2] == "2024-01-05"
    assert np.isfinite(result.returns).all()
    assert result.positions.shape == prices.shape


def test_flat_prices_produce_zero_returns_despite_trading() -> None:
    prices, signals = deterministic_flat_prices_case()

    result = backtest_vectorized(prices, signals)

    assert np.allclose(result.returns, 0.0)
    assert np.allclose(result.pnl, 0.0)
    assert np.allclose(result.equity_curve, result.equity_curve[0])


def test_zero_volume_symbol_stays_untraded() -> None:
    prices, signals = deterministic_zero_volume_case()

    result = backtest_vectorized(prices, signals)

    # Illiquid second asset remains untouched.
    assert np.allclose(result.positions[:, 1], 0.0)
    assert np.allclose(result.trades[:, 1], 0.0)


def test_tiny_universe_single_symbol_case() -> None:
    prices = np.array([100.0, 101.0, 102.0], dtype=float)
    signals = np.array([1.0, 1.0, 0.0], dtype=float)

    result = backtest_vectorized(prices, signals)

    assert result.positions.shape == (3,)
    assert result.trades.shape == (3,)
    assert result.returns.shape == (3,)


def test_vectorized_and_event_driven_agree_on_deterministic_path() -> None:
    bars = deterministic_event_bars()
    targets = deterministic_event_target_positions()
    prices = np.array([float(bar["open"]) for bar in bars], dtype=float)

    vec = backtest_vectorized(prices=prices, signals=targets, initial_equity=1_000.0)

    strategy = TargetPositionStrategy(symbol="ABC", target_positions=targets)
    ev = EventDrivenBacktester(initial_cash=1_000.0)
    event_result = ev.run(bars, strategy)

    final_qty = event_result.portfolio.positions.get("ABC").quantity if "ABC" in event_result.portfolio.positions else 0.0
    event_final_equity = event_result.portfolio.cash + final_qty * prices[-1]

    assert np.isclose(event_final_equity, vec.equity_curve[-1])

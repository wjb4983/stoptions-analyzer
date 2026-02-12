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


def test_compute_metrics_include_rich_period_aware_fields() -> None:
    prices = np.array([100.0, 101.0, 102.0, 99.0, 100.0, 104.0, 103.0], dtype=float)
    signals = np.array([1.0, 1.0, 1.0, 0.0, 1.0, 1.0, 1.0], dtype=float)

    result_daily = backtest_vectorized(prices, signals, timeframe="1d")
    result_minute = backtest_vectorized(prices, signals, timeframe="1m")

    required = {
        "cagr",
        "max_drawdown",
        "calmar",
        "sortino",
        "downside_deviation",
        "skew",
        "kurtosis",
        "hit_rate",
        "profit_factor",
        "exposure_time",
        "turnover_adjusted_return",
        "rolling_sharpe_mean",
        "rolling_drawdown_worst",
    }
    assert required.issubset(result_daily.metrics.keys())
    assert float(result_daily.metrics["periods_per_year"]) == 252.0
    assert float(result_minute.metrics["periods_per_year"]) == 252.0 * 390.0
    assert 0.0 <= float(result_daily.metrics["hit_rate"]) <= 1.0
    assert 0.0 <= float(result_daily.metrics["exposure_time"]) <= 1.0
    assert np.isfinite(float(result_daily.metrics["turnover_adjusted_return"]))


def test_split_transform_and_dividend_cashflow_keep_value_continuity() -> None:
    prices = np.array([100.0, 50.0, 51.0, 52.0], dtype=float)
    signals = np.array([1.0, 1.0, 1.0, 1.0], dtype=float)
    splits = np.array([1.0, 2.0, 1.0, 1.0], dtype=float)
    dividends = np.array([0.0, 0.0, 1.0, 0.0], dtype=float)

    result = backtest_vectorized(
        prices=prices,
        signals=signals,
        initial_equity=1000.0,
        corporate_action_splits=splits,
        corporate_action_dividends=dividends,
    )

    # 2-for-1 split should scale position units going forward.
    assert np.isclose(result.positions[1], 2.0)
    # Dividend cashflow should contribute positive return on dividend bar.
    assert float(result.cost_breakdown["dividend_return"][2]) > 0.0


def test_split_pnl_is_neutral_without_dividends() -> None:
    prices = np.array([100.0, 100.0, 100.0, 100.0], dtype=float)
    signals = np.array([1.0, 1.0, 1.0, 1.0], dtype=float)
    splits = np.array([1.0, 2.0, 1.0, 1.0], dtype=float)

    split_result = backtest_vectorized(
        prices=prices,
        signals=signals,
        initial_equity=1000.0,
        corporate_action_splits=splits,
    )
    baseline_result = backtest_vectorized(
        prices=prices,
        signals=signals,
        initial_equity=1000.0,
    )

    assert np.isclose(float(split_result.equity_curve[-1]), float(baseline_result.equity_curve[-1]))
    assert np.isclose(float(split_result.positions[1]), 2.0)


def test_margin_constraints_force_deleveraging_and_keep_utilization_bounded() -> None:
    prices = np.array(
        [
            [100.0, 100.0],
            [100.0, 100.0],
            [100.0, 100.0],
            [100.0, 100.0],
        ],
        dtype=float,
    )
    # Deliberately over-sized to trigger forced deleveraging under tight schedules.
    signals = np.array(
        [
            [0.0, 0.0],
            [4.0, -4.0],
            [4.0, -4.0],
            [2.0, -2.0],
        ],
        dtype=float,
    )

    result = backtest_vectorized(
        prices=prices,
        signals=signals,
        margin_schedule_by_asset=np.array([0.9, 0.9], dtype=float),
        stress_addon_by_asset=np.array([0.3, 0.3], dtype=float),
        concentration_addon=0.2,
        hard_to_borrow_flags=np.array([0.0, 1.0], dtype=float),
        hard_to_borrow_addon=0.2,
    )

    account_state = result.cost_breakdown["account_state"]
    margin_util = np.asarray(account_state["margin_utilization"], dtype=float)
    forced = np.asarray(account_state["forced_liquidation"], dtype=float)
    scale = np.asarray(account_state["deleveraging_scale"], dtype=float)

    assert np.all(margin_util <= 1.0 + 1e-9)
    assert np.any(forced > 0.0)
    assert np.any(scale < 1.0)


def test_hard_to_borrow_short_block_prevents_impossible_short_states() -> None:
    prices = np.array([[100.0, 100.0], [100.0, 100.0], [100.0, 100.0]], dtype=float)
    signals = np.array([[0.0, 0.0], [1.0, -1.0], [1.0, -1.0]], dtype=float)

    result = backtest_vectorized(
        prices=prices,
        signals=signals,
        hard_to_borrow_flags=np.array([0.0, 1.0], dtype=float),
        hard_to_borrow_short_block=True,
    )

    positions = np.asarray(result.positions, dtype=float)
    # Asset 1 is flagged HTB, so no short position should survive post-constraints.
    assert np.all(positions[:, 1] >= -1e-12)

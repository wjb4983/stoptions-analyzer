from __future__ import annotations

import numpy as np
import pytest

from src.backtesting.signals.config import (
    MaxHoldExitConfig,
    MomentumFlipExitConfig,
    NoExitConfig,
    TimeSeriesMomentumEntryConfig,
    TrailingStopExitConfig,
    parse_entry_signal_config,
    parse_exit_signal_config,
)
from src.backtesting.signals.engine import build_targets
from src.backtesting.signals.entry import BreakoutEntry, MovingAverageTrendEntry, TimeSeriesMomentumEntry
from src.backtesting.signals.exit import MaxHoldExit, MomentumFlipExit, PositionState, TrailingStopExit


def test_ts_momentum_entry_signal_positive_and_negative() -> None:
    signal = TimeSeriesMomentumEntry(TimeSeriesMomentumEntryConfig(lookback_days=2, skip_days=0))
    prices = np.array([100.0, 101.0, 103.0, 102.0])
    missing = np.zeros_like(prices, dtype=bool)

    assert signal.value_at(2, prices, missing) == 1
    assert signal.value_at(3, prices, missing) == 1

    prices2 = np.array([103.0, 101.0, 99.0, 98.0])
    assert signal.value_at(2, prices2, missing) == -1


def test_ma_trend_entry_signal() -> None:
    signal = MovingAverageTrendEntry(config=parse_entry_signal_config("ma_trend", {"ma_window": 3}, default_lookback_days=90, default_skip_days=5))
    prices = np.array([100.0, 100.0, 101.0, 99.0])
    missing = np.zeros_like(prices, dtype=bool)

    assert signal.value_at(2, prices, missing) == 1
    assert signal.value_at(3, prices, missing) == 0


def test_breakout_entry_signal() -> None:
    signal = BreakoutEntry(config=parse_entry_signal_config("breakout", {"breakout_window": 3}, default_lookback_days=90, default_skip_days=5))
    prices = np.array([10.0, 11.0, 12.0, 13.0, 12.0])
    missing = np.zeros_like(prices, dtype=bool)

    assert signal.value_at(3, prices, missing) == 1
    assert signal.value_at(4, prices, missing) == 0


def test_momentum_flip_exit_signal() -> None:
    exit_signal = MomentumFlipExit(MomentumFlipExitConfig(lookback_days=2, skip_days=0))
    prices = np.array([100.0, 101.0, 99.0, 98.0])
    missing = np.zeros_like(prices, dtype=bool)
    long_pos = PositionState(side=1, entry_price=100.0, peak_price=101.0, bars_held=1)

    assert bool(exit_signal.should_exit(2, prices, missing, long_pos)) is True


def test_trailing_stop_exit_signal() -> None:
    exit_signal = TrailingStopExit(TrailingStopExitConfig(trailing_stop_pct=0.1))
    prices = np.array([100.0, 110.0, 98.0])
    missing = np.zeros_like(prices, dtype=bool)
    long_pos = PositionState(side=1, entry_price=100.0, peak_price=110.0, bars_held=2)

    assert exit_signal.should_exit(2, prices, missing, long_pos) is True


def test_max_hold_exit_signal() -> None:
    exit_signal = MaxHoldExit(MaxHoldExitConfig(max_hold_bars=3))
    prices = np.array([100.0, 101.0, 102.0, 103.0])
    missing = np.zeros_like(prices, dtype=bool)
    position = PositionState(side=1, entry_price=100.0, peak_price=103.0, bars_held=3)

    assert exit_signal.should_exit(3, prices, missing, position) is True


def test_build_targets_with_selected_signals_and_exit_wiring() -> None:
    prices = np.array([[100.0], [102.0], [104.0], [99.0], [98.0], [97.0]])
    missing = np.zeros_like(prices, dtype=bool)

    entry_cfg = parse_entry_signal_config(
        "ts_momentum",
        {"lookback_days": 2, "skip_days": 0},
        default_lookback_days=90,
        default_skip_days=5,
    )
    exit_cfg = parse_exit_signal_config(
        "momentum_flip",
        {"lookback_days": 1, "skip_days": 0},
        default_lookback_days=90,
        default_skip_days=5,
    )
    signals = build_targets(
        close_prices=prices,
        missing_mask=missing,
        entry_config=entry_cfg,
        exit_config=exit_cfg,
    )

    assert signals.shape == prices.shape
    assert signals[2, 0] == 1.0
    assert signals[3, 0] == -1.0
    assert signals[-1, 0] == -1.0


def test_config_parsing_and_validation() -> None:
    entry = parse_entry_signal_config(
        "ts_momentum",
        None,
        default_lookback_days=21,
        default_skip_days=2,
    )
    assert isinstance(entry, TimeSeriesMomentumEntryConfig)
    assert entry.lookback_days == 21

    exit_cfg = parse_exit_signal_config(
        "none",
        None,
        default_lookback_days=21,
        default_skip_days=2,
    )
    assert isinstance(exit_cfg, NoExitConfig)

    with pytest.raises(ValueError):
        parse_entry_signal_config("breakout", {"breakout_window": 0}, default_lookback_days=21, default_skip_days=2)

    with pytest.raises(ValueError):
        parse_exit_signal_config("trailing_stop", {"trailing_stop_pct": 1.5}, default_lookback_days=21, default_skip_days=2)

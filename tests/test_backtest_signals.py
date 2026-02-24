from __future__ import annotations

import numpy as np
import pytest

from src.backtesting.signals.config import (
    MaxHoldExitConfig,
    MeanReversionEntryConfig,
    MomentumFlipExitConfig,
    NoExitConfig,
    SeasonalityEventEntryConfig,
    TimeSeriesMomentumEntryConfig,
    TrendStrengthRegimeEntryConfig,
    TrailingStopExitConfig,
    VRPHarvestEntryConfig,
    VolatilityCarryEntryConfig,
    parse_entry_signal_config,
    parse_execution_model_config,
    parse_exit_signal_config,
    parse_strategy_knobs,
)
from src.backtesting.signals.engine import build_standardized_targets, build_targets
from src.backtesting.signals.entry import (
    BreakoutEntry,
    MeanReversionEntry,
    MovingAverageTrendEntry,
    SeasonalityEventEntry,
    TimeSeriesMomentumEntry,
    TrendStrengthRegimeEntry,
    VRPHarvestEntry,
    VolatilityCarryEntry,
)
from src.backtesting.signals.exit import MaxHoldExit, MomentumFlipExit, PositionState, TrailingStopExit
from src.backtesting.strategies.ensemble import meta_model_weighting, risk_budgeted_blend, weighted_voting


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


def test_new_signal_families_generate_expected_sides() -> None:
    missing = np.zeros(8, dtype=bool)

    mean_rev = MeanReversionEntry(MeanReversionEntryConfig(lookback_days=5, zscore_threshold=0.8))
    prices_mean = np.array([100.0, 100.0, 100.0, 100.0, 108.0, 95.0, 94.0, 102.0])
    assert mean_rev.value_at(4, prices_mean, missing) == -1
    assert mean_rev.value_at(6, prices_mean, missing) == 1

    vol_carry = VolatilityCarryEntry(VolatilityCarryEntryConfig(short_vol_window=2, long_vol_window=4, min_carry_spread=0.0))
    prices_vol = np.array([100.0, 103.0, 97.0, 102.0, 101.0, 101.5, 101.0, 101.4])
    assert vol_carry.value_at(4, prices_vol, missing) == 1

    trend_strength = TrendStrengthRegimeEntry(TrendStrengthRegimeEntryConfig(trend_window=4, strength_window=4, min_strength=0.7))
    prices_trend = np.array([100.0, 101.0, 102.0, 103.0, 104.0, 105.0, 104.0, 106.0])
    assert trend_strength.value_at(4, prices_trend, missing) == 1

    seasonality = SeasonalityEventEntry(SeasonalityEventEntryConfig(seasonal_period=3, event_offset=0, event_window=1))
    prices_seasonal = np.array([100.0, 101.0, 99.0, 104.0, 102.0, 100.0, 106.0, 103.0])
    assert seasonality.value_at(3, prices_seasonal, missing) == 1
    assert seasonality.value_at(4, prices_seasonal, missing) == 0



def test_vrp_harvest_threshold_behavior_and_long_only() -> None:
    prices = np.array([0.50, 0.50, 0.50, 0.51, 0.52, 0.53])
    missing = np.zeros_like(prices, dtype=bool)

    signal = VRPHarvestEntry(VRPHarvestEntryConfig(realized_vol_lookback=3, vrp_threshold=0.01))
    assert signal.value_at(3, prices, missing) == -1

    low_iv = np.array([0.03, 0.02, 0.025, 0.015, 0.02, 0.018])
    assert signal.value_at(4, low_iv, missing) == 1

    long_only_signal = VRPHarvestEntry(VRPHarvestEntryConfig(realized_vol_lookback=3, vrp_threshold=0.01, long_only=True))
    assert long_only_signal.value_at(3, prices, missing) == 0


def test_vrp_harvest_nan_and_missing_safety() -> None:
    prices = np.array([0.10, 0.11, np.nan, 0.12, 0.13])
    missing = np.zeros_like(prices, dtype=bool)
    signal = VRPHarvestEntry(VRPHarvestEntryConfig(realized_vol_lookback=3, vrp_threshold=0.0))

    assert signal.value_at(3, prices, missing) == 0

    prices2 = np.array([0.10, 0.11, 0.09, 0.12, 0.13])
    missing2 = np.zeros_like(prices2, dtype=bool)
    missing2[2] = True
    assert signal.value_at(4, prices2, missing2) == 0

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

    vrp_entry = parse_entry_signal_config(
        "vrp_harvest",
        {"iv_feature_name": "iv_1m", "realized_vol_lookback": 15, "vrp_threshold": 0.02, "regime_filter": True},
        default_lookback_days=21,
        default_skip_days=2,
    )
    assert isinstance(vrp_entry, VRPHarvestEntryConfig)
    assert vrp_entry.realized_vol_lookback == 15

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


def test_standardized_output_contains_confidence_and_horizon() -> None:
    prices = np.array([[100.0], [101.0], [102.0], [103.0], [104.0]])
    missing = np.zeros_like(prices, dtype=bool)
    output = build_standardized_targets(
        close_prices=prices,
        missing_mask=missing,
        entry_config=parse_entry_signal_config("ma_trend", {"ma_window": 3}, default_lookback_days=10, default_skip_days=1),
        exit_config=NoExitConfig(),
    )

    assert output.values.shape == prices.shape
    assert output.confidence.shape == prices.shape
    assert output.horizon_bars.shape == prices.shape
    assert np.all(output.confidence >= 0.0)
    assert np.all(output.horizon_bars >= 0)


def test_ensemble_combiner_variants() -> None:
    signals = np.array(
        [
            [[1.0, 0.0, -1.0]],
            [[1.0, 1.0, -1.0]],
        ]
    )
    voted = weighted_voting(signals, np.array([0.5, 0.3, 0.2]))
    risked = risk_budgeted_blend(signals, np.array([0.4, 0.4, 0.2]), np.array([0.2, 0.3, 0.5]))
    meta = meta_model_weighting(signals, np.array([0.4, 0.2, 0.1]), bias=0.1)

    assert voted.shape == (2, 1)
    assert risked.shape == (2, 1)
    assert meta.shape == (2, 1)
    assert np.all(np.abs(meta) <= 1.0)


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

    knobs = parse_strategy_knobs("mean_reversion", {"lookback_days": 12, "zscore_threshold": 1.1})
    assert knobs.lookback_days == 12

    with pytest.raises(ValueError):
        parse_entry_signal_config("breakout", {"breakout_window": 0}, default_lookback_days=21, default_skip_days=2)

    with pytest.raises(ValueError):
        parse_exit_signal_config("trailing_stop", {"trailing_stop_pct": 1.5}, default_lookback_days=21, default_skip_days=2)

    with pytest.raises(ValueError):
        parse_strategy_knobs("vol_carry", {"short_vol_window": 10, "long_vol_window": 5})


def test_parse_execution_model_config_supports_modular_components() -> None:
    cfg = parse_execution_model_config("modular", {"spread_bps": 3.0, "impact_bps": 7.0})
    assert cfg.name == "modular"
    assert float(cfg.params["spread_bps"]) == 3.0


def test_parse_execution_model_config_rejects_unknown_model() -> None:
    with pytest.raises(ValueError):
        parse_execution_model_config("bogus", {})

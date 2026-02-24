from __future__ import annotations

from dataclasses import dataclass
from typing import Protocol

import numpy as np

from .config import (
    BreakoutEntryConfig,
    EntrySignalConfig,
    MeanReversionEntryConfig,
    MovingAverageTrendEntryConfig,
    SeasonalityEventEntryConfig,
    TimeSeriesMomentumEntryConfig,
    TrendStrengthRegimeEntryConfig,
    VRPHarvestEntryConfig,
    VolatilityCarryEntryConfig,
)


class EntrySignal(Protocol):
    def value_at(self, idx: int, prices: np.ndarray, missing: np.ndarray) -> int:
        ...


@dataclass(frozen=True)
class TimeSeriesMomentumEntry:
    config: TimeSeriesMomentumEntryConfig

    def value_at(self, idx: int, prices: np.ndarray, missing: np.ndarray) -> int:
        end_idx = idx - self.config.skip_days
        start_idx = end_idx - self.config.lookback_days
        if start_idx < 0 or end_idx < 0:
            return 0
        if missing[start_idx] or missing[end_idx]:
            return 0
        start_px = prices[start_idx]
        end_px = prices[end_idx]
        if start_px <= 0.0 or end_px <= 0.0:
            return 0
        score = end_px / start_px - 1.0
        if abs(score) < self.config.min_abs_return:
            return 0
        side = int(np.sign(score))
        if self.config.long_only and side < 0:
            return 0
        return side


@dataclass(frozen=True)
class MovingAverageTrendEntry:
    config: MovingAverageTrendEntryConfig

    def value_at(self, idx: int, prices: np.ndarray, missing: np.ndarray) -> int:
        start_idx = idx - self.config.ma_window + 1
        if start_idx < 0:
            return 0
        window_missing = missing[start_idx : idx + 1]
        if bool(np.any(window_missing)):
            return 0
        window = prices[start_idx : idx + 1]
        ma = float(np.mean(window))
        return 1 if prices[idx] > ma else 0


@dataclass(frozen=True)
class BreakoutEntry:
    config: BreakoutEntryConfig

    def value_at(self, idx: int, prices: np.ndarray, missing: np.ndarray) -> int:
        start_idx = idx - self.config.breakout_window
        if start_idx < 0:
            return 0
        window_missing = missing[start_idx : idx + 1]
        if bool(np.any(window_missing)):
            return 0
        prior_window = prices[start_idx:idx]
        if prior_window.size == 0:
            return 0
        prior_high = float(np.max(prior_window))
        return 1 if prices[idx] > prior_high else 0


@dataclass(frozen=True)
class MeanReversionEntry:
    config: MeanReversionEntryConfig

    def value_at(self, idx: int, prices: np.ndarray, missing: np.ndarray) -> int:
        start_idx = idx - self.config.lookback_days + 1
        if start_idx < 0:
            return 0
        window_missing = missing[start_idx : idx + 1]
        if bool(np.any(window_missing)):
            return 0
        window = prices[start_idx : idx + 1]
        mean = float(np.mean(window))
        std = float(np.std(window))
        if std <= 0.0:
            return 0
        zscore = (float(prices[idx]) - mean) / std
        if zscore >= self.config.zscore_threshold:
            return -1 if not self.config.long_only else 0
        if zscore <= -self.config.zscore_threshold:
            return 1
        return 0


@dataclass(frozen=True)
class VolatilityCarryEntry:
    config: VolatilityCarryEntryConfig

    def value_at(self, idx: int, prices: np.ndarray, missing: np.ndarray) -> int:
        start_idx = idx - self.config.long_vol_window
        if start_idx < 0:
            return 0
        window_missing = missing[start_idx : idx + 1]
        if bool(np.any(window_missing)):
            return 0
        window = prices[start_idx : idx + 1]
        prev = window[:-1]
        curr = window[1:]
        valid = prev > 0.0
        if not bool(np.all(valid)):
            return 0
        rets = curr / prev - 1.0
        short = rets[-self.config.short_vol_window :]
        long = rets[-self.config.long_vol_window :]
        short_vol = float(np.std(short))
        long_vol = float(np.std(long))
        spread = long_vol - short_vol
        if spread > self.config.min_carry_spread:
            return 1
        if spread < -self.config.min_carry_spread:
            return -1
        return 0


@dataclass(frozen=True)
class TrendStrengthRegimeEntry:
    config: TrendStrengthRegimeEntryConfig

    def value_at(self, idx: int, prices: np.ndarray, missing: np.ndarray) -> int:
        start_idx = idx - max(self.config.trend_window, self.config.strength_window) + 1
        if start_idx < 0:
            return 0
        window_missing = missing[start_idx : idx + 1]
        if bool(np.any(window_missing)):
            return 0

        trend_window = prices[idx - self.config.trend_window + 1 : idx + 1]
        strength_window = prices[idx - self.config.strength_window + 1 : idx + 1]
        trend = float(trend_window[-1] - trend_window[0])
        diffs = np.diff(strength_window)
        directional_moves = np.sign(diffs)
        strength = float(np.mean(directional_moves >= 0))
        if strength < self.config.min_strength:
            strength = max(strength, float(np.mean(directional_moves <= 0)))
            if strength < self.config.min_strength:
                return 0
        return 1 if trend > 0.0 else (-1 if trend < 0.0 else 0)




@dataclass(frozen=True)
class VRPHarvestEntry:
    config: VRPHarvestEntryConfig

    def value_at(self, idx: int, prices: np.ndarray, missing: np.ndarray) -> int:
        start_idx = idx - self.config.realized_vol_lookback
        if start_idx < 0:
            return 0
        window_missing = missing[start_idx : idx + 1]
        if bool(np.any(window_missing)):
            return 0
        window = prices[start_idx : idx + 1]
        prev = window[:-1]
        curr = window[1:]
        if prev.size == 0 or not bool(np.all(prev > 0.0)):
            return 0
        realized_vol = float(np.std(curr / prev - 1.0))

        iv_idx = idx - 1 if self.config.regime_filter else idx
        if iv_idx < 0 or missing[iv_idx]:
            return 0
        implied_vol = float(prices[iv_idx])
        if not np.isfinite(implied_vol) or implied_vol <= 0.0 or not np.isfinite(realized_vol):
            return 0

        vrp = implied_vol - realized_vol
        threshold = self.config.vrp_threshold
        if vrp > threshold:
            side = -1
        elif vrp < -threshold:
            side = 1
        else:
            return 0

        if self.config.long_only and side < 0:
            return 0
        return side

@dataclass(frozen=True)
class SeasonalityEventEntry:
    config: SeasonalityEventEntryConfig

    def value_at(self, idx: int, prices: np.ndarray, missing: np.ndarray) -> int:
        event_idx = idx - self.config.event_offset
        if event_idx < 0:
            return 0
        if event_idx % self.config.seasonal_period >= self.config.event_window:
            return 0
        prev_idx = idx - self.config.seasonal_period
        if prev_idx < 0 or missing[idx] or missing[prev_idx]:
            return 0
        prev = float(prices[prev_idx])
        curr = float(prices[idx])
        if prev <= 0.0:
            return 0
        ret = curr / prev - 1.0
        side = int(np.sign(ret))
        if self.config.long_only and side < 0:
            return 0
        return side


def build_entry_signal(config: EntrySignalConfig) -> EntrySignal:
    if isinstance(config, TimeSeriesMomentumEntryConfig):
        return TimeSeriesMomentumEntry(config)
    if isinstance(config, MovingAverageTrendEntryConfig):
        return MovingAverageTrendEntry(config)
    if isinstance(config, BreakoutEntryConfig):
        return BreakoutEntry(config)
    if isinstance(config, MeanReversionEntryConfig):
        return MeanReversionEntry(config)
    if isinstance(config, VolatilityCarryEntryConfig):
        return VolatilityCarryEntry(config)
    if isinstance(config, TrendStrengthRegimeEntryConfig):
        return TrendStrengthRegimeEntry(config)
    if isinstance(config, SeasonalityEventEntryConfig):
        return SeasonalityEventEntry(config)
    if isinstance(config, VRPHarvestEntryConfig):
        return VRPHarvestEntry(config)
    raise TypeError(f"Unsupported entry signal config: {type(config).__name__}")

from __future__ import annotations

from dataclasses import dataclass
from typing import Protocol

import numpy as np

from .config import BreakoutEntryConfig, EntrySignalConfig, MovingAverageTrendEntryConfig, TimeSeriesMomentumEntryConfig


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
        return int(np.sign(score))


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


def build_entry_signal(config: EntrySignalConfig) -> EntrySignal:
    if isinstance(config, TimeSeriesMomentumEntryConfig):
        return TimeSeriesMomentumEntry(config)
    if isinstance(config, MovingAverageTrendEntryConfig):
        return MovingAverageTrendEntry(config)
    if isinstance(config, BreakoutEntryConfig):
        return BreakoutEntry(config)
    raise TypeError(f"Unsupported entry signal config: {type(config).__name__}")

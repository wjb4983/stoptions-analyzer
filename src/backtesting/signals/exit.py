from __future__ import annotations

from dataclasses import dataclass
from typing import Protocol

from .config import ExitSignalConfig, MaxHoldExitConfig, MomentumFlipExitConfig, NoExitConfig, TrailingStopExitConfig


@dataclass
class PositionState:
    side: int
    entry_price: float
    peak_price: float
    bars_held: int = 0


class ExitSignal(Protocol):
    def should_exit(self, idx: int, prices, missing, position: PositionState) -> bool:
        ...


@dataclass(frozen=True)
class NoExit:
    def should_exit(self, idx: int, prices, missing, position: PositionState) -> bool:
        return False


@dataclass(frozen=True)
class MomentumFlipExit:
    config: MomentumFlipExitConfig

    def should_exit(self, idx: int, prices, missing, position: PositionState) -> bool:
        end_idx = idx - self.config.skip_days
        start_idx = end_idx - self.config.lookback_days
        if start_idx < 0 or end_idx < 0:
            return False
        if missing[start_idx] or missing[end_idx]:
            return False
        start_px = prices[start_idx]
        end_px = prices[end_idx]
        if start_px <= 0.0 or end_px <= 0.0:
            return False
        score = end_px / start_px - 1.0
        if position.side > 0:
            return score < 0.0
        if position.side < 0:
            return score > 0.0
        return False


@dataclass(frozen=True)
class TrailingStopExit:
    config: TrailingStopExitConfig

    def should_exit(self, idx: int, prices, missing, position: PositionState) -> bool:
        if missing[idx]:
            return False
        current_price = float(prices[idx])
        if position.side > 0:
            stop_level = position.peak_price * (1.0 - self.config.trailing_stop_pct)
            return current_price <= stop_level
        stop_level = position.peak_price * (1.0 + self.config.trailing_stop_pct)
        return current_price >= stop_level


@dataclass(frozen=True)
class MaxHoldExit:
    config: MaxHoldExitConfig

    def should_exit(self, idx: int, prices, missing, position: PositionState) -> bool:
        return position.bars_held >= self.config.max_hold_bars


def build_exit_signal(config: ExitSignalConfig) -> ExitSignal:
    if isinstance(config, NoExitConfig):
        return NoExit()
    if isinstance(config, MomentumFlipExitConfig):
        return MomentumFlipExit(config)
    if isinstance(config, TrailingStopExitConfig):
        return TrailingStopExit(config)
    if isinstance(config, MaxHoldExitConfig):
        return MaxHoldExit(config)
    raise TypeError(f"Unsupported exit signal config: {type(config).__name__}")

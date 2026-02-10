from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Mapping


@dataclass(frozen=True)
class TimeSeriesMomentumEntryConfig:
    name: str = "ts_momentum"
    lookback_days: int = 90
    skip_days: int = 5
    min_abs_return: float = 0.0
    long_only: bool = False


@dataclass(frozen=True)
class MovingAverageTrendEntryConfig:
    name: str = "ma_trend"
    ma_window: int = 50


@dataclass(frozen=True)
class BreakoutEntryConfig:
    name: str = "breakout"
    breakout_window: int = 20


EntrySignalConfig = (
    TimeSeriesMomentumEntryConfig | MovingAverageTrendEntryConfig | BreakoutEntryConfig
)


@dataclass(frozen=True)
class NoExitConfig:
    name: str = "none"


@dataclass(frozen=True)
class MomentumFlipExitConfig:
    name: str = "momentum_flip"
    lookback_days: int = 90
    skip_days: int = 5
    min_abs_return: float = 0.0


@dataclass(frozen=True)
class TrailingStopExitConfig:
    name: str = "trailing_stop"
    trailing_stop_pct: float = 0.05


@dataclass(frozen=True)
class MaxHoldExitConfig:
    name: str = "max_hold"
    max_hold_bars: int = 20


ExitSignalConfig = NoExitConfig | MomentumFlipExitConfig | TrailingStopExitConfig | MaxHoldExitConfig


def _int(value: Any, field_name: str) -> int:
    try:
        parsed = int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{field_name} must be an integer") from exc
    if parsed <= 0:
        raise ValueError(f"{field_name} must be > 0")
    return parsed


def parse_entry_signal_config(
    signal_name: str,
    params: Mapping[str, Any] | None,
    *,
    default_lookback_days: int,
    default_skip_days: int,
) -> EntrySignalConfig:
    payload = dict(params or {})
    if signal_name == "ts_momentum":
        lookback_days = _int(payload.get("lookback_days", default_lookback_days), "lookback_days")
        skip_days = int(payload.get("skip_days", default_skip_days))
        if skip_days < 0:
            raise ValueError("skip_days must be >= 0")
        if skip_days >= lookback_days:
            raise ValueError("skip_days must be < lookback_days")
        min_abs_return = float(payload.get("min_abs_return", 0.0))
        if min_abs_return < 0:
            raise ValueError("min_abs_return must be >= 0")
        long_only = bool(payload.get("long_only", False))
        return TimeSeriesMomentumEntryConfig(
            lookback_days=lookback_days,
            skip_days=skip_days,
            min_abs_return=min_abs_return,
            long_only=long_only,
        )
    if signal_name == "ma_trend":
        return MovingAverageTrendEntryConfig(ma_window=_int(payload.get("ma_window", 50), "ma_window"))
    if signal_name == "breakout":
        return BreakoutEntryConfig(
            breakout_window=_int(payload.get("breakout_window", 20), "breakout_window")
        )
    raise ValueError(f"Unsupported entry signal: {signal_name}")


def parse_exit_signal_config(
    signal_name: str,
    params: Mapping[str, Any] | None,
    *,
    default_lookback_days: int,
    default_skip_days: int,
) -> ExitSignalConfig:
    payload = dict(params or {})
    if signal_name == "none":
        return NoExitConfig()
    if signal_name == "momentum_flip":
        lookback_days = _int(payload.get("lookback_days", default_lookback_days), "lookback_days")
        skip_days = int(payload.get("skip_days", default_skip_days))
        if skip_days < 0:
            raise ValueError("skip_days must be >= 0")
        if skip_days >= lookback_days:
            raise ValueError("skip_days must be < lookback_days")
        min_abs_return = float(payload.get("min_abs_return", 0.0))
        if min_abs_return < 0:
            raise ValueError("min_abs_return must be >= 0")
        return MomentumFlipExitConfig(
            lookback_days=lookback_days,
            skip_days=skip_days,
            min_abs_return=min_abs_return,
        )
    if signal_name == "trailing_stop":
        trailing_stop_pct = float(payload.get("trailing_stop_pct", 0.05))
        if trailing_stop_pct <= 0.0 or trailing_stop_pct >= 1.0:
            raise ValueError("trailing_stop_pct must be between 0 and 1")
        return TrailingStopExitConfig(trailing_stop_pct=trailing_stop_pct)
    if signal_name == "max_hold":
        return MaxHoldExitConfig(max_hold_bars=_int(payload.get("max_hold_bars", 20), "max_hold_bars"))
    raise ValueError(f"Unsupported exit signal: {signal_name}")


def required_lookback_window(entry: EntrySignalConfig, exit_cfg: ExitSignalConfig) -> int:
    def _entry_window() -> int:
        if isinstance(entry, TimeSeriesMomentumEntryConfig):
            return entry.lookback_days + entry.skip_days + 1
        if isinstance(entry, MovingAverageTrendEntryConfig):
            return entry.ma_window + 1
        if isinstance(entry, BreakoutEntryConfig):
            return entry.breakout_window + 1
        return 1

    def _exit_window() -> int:
        if isinstance(exit_cfg, MomentumFlipExitConfig):
            return exit_cfg.lookback_days + exit_cfg.skip_days + 1
        return 1

    return max(_entry_window(), _exit_window())

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


@dataclass(frozen=True)
class MeanReversionEntryConfig:
    name: str = "mean_reversion"
    lookback_days: int = 20
    zscore_threshold: float = 1.0
    long_only: bool = False


@dataclass(frozen=True)
class VolatilityCarryEntryConfig:
    name: str = "vol_carry"
    short_vol_window: int = 10
    long_vol_window: int = 30
    min_carry_spread: float = 0.0


@dataclass(frozen=True)
class TrendStrengthRegimeEntryConfig:
    name: str = "trend_strength"
    trend_window: int = 20
    strength_window: int = 20
    min_strength: float = 0.55


@dataclass(frozen=True)
class SeasonalityEventEntryConfig:
    name: str = "seasonality_event"
    seasonal_period: int = 5
    event_offset: int = 0
    event_window: int = 1
    long_only: bool = False


@dataclass(frozen=True)
class VRPHarvestEntryConfig:
    name: str = "vrp_harvest"
    iv_feature_name: str = "iv_1m"
    realized_vol_lookback: int = 21
    vrp_threshold: float = 0.0
    regime_filter: bool = False
    long_only: bool = False


@dataclass(frozen=True)
class CheapVolEntryConfig:
    name: str = "cheap_vol_long"
    iv_feature_name: str = "iv_1m"
    iv_z_window: int = 60
    cheap_z_cutoff: float = -1.0
    cross_section_rank_max: float | None = None
    term_structure_feature_name: str | None = None
    term_structure_mode: str = "any"
    max_holding_bars: int = 10


EntrySignalConfig = TimeSeriesMomentumEntryConfig | MovingAverageTrendEntryConfig | BreakoutEntryConfig | MeanReversionEntryConfig | VolatilityCarryEntryConfig | TrendStrengthRegimeEntryConfig | SeasonalityEventEntryConfig | VRPHarvestEntryConfig | CheapVolEntryConfig


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


def _float(value: Any, field_name: str) -> float:
    try:
        parsed = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{field_name} must be a float") from exc
    if not (parsed == parsed):
        raise ValueError(f"{field_name} must be finite")
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
    if signal_name == "mean_reversion":
        lookback_days = _int(payload.get("lookback_days", 20), "lookback_days")
        zscore_threshold = _float(payload.get("zscore_threshold", 1.0), "zscore_threshold")
        if zscore_threshold <= 0.0:
            raise ValueError("zscore_threshold must be > 0")
        return MeanReversionEntryConfig(
            lookback_days=lookback_days,
            zscore_threshold=zscore_threshold,
            long_only=bool(payload.get("long_only", False)),
        )
    if signal_name == "vol_carry":
        short_vol_window = _int(payload.get("short_vol_window", 10), "short_vol_window")
        long_vol_window = _int(payload.get("long_vol_window", 30), "long_vol_window")
        if short_vol_window >= long_vol_window:
            raise ValueError("short_vol_window must be < long_vol_window")
        min_carry_spread = _float(payload.get("min_carry_spread", 0.0), "min_carry_spread")
        if min_carry_spread < 0.0:
            raise ValueError("min_carry_spread must be >= 0")
        return VolatilityCarryEntryConfig(
            short_vol_window=short_vol_window,
            long_vol_window=long_vol_window,
            min_carry_spread=min_carry_spread,
        )
    if signal_name == "trend_strength":
        trend_window = _int(payload.get("trend_window", 20), "trend_window")
        strength_window = _int(payload.get("strength_window", 20), "strength_window")
        min_strength = _float(payload.get("min_strength", 0.55), "min_strength")
        if min_strength < 0.5 or min_strength > 1.0:
            raise ValueError("min_strength must be between 0.5 and 1.0")
        return TrendStrengthRegimeEntryConfig(
            trend_window=trend_window,
            strength_window=strength_window,
            min_strength=min_strength,
        )
    if signal_name == "seasonality_event":
        seasonal_period = _int(payload.get("seasonal_period", 5), "seasonal_period")
        event_window = _int(payload.get("event_window", 1), "event_window")
        event_offset = int(payload.get("event_offset", 0))
        if event_offset < 0:
            raise ValueError("event_offset must be >= 0")
        return SeasonalityEventEntryConfig(
            seasonal_period=seasonal_period,
            event_offset=event_offset,
            event_window=event_window,
            long_only=bool(payload.get("long_only", False)),
        )
    if signal_name == "vrp_harvest":
        iv_feature_name = str(payload.get("iv_feature_name", "iv_1m")).strip()
        if not iv_feature_name:
            raise ValueError("iv_feature_name must be a non-empty string")
        realized_vol_lookback = _int(payload.get("realized_vol_lookback", 21), "realized_vol_lookback")
        vrp_threshold = _float(payload.get("vrp_threshold", 0.0), "vrp_threshold")
        return VRPHarvestEntryConfig(
            iv_feature_name=iv_feature_name,
            realized_vol_lookback=realized_vol_lookback,
            vrp_threshold=vrp_threshold,
            regime_filter=bool(payload.get("regime_filter", False)),
            long_only=bool(payload.get("long_only", False)),
        )
    if signal_name == "cheap_vol_long":
        iv_feature_name = str(payload.get("iv_feature_name", "iv_1m")).strip()
        if not iv_feature_name:
            raise ValueError("iv_feature_name must be a non-empty string")
        iv_z_window = _int(payload.get("iv_z_window", 60), "iv_z_window")
        cheap_z_cutoff = _float(payload.get("cheap_z_cutoff", -1.0), "cheap_z_cutoff")
        cross_section_rank_raw = payload.get("cross_section_rank_max", None)
        cross_section_rank_max: float | None
        if cross_section_rank_raw is None:
            cross_section_rank_max = None
        else:
            cross_section_rank_max = _float(cross_section_rank_raw, "cross_section_rank_max")
            if cross_section_rank_max < 0.0 or cross_section_rank_max > 1.0:
                raise ValueError("cross_section_rank_max must be between 0 and 1")
        term_structure_feature_raw = payload.get("term_structure_feature_name", None)
        term_structure_feature_name = None if term_structure_feature_raw is None else str(term_structure_feature_raw).strip() or None
        term_structure_mode = str(payload.get("term_structure_mode", "any")).strip().lower()
        if term_structure_mode not in {"any", "contango", "backwardation"}:
            raise ValueError("term_structure_mode must be one of: any, contango, backwardation")
        max_holding_bars = _int(payload.get("max_holding_bars", 10), "max_holding_bars")
        return CheapVolEntryConfig(
            iv_feature_name=iv_feature_name,
            iv_z_window=iv_z_window,
            cheap_z_cutoff=cheap_z_cutoff,
            cross_section_rank_max=cross_section_rank_max,
            term_structure_feature_name=term_structure_feature_name,
            term_structure_mode=term_structure_mode,
            max_holding_bars=max_holding_bars,
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
        if isinstance(entry, MeanReversionEntryConfig):
            return entry.lookback_days + 1
        if isinstance(entry, VolatilityCarryEntryConfig):
            return entry.long_vol_window + 1
        if isinstance(entry, TrendStrengthRegimeEntryConfig):
            return max(entry.trend_window, entry.strength_window) + 1
        if isinstance(entry, SeasonalityEventEntryConfig):
            return max(entry.seasonal_period + entry.event_offset + entry.event_window, 1)
        if isinstance(entry, VRPHarvestEntryConfig):
            return entry.realized_vol_lookback + 1
        if isinstance(entry, CheapVolEntryConfig):
            return entry.iv_z_window + 1
        return 1

    def _exit_window() -> int:
        if isinstance(exit_cfg, MomentumFlipExitConfig):
            return exit_cfg.lookback_days + exit_cfg.skip_days + 1
        return 1

    return max(_entry_window(), _exit_window())


@dataclass(frozen=True)
class TimeSeriesMomentumKnobs:
    lookback_days: int
    skip_days: int
    min_abs_return: float


@dataclass(frozen=True)
class MeanReversionKnobs:
    lookback_days: int
    zscore_threshold: float


@dataclass(frozen=True)
class VolatilityCarryKnobs:
    short_vol_window: int
    long_vol_window: int
    min_carry_spread: float


@dataclass(frozen=True)
class TrendStrengthKnobs:
    trend_window: int
    strength_window: int
    min_strength: float


@dataclass(frozen=True)
class SeasonalityEventKnobs:
    seasonal_period: int
    event_offset: int
    event_window: int


@dataclass(frozen=True)
class VRPHarvestKnobs:
    realized_vol_lookback: int
    vrp_threshold: float


@dataclass(frozen=True)
class CheapVolKnobs:
    iv_z_window: int
    cheap_z_cutoff: float
    max_holding_bars: int


StrategyKnobSchema = (
    TimeSeriesMomentumKnobs
    | MeanReversionKnobs
    | VolatilityCarryKnobs
    | TrendStrengthKnobs
    | SeasonalityEventKnobs
    | VRPHarvestKnobs
    | CheapVolKnobs
)


def parse_strategy_knobs(strategy_name: str, params: Mapping[str, Any] | None) -> StrategyKnobSchema:
    payload = dict(params or {})
    if strategy_name == "ts_momentum":
        lookback_days = _int(payload.get("lookback_days", 90), "lookback_days")
        skip_days = int(payload.get("skip_days", 5))
        if skip_days < 0 or skip_days >= lookback_days:
            raise ValueError("skip_days must be >=0 and < lookback_days")
        min_abs_return = _float(payload.get("min_abs_return", 0.0), "min_abs_return")
        if min_abs_return < 0:
            raise ValueError("min_abs_return must be >= 0")
        return TimeSeriesMomentumKnobs(lookback_days=lookback_days, skip_days=skip_days, min_abs_return=min_abs_return)
    if strategy_name == "mean_reversion":
        lookback_days = _int(payload.get("lookback_days", 20), "lookback_days")
        zscore_threshold = _float(payload.get("zscore_threshold", 1.0), "zscore_threshold")
        if zscore_threshold <= 0:
            raise ValueError("zscore_threshold must be > 0")
        return MeanReversionKnobs(lookback_days=lookback_days, zscore_threshold=zscore_threshold)
    if strategy_name == "vol_carry":
        short_vol_window = _int(payload.get("short_vol_window", 10), "short_vol_window")
        long_vol_window = _int(payload.get("long_vol_window", 30), "long_vol_window")
        if short_vol_window >= long_vol_window:
            raise ValueError("short_vol_window must be < long_vol_window")
        min_carry_spread = _float(payload.get("min_carry_spread", 0.0), "min_carry_spread")
        return VolatilityCarryKnobs(
            short_vol_window=short_vol_window,
            long_vol_window=long_vol_window,
            min_carry_spread=min_carry_spread,
        )
    if strategy_name == "trend_strength":
        trend_window = _int(payload.get("trend_window", 20), "trend_window")
        strength_window = _int(payload.get("strength_window", 20), "strength_window")
        min_strength = _float(payload.get("min_strength", 0.55), "min_strength")
        if min_strength < 0.5 or min_strength > 1.0:
            raise ValueError("min_strength must be between 0.5 and 1.0")
        return TrendStrengthKnobs(
            trend_window=trend_window,
            strength_window=strength_window,
            min_strength=min_strength,
        )
    if strategy_name == "seasonality_event":
        seasonal_period = _int(payload.get("seasonal_period", 5), "seasonal_period")
        event_offset = int(payload.get("event_offset", 0))
        if event_offset < 0:
            raise ValueError("event_offset must be >= 0")
        event_window = _int(payload.get("event_window", 1), "event_window")
        return SeasonalityEventKnobs(
            seasonal_period=seasonal_period,
            event_offset=event_offset,
            event_window=event_window,
        )
    if strategy_name == "vrp_harvest":
        realized_vol_lookback = _int(payload.get("realized_vol_lookback", 21), "realized_vol_lookback")
        vrp_threshold = _float(payload.get("vrp_threshold", 0.0), "vrp_threshold")
        return VRPHarvestKnobs(realized_vol_lookback=realized_vol_lookback, vrp_threshold=vrp_threshold)
    if strategy_name == "cheap_vol_long":
        iv_z_window = _int(payload.get("iv_z_window", 60), "iv_z_window")
        cheap_z_cutoff = _float(payload.get("cheap_z_cutoff", -1.0), "cheap_z_cutoff")
        max_holding_bars = _int(payload.get("max_holding_bars", 10), "max_holding_bars")
        return CheapVolKnobs(
            iv_z_window=iv_z_window,
            cheap_z_cutoff=cheap_z_cutoff,
            max_holding_bars=max_holding_bars,
        )
    raise ValueError(f"Unsupported strategy knobs for: {strategy_name}")


@dataclass(frozen=True)
class ExecutionModelConfig:
    name: str = "bps"
    params: dict[str, float] | dict[str, object] | None = None


def parse_execution_model_config(model_name: str, params: Mapping[str, Any] | None) -> ExecutionModelConfig:
    payload = dict(params or {})
    name = str(model_name or "bps").strip().lower()
    if name in {"bps", "spread", "participation", "volatility_scaled", "square_root", "latency_drift", "modular"}:
        return ExecutionModelConfig(name=name, params=payload)
    raise ValueError(f"Unsupported execution model: {model_name}")

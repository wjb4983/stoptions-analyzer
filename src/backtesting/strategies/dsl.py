from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path
from typing import Any, Mapping

from ..signals.config import (
    EntrySignalConfig,
    ExecutionModelConfig,
    ExitSignalConfig,
    parse_entry_signal_config,
    parse_execution_model_config,
    parse_exit_signal_config,
)


class StrategyDSLValidationError(ValueError):
    """Raised when a strategy DSL payload fails schema validation."""


@dataclass(frozen=True)
class UniverseSpec:
    symbols: tuple[str, ...]
    asset_class: str = "equity"
    timezone: str = "UTC"
    rebalance_frequency: str = "1d"


@dataclass(frozen=True)
class FeatureSpec:
    name: str
    source: str
    params: Mapping[str, Any]


@dataclass(frozen=True)
class LogicSpec:
    name: str
    params: Mapping[str, Any]


@dataclass(frozen=True)
class RiskRule:
    name: str
    params: Mapping[str, Any]


@dataclass(frozen=True)
class ExecutionSettings:
    model: str
    params: Mapping[str, Any]


@dataclass(frozen=True)
class StrategyDefinition:
    name: str
    family: str
    universe: UniverseSpec
    features: tuple[FeatureSpec, ...]
    entry: LogicSpec
    exit: LogicSpec
    risk_rules: tuple[RiskRule, ...]
    execution: ExecutionSettings


@dataclass(frozen=True)
class ExecutableStrategy:
    definition: StrategyDefinition
    entry_config: EntrySignalConfig
    exit_config: ExitSignalConfig
    execution_model: ExecutionModelConfig


_ALLOWED_TOP_LEVEL = {
    "name",
    "family",
    "universe",
    "features",
    "logic",
    "risk",
    "execution",
}


def _err(path: str, message: str) -> StrategyDSLValidationError:
    return StrategyDSLValidationError(f"{path}: {message}")


def _ensure_mapping(value: Any, *, path: str) -> Mapping[str, Any]:
    if not isinstance(value, Mapping):
        raise _err(path, f"expected object/dict, got {type(value).__name__}")
    return value


def _ensure_non_empty_str(value: Any, *, path: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise _err(path, "expected non-empty string")
    return value.strip()


def _ensure_list(value: Any, *, path: str) -> list[Any]:
    if not isinstance(value, list):
        raise _err(path, f"expected array/list, got {type(value).__name__}")
    return value


def _parse_universe(node: Mapping[str, Any]) -> UniverseSpec:
    symbols_raw = _ensure_list(node.get("symbols"), path="universe.symbols")
    if not symbols_raw:
        raise _err("universe.symbols", "must include at least one symbol")
    symbols: list[str] = []
    for idx, symbol in enumerate(symbols_raw):
        symbols.append(_ensure_non_empty_str(symbol, path=f"universe.symbols[{idx}]"))

    return UniverseSpec(
        symbols=tuple(symbols),
        asset_class=_ensure_non_empty_str(node.get("asset_class", "equity"), path="universe.asset_class"),
        timezone=_ensure_non_empty_str(node.get("timezone", "UTC"), path="universe.timezone"),
        rebalance_frequency=_ensure_non_empty_str(
            node.get("rebalance_frequency", "1d"), path="universe.rebalance_frequency"
        ),
    )


def _parse_features(node: Any) -> tuple[FeatureSpec, ...]:
    features_raw = _ensure_list(node, path="features")
    parsed: list[FeatureSpec] = []
    for idx, item in enumerate(features_raw):
        feature = _ensure_mapping(item, path=f"features[{idx}]")
        params = _ensure_mapping(feature.get("params", {}), path=f"features[{idx}].params")
        parsed.append(
            FeatureSpec(
                name=_ensure_non_empty_str(feature.get("name"), path=f"features[{idx}].name"),
                source=_ensure_non_empty_str(feature.get("source", "price"), path=f"features[{idx}].source"),
                params=dict(params),
            )
        )
    return tuple(parsed)


def _parse_logic(node: Mapping[str, Any], *, leg: str) -> LogicSpec:
    logic_node = _ensure_mapping(node.get(leg), path=f"logic.{leg}")
    params = _ensure_mapping(logic_node.get("params", {}), path=f"logic.{leg}.params")
    return LogicSpec(
        name=_ensure_non_empty_str(logic_node.get("name"), path=f"logic.{leg}.name"),
        params=dict(params),
    )


def _parse_risk(node: Any) -> tuple[RiskRule, ...]:
    rules_raw = _ensure_list(node, path="risk.rules")
    parsed: list[RiskRule] = []
    for idx, item in enumerate(rules_raw):
        rule = _ensure_mapping(item, path=f"risk.rules[{idx}]")
        params = _ensure_mapping(rule.get("params", {}), path=f"risk.rules[{idx}].params")
        parsed.append(
            RiskRule(
                name=_ensure_non_empty_str(rule.get("name"), path=f"risk.rules[{idx}].name"),
                params=dict(params),
            )
        )
    return tuple(parsed)


def _parse_execution(node: Mapping[str, Any]) -> ExecutionSettings:
    params = _ensure_mapping(node.get("params", {}), path="execution.params")
    return ExecutionSettings(
        model=_ensure_non_empty_str(node.get("model", "bps"), path="execution.model"),
        params=dict(params),
    )


def validate_strategy_payload(payload: Mapping[str, Any]) -> None:
    unknown = sorted(set(payload.keys()) - _ALLOWED_TOP_LEVEL)
    if unknown:
        raise _err("$", f"unknown top-level keys: {unknown}")

    _ensure_non_empty_str(payload.get("name"), path="name")
    _ensure_non_empty_str(payload.get("family", "generic"), path="family")
    _parse_universe(_ensure_mapping(payload.get("universe"), path="universe"))
    _parse_features(payload.get("features", []))

    logic = _ensure_mapping(payload.get("logic"), path="logic")
    _parse_logic(logic, leg="entry")
    _parse_logic(logic, leg="exit")

    risk = _ensure_mapping(payload.get("risk", {}), path="risk")
    _parse_risk(risk.get("rules", []))

    _parse_execution(_ensure_mapping(payload.get("execution", {}), path="execution"))


def parse_strategy_payload(payload: Mapping[str, Any]) -> StrategyDefinition:
    validate_strategy_payload(payload)

    universe = _parse_universe(_ensure_mapping(payload["universe"], path="universe"))
    features = _parse_features(payload.get("features", []))
    logic = _ensure_mapping(payload["logic"], path="logic")
    risk = _ensure_mapping(payload.get("risk", {}), path="risk")

    return StrategyDefinition(
        name=_ensure_non_empty_str(payload["name"], path="name"),
        family=_ensure_non_empty_str(payload.get("family", "generic"), path="family"),
        universe=universe,
        features=features,
        entry=_parse_logic(logic, leg="entry"),
        exit=_parse_logic(logic, leg="exit"),
        risk_rules=_parse_risk(risk.get("rules", [])),
        execution=_parse_execution(_ensure_mapping(payload.get("execution", {}), path="execution")),
    )


def _load_yaml(text: str) -> Mapping[str, Any]:
    try:
        import yaml  # type: ignore
    except ModuleNotFoundError as exc:
        raise StrategyDSLValidationError(
            "YAML parsing requested but PyYAML is not installed; use JSON input or install pyyaml"
        ) from exc
    data = yaml.safe_load(text)
    return _ensure_mapping(data, path="$")


def parse_strategy_text(text: str, *, format_hint: str | None = None) -> StrategyDefinition:
    fmt = (format_hint or "").strip().lower()
    if fmt not in {"", "json", "yaml", "yml"}:
        raise StrategyDSLValidationError("format_hint must be one of: json, yaml, yml")

    if fmt in {"yaml", "yml"}:
        payload = _load_yaml(text)
    else:
        try:
            payload = json.loads(text)
        except json.JSONDecodeError:
            payload = _load_yaml(text)

    return parse_strategy_payload(_ensure_mapping(payload, path="$"))


def parse_strategy_file(path: str | Path) -> StrategyDefinition:
    file_path = Path(path)
    text = file_path.read_text(encoding="utf-8")
    suffix = file_path.suffix.lower()
    hint = "json" if suffix == ".json" else "yaml" if suffix in {".yaml", ".yml"} else None
    return parse_strategy_text(text, format_hint=hint)


def compile_strategy(definition: StrategyDefinition) -> ExecutableStrategy:
    entry_cfg = parse_entry_signal_config(
        definition.entry.name,
        definition.entry.params,
        default_lookback_days=90,
        default_skip_days=5,
    )
    exit_cfg = parse_exit_signal_config(
        definition.exit.name,
        definition.exit.params,
        default_lookback_days=90,
        default_skip_days=5,
    )
    execution_model = parse_execution_model_config(definition.execution.model, definition.execution.params)
    return ExecutableStrategy(
        definition=definition,
        entry_config=entry_cfg,
        exit_config=exit_cfg,
        execution_model=execution_model,
    )


def load_compiled_strategy(path: str | Path) -> ExecutableStrategy:
    return compile_strategy(parse_strategy_file(path))


BUILTIN_STRATEGY_TEMPLATES: dict[str, dict[str, Any]] = {
    "ts_momentum_core": {
        "name": "TS Momentum Core",
        "family": "trend_following",
        "universe": {"symbols": ["SPY", "QQQ", "TLT", "GLD"]},
        "features": [{"name": "returns", "source": "price", "params": {"lookback_days": 90}}],
        "logic": {
            "entry": {"name": "ts_momentum", "params": {"lookback_days": 90, "skip_days": 5}},
            "exit": {"name": "momentum_flip", "params": {"lookback_days": 90, "skip_days": 5}},
        },
        "risk": {"rules": [{"name": "max_position", "params": {"weight": 0.25}}]},
        "execution": {"model": "bps", "params": {"bps": 8.0}},
    },
    "ts_momentum_fast": {
        "name": "TS Momentum Fast",
        "family": "trend_following",
        "universe": {"symbols": ["SPY", "IWM", "EEM"]},
        "features": [{"name": "returns", "source": "price", "params": {"lookback_days": 45}}],
        "logic": {
            "entry": {"name": "ts_momentum", "params": {"lookback_days": 45, "skip_days": 2}},
            "exit": {"name": "trailing_stop", "params": {"trailing_stop_pct": 0.04}},
        },
        "risk": {"rules": [{"name": "vol_target", "params": {"target": 0.1}}]},
        "execution": {"model": "spread", "params": {"half_spread_bps": 2.0}},
    },
    "ts_momentum_long_only": {
        "name": "TS Momentum Long Only",
        "family": "trend_following",
        "universe": {"symbols": ["SPY", "QQQ", "DIA"]},
        "features": [{"name": "returns", "source": "price", "params": {"lookback_days": 120}}],
        "logic": {
            "entry": {"name": "ts_momentum", "params": {"lookback_days": 120, "skip_days": 5, "long_only": True}},
            "exit": {"name": "max_hold", "params": {"max_hold_bars": 60}},
        },
        "risk": {"rules": [{"name": "max_drawdown", "params": {"limit": 0.15}}]},
        "execution": {"model": "participation", "params": {"max_participation": 0.1}},
    },
    "ma_trend_swing": {
        "name": "MA Trend Swing",
        "family": "trend_following",
        "universe": {"symbols": ["AAPL", "MSFT", "NVDA"]},
        "features": [{"name": "moving_average", "source": "price", "params": {"window": 50}}],
        "logic": {"entry": {"name": "ma_trend", "params": {"ma_window": 50}}, "exit": {"name": "momentum_flip", "params": {}}},
        "risk": {"rules": [{"name": "max_position", "params": {"weight": 0.1}}]},
        "execution": {"model": "bps", "params": {"bps": 6.0}},
    },
    "breakout_20d": {
        "name": "Breakout 20D",
        "family": "breakout",
        "universe": {"symbols": ["CL=F", "GC=F", "SI=F"]},
        "features": [{"name": "range", "source": "price", "params": {"window": 20}}],
        "logic": {"entry": {"name": "breakout", "params": {"breakout_window": 20}}, "exit": {"name": "trailing_stop", "params": {"trailing_stop_pct": 0.06}}},
        "risk": {"rules": [{"name": "stop_loss", "params": {"pct": 0.06}}]},
        "execution": {"model": "volatility_scaled", "params": {"base_bps": 5.0}},
    },
    "mean_reversion_intraday": {
        "name": "Mean Reversion Intraday",
        "family": "mean_reversion",
        "universe": {"symbols": ["SPY", "QQQ", "IWM"], "rebalance_frequency": "30m"},
        "features": [{"name": "zscore", "source": "price", "params": {"lookback_days": 10}}],
        "logic": {"entry": {"name": "mean_reversion", "params": {"lookback_days": 10, "zscore_threshold": 1.8}}, "exit": {"name": "max_hold", "params": {"max_hold_bars": 8}}},
        "risk": {"rules": [{"name": "max_gross", "params": {"value": 1.5}}]},
        "execution": {"model": "latency_drift", "params": {"latency_ms": 30}},
    },
    "mean_reversion_swing": {
        "name": "Mean Reversion Swing",
        "family": "mean_reversion",
        "universe": {"symbols": ["XLF", "XLK", "XLE", "XLV"]},
        "features": [{"name": "zscore", "source": "price", "params": {"lookback_days": 20}}],
        "logic": {"entry": {"name": "mean_reversion", "params": {"lookback_days": 20, "zscore_threshold": 1.2}}, "exit": {"name": "momentum_flip", "params": {"lookback_days": 30}}},
        "risk": {"rules": [{"name": "sector_cap", "params": {"max_weight": 0.3}}]},
        "execution": {"model": "spread", "params": {"half_spread_bps": 1.5}},
    },
    "vol_carry_balanced": {
        "name": "Vol Carry Balanced",
        "family": "volatility_carry",
        "universe": {"symbols": ["VXX", "SVXY"]},
        "features": [{"name": "realized_vol", "source": "price", "params": {"short": 10, "long": 30}}],
        "logic": {"entry": {"name": "vol_carry", "params": {"short_vol_window": 10, "long_vol_window": 30}}, "exit": {"name": "trailing_stop", "params": {"trailing_stop_pct": 0.08}}},
        "risk": {"rules": [{"name": "max_notional", "params": {"value": 0.5}}]},
        "execution": {"model": "square_root", "params": {"eta": 0.5}},
    },
    "vol_carry_defensive": {
        "name": "Vol Carry Defensive",
        "family": "volatility_carry",
        "universe": {"symbols": ["VIXY", "VXZ"]},
        "features": [{"name": "term_structure", "source": "options", "params": {"tenors": [1, 2, 3]}}],
        "logic": {"entry": {"name": "vol_carry", "params": {"short_vol_window": 8, "long_vol_window": 40, "min_carry_spread": 0.02}}, "exit": {"name": "max_hold", "params": {"max_hold_bars": 15}}},
        "risk": {"rules": [{"name": "vol_target", "params": {"target": 0.08}}]},
        "execution": {"model": "bps", "params": {"bps": 12.0}},
    },
    "trend_strength_filter": {
        "name": "Trend Strength Filter",
        "family": "regime",
        "universe": {"symbols": ["SPY", "QQQ", "TLT", "DBC"]},
        "features": [{"name": "trend_strength", "source": "price", "params": {"trend_window": 20, "strength_window": 20}}],
        "logic": {"entry": {"name": "trend_strength", "params": {"trend_window": 20, "strength_window": 20, "min_strength": 0.6}}, "exit": {"name": "none", "params": {}}},
        "risk": {"rules": [{"name": "regime_deleverage", "params": {"factor": 0.5}}]},
        "execution": {"model": "modular", "params": {"slippage": {"model": "spread"}}},
    },
    "seasonality_turn_of_month": {
        "name": "Seasonality Turn Of Month",
        "family": "seasonality",
        "universe": {"symbols": ["SPY", "DIA", "IWM"]},
        "features": [{"name": "calendar", "source": "event", "params": {"event": "turn_of_month"}}],
        "logic": {"entry": {"name": "seasonality_event", "params": {"seasonal_period": 21, "event_offset": 19, "event_window": 3, "long_only": True}}, "exit": {"name": "max_hold", "params": {"max_hold_bars": 3}}},
        "risk": {"rules": [{"name": "max_position", "params": {"weight": 0.2}}]},
        "execution": {"model": "bps", "params": {"bps": 4.0}},
    },
    "seasonality_weekday_effect": {
        "name": "Seasonality Weekday Effect",
        "family": "seasonality",
        "universe": {"symbols": ["SPY", "QQQ"]},
        "features": [{"name": "calendar", "source": "event", "params": {"event": "weekday"}}],
        "logic": {"entry": {"name": "seasonality_event", "params": {"seasonal_period": 5, "event_offset": 0, "event_window": 1}}, "exit": {"name": "none", "params": {}}},
        "risk": {"rules": [{"name": "max_turnover", "params": {"daily": 0.4}}]},
        "execution": {"model": "spread", "params": {"half_spread_bps": 1.0}},
    },
    "vrp_harvest_core": {
        "name": "VRP Harvest Core",
        "family": "volatility_risk_premium",
        "universe": {"symbols": ["VIXY", "SVXY"]},
        "features": [{"name": "iv_surface", "source": "options", "params": {"iv_feature_name": "iv_1m"}}],
        "logic": {"entry": {"name": "vrp_harvest", "params": {"iv_feature_name": "iv_1m", "realized_vol_lookback": 21, "vrp_threshold": 0.02}}, "exit": {"name": "max_hold", "params": {"max_hold_bars": 5}}},
        "risk": {"rules": [{"name": "max_notional", "params": {"value": 0.35}}]},
        "execution": {"model": "bps", "params": {"bps": 10.0}},
    },
    "equity_index_rotation": {
        "name": "Equity Index Rotation",
        "family": "rotation",
        "universe": {"symbols": ["SPY", "QQQ", "IWM", "EFA", "EEM"], "rebalance_frequency": "1w"},
        "features": [{"name": "relative_strength", "source": "price", "params": {"lookback_days": 60}}],
        "logic": {"entry": {"name": "ts_momentum", "params": {"lookback_days": 60, "skip_days": 5, "long_only": True}}, "exit": {"name": "momentum_flip", "params": {"lookback_days": 60, "skip_days": 5}}},
        "risk": {"rules": [{"name": "top_k", "params": {"k": 2}}]},
        "execution": {"model": "participation", "params": {"max_participation": 0.15}},
    },
    "futures_trend": {
        "name": "Futures Trend",
        "family": "trend_following",
        "universe": {"symbols": ["ES=F", "NQ=F", "ZN=F", "CL=F", "GC=F"]},
        "features": [{"name": "returns", "source": "price", "params": {"lookback_days": 180}}],
        "logic": {"entry": {"name": "ts_momentum", "params": {"lookback_days": 180, "skip_days": 10}}, "exit": {"name": "trailing_stop", "params": {"trailing_stop_pct": 0.07}}},
        "risk": {"rules": [{"name": "risk_parity", "params": {"target_vol": 0.12}}]},
        "execution": {"model": "square_root", "params": {"eta": 0.7}},
    },
    "etf_risk_parity_momentum": {
        "name": "ETF Risk Parity Momentum",
        "family": "allocation",
        "universe": {"symbols": ["SPY", "TLT", "GLD", "DBC", "VNQ"], "rebalance_frequency": "1w"},
        "features": [{"name": "returns", "source": "price", "params": {"lookback_days": 120}}, {"name": "realized_vol", "source": "price", "params": {"window": 30}}],
        "logic": {"entry": {"name": "ts_momentum", "params": {"lookback_days": 120, "skip_days": 5, "long_only": True}}, "exit": {"name": "none", "params": {}}},
        "risk": {"rules": [{"name": "risk_parity", "params": {"target_vol": 0.1}}, {"name": "max_position", "params": {"weight": 0.35}}]},
        "execution": {"model": "modular", "params": {"slippage": {"model": "square_root", "eta": 0.5}}},
    },
    "macro_regime_switch": {
        "name": "Macro Regime Switch",
        "family": "regime",
        "universe": {"symbols": ["SPY", "TLT", "GLD", "UUP", "HYG"], "rebalance_frequency": "1w"},
        "features": [{"name": "macro_regime", "source": "macro", "params": {"regimes": ["risk_on", "risk_off"]}}],
        "logic": {"entry": {"name": "trend_strength", "params": {"trend_window": 30, "strength_window": 30, "min_strength": 0.65}}, "exit": {"name": "max_hold", "params": {"max_hold_bars": 10}}},
        "risk": {"rules": [{"name": "regime_budget", "params": {"risk_on": 1.0, "risk_off": 0.5}}]},
        "execution": {"model": "volatility_scaled", "params": {"base_bps": 7.0}},
    },
}


def list_template_names() -> tuple[str, ...]:
    return tuple(sorted(BUILTIN_STRATEGY_TEMPLATES.keys()))


def get_template_payload(name: str) -> dict[str, Any]:
    if name not in BUILTIN_STRATEGY_TEMPLATES:
        allowed = ", ".join(list_template_names())
        raise KeyError(f"Unknown template '{name}'. Available templates: {allowed}")
    return dict(BUILTIN_STRATEGY_TEMPLATES[name])


def compile_template(name: str) -> ExecutableStrategy:
    payload = get_template_payload(name)
    return compile_strategy(parse_strategy_payload(payload))

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from backtesting.regime_builder import RegimeLegSpec, _REQUIRED_KNOBS_BY_LEG_FAMILY
from models.regime_catalog import validate_model_leg_pairing

_UI_TO_LEG_FAMILY: dict[str, str] = {
    "Trend Following": "timeseries_momentum",
    "Mean Reversion": "cheap_vol_buying",
    "Volatility Breakout": "volatility_risk_premium_selling",
    "Regime Change": "regime_change_detection",
    "Volatility Clustering": "volatility_clustering",
    "IV/EV Spread": "iv_ev_spread_term_structure",
    "Event Intensity": "self_exciting_event_intensity",
    "Vol Surface": "vol_surface_calibration",
}

_DEFAULT_MODEL_CANDIDATES_BY_LEG_FAMILY: dict[str, tuple[str, ...]] = {
    "timeseries_momentum": ("momentum", "momentum_forecasting", "meta_label_classifier"),
    "cheap_vol_buying": (
        "mean_reversion",
        "cheap_vol_mean_reversion_timing",
        "cheap_vol_event_timing",
        "meta_label_classifier",
    ),
    "volatility_risk_premium_selling": (
        "volatility_carry",
        "vrp_carry_relative_value",
        "meta_label_classifier",
    ),
    "regime_change_detection": (
        "hmm_regime_change",
        "markov_regime_switching",
        "changepoint_regime_change",
        "meta_label_classifier",
    ),
    "volatility_clustering": (
        "volatility_carry",
        "options_volatility",
        "term_structure_slope",
        "meta_label_classifier",
    ),
    "iv_ev_spread_term_structure": (
        "term_structure_slope",
        "vrp_carry_relative_value",
        "dispersion",
        "meta_label_classifier",
    ),
    "self_exciting_event_intensity": (
        "event_driven",
        "options_flow_driven",
        "microstructure_imbalance",
        "meta_label_classifier",
    ),
    "vol_surface_calibration": (
        "options_volatility",
        "term_structure_slope",
        "volatility_carry",
        "meta_label_classifier",
    ),
}


@dataclass(frozen=True)
class RegimeLegMapping:
    leg_spec: RegimeLegSpec
    default_model_candidates: tuple[str, ...]


_KNOB_TRANSLATIONS_BY_LEG_FAMILY: dict[str, dict[str, str]] = {
    "timeseries_momentum": {
        "max_position_pct": "sizing_cap",
        "max_drawdown_stop": "stop_loss_pct",
    },
    "cheap_vol_buying": {
        "entry_zscore": "carry_threshold",
        "max_position_pct": "sizing_cap",
        "max_drawdown_stop": "stop_loss_pct",
        "model_confidence_min": "vol_filter_max",
    },
    "volatility_risk_premium_selling": {
        "entry_zscore": "carry_threshold",
        "max_position_pct": "sizing_cap",
        "max_drawdown_stop": "stop_loss_pct",
        "model_confidence_min": "vol_filter_min",
    },
    "regime_change_detection": {
        "entry_zscore": "detection_threshold",
        "max_position_pct": "sizing_cap",
        "max_drawdown_stop": "stop_loss_pct",
    },
    "volatility_clustering": {
        "entry_zscore": "vol_filter_min",
        "model_confidence_min": "vol_filter_max",
        "max_position_pct": "sizing_cap",
        "max_drawdown_stop": "stop_loss_pct",
    },
    "iv_ev_spread_term_structure": {
        "entry_zscore": "carry_threshold",
        "model_confidence_min": "vol_filter_min",
        "max_position_pct": "sizing_cap",
        "max_drawdown_stop": "stop_loss_pct",
    },
    "self_exciting_event_intensity": {
        "entry_zscore": "detection_threshold",
        "model_confidence_min": "vol_filter_min",
        "max_position_pct": "sizing_cap",
        "max_drawdown_stop": "stop_loss_pct",
    },
    "vol_surface_calibration": {
        "entry_zscore": "detection_threshold",
        "model_confidence_min": "vol_filter_max",
        "max_position_pct": "sizing_cap",
        "max_drawdown_stop": "stop_loss_pct",
    },
}

_KNOB_DEFAULTS_BY_LEG_FAMILY: dict[str, dict[str, float]] = {
    "timeseries_momentum": {
        "lookback_days": 90.0,
        "vol_filter_max": 0.65,
        "sizing_cap": 0.08,
        "stop_loss_pct": 0.12,
    },
    "cheap_vol_buying": {
        "lookback_days": 30.0,
        "carry_threshold": -0.15,
        "vol_filter_max": 0.60,
        "sizing_cap": 0.05,
        "stop_loss_pct": 0.09,
    },
    "volatility_risk_premium_selling": {
        "lookback_days": 45.0,
        "carry_threshold": 0.15,
        "vol_filter_min": 0.70,
        "sizing_cap": 0.06,
        "stop_loss_pct": 0.10,
    },
    "regime_change_detection": {
        "lookback_days": 60.0,
        "detection_threshold": 1.60,
        "sizing_cap": 0.05,
        "stop_loss_pct": 0.08,
    },
    "volatility_clustering": {
        "lookback_days": 63.0,
        "vol_filter_min": 0.2,
        "vol_filter_max": 0.85,
        "sizing_cap": 0.05,
        "stop_loss_pct": 0.1,
    },
    "iv_ev_spread_term_structure": {
        "lookback_days": 45.0,
        "carry_threshold": 0.05,
        "vol_filter_min": 0.6,
        "sizing_cap": 0.05,
        "stop_loss_pct": 0.09,
    },
    "self_exciting_event_intensity": {
        "lookback_days": 15.0,
        "detection_threshold": 1.4,
        "vol_filter_min": 0.55,
        "sizing_cap": 0.03,
        "stop_loss_pct": 0.07,
    },
    "vol_surface_calibration": {
        "lookback_days": 30.0,
        "detection_threshold": 1.2,
        "vol_filter_max": 0.75,
        "sizing_cap": 0.04,
        "stop_loss_pct": 0.08,
    },
}


def supported_ui_legs() -> tuple[str, ...]:
    return tuple(_UI_TO_LEG_FAMILY)


def to_regime_leg_spec(ui_leg: dict[str, Any]) -> RegimeLegMapping:
    ui_leg_type = str(ui_leg.get("model_type", "")).strip()
    if ui_leg_type not in _UI_TO_LEG_FAMILY:
        supported = ", ".join(sorted(_UI_TO_LEG_FAMILY))
        raise ValueError(f"Unsupported UI leg '{ui_leg_type}'. Supported options: {supported}")

    leg_family = _UI_TO_LEG_FAMILY[ui_leg_type]
    controls = ui_leg.get("controls", {})
    if not isinstance(controls, dict):
        raise ValueError(f"Invalid controls payload for UI leg '{ui_leg_type}'.")

    translated_knobs = _translate_knobs(leg_family, controls)
    required = _REQUIRED_KNOBS_BY_LEG_FAMILY[leg_family]
    missing = [key for key in required if key not in translated_knobs]
    if missing:
        raise ValueError(
            f"Mapping for '{ui_leg_type}' is incomplete; missing backend knobs: {', '.join(missing)}"
        )

    model_candidates = _DEFAULT_MODEL_CANDIDATES_BY_LEG_FAMILY[leg_family]
    for model_name in model_candidates:
        validate_model_leg_pairing(leg_family, model_name)

    leg_spec = RegimeLegSpec(
        name=str(ui_leg.get("name", f"{ui_leg_type} leg")),
        leg_family=leg_family,
        knobs=translated_knobs,
        enabled=True,
    )
    return RegimeLegMapping(leg_spec=leg_spec, default_model_candidates=model_candidates)


def _translate_knobs(leg_family: str, controls: dict[str, Any]) -> dict[str, float]:
    translated = dict(_KNOB_DEFAULTS_BY_LEG_FAMILY[leg_family])
    key_map = _KNOB_TRANSLATIONS_BY_LEG_FAMILY[leg_family]

    for ui_key, raw_value in controls.items():
        if ui_key in translated:
            translated[ui_key] = float(raw_value)
            continue
        backend_key = key_map.get(ui_key)
        if backend_key is None:
            continue
        translated[backend_key] = float(raw_value)

    return translated

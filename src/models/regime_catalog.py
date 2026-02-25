from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from .registry import MODEL_REGISTRY


@dataclass(frozen=True)
class ModelDescriptor:
    model_name: str
    display_name: str
    hyperparameter_template: dict[str, Any]


_LEG_MODEL_CATALOG: dict[str, tuple[ModelDescriptor, ...]] = {
    "timeseries_momentum": (
        ModelDescriptor(
            model_name="momentum",
            display_name="Classic Momentum",
            hyperparameter_template={"lookback_days": 126, "smoothing_window": 20},
        ),
        ModelDescriptor(
            model_name="momentum_forecasting",
            display_name="Momentum Forecasting",
            hyperparameter_template={"lookback_days": 252, "trend_window": 20, "vol_window": 20},
        ),
        ModelDescriptor(
            model_name="meta_label_classifier",
            display_name="Meta Label Classifier",
            hyperparameter_template={"confidence_threshold": 0.55, "calibration": "conformal"},
        ),
    ),
    "volatility_risk_premium_selling": (
        ModelDescriptor(
            model_name="volatility_carry",
            display_name="Volatility Carry",
            hyperparameter_template={"carry_threshold": 0.1, "rebalance_days": 5},
        ),
        ModelDescriptor(
            model_name="vrp_carry_relative_value",
            display_name="VRP Carry/Relative Value",
            hyperparameter_template={"carry_threshold": 0.08, "relative_value_quantile": 0.7},
        ),
        ModelDescriptor(
            model_name="meta_label_classifier",
            display_name="Meta Label Classifier",
            hyperparameter_template={"confidence_threshold": 0.6, "calibration": "conformal"},
        ),
    ),
    "cheap_vol_buying": (
        ModelDescriptor(
            model_name="mean_reversion",
            display_name="Mean Reversion",
            hyperparameter_template={"reversion_half_life": 5, "entry_zscore": -1.25},
        ),
        ModelDescriptor(
            model_name="cheap_vol_event_timing",
            display_name="Cheap-Vol Event Timing",
            hyperparameter_template={"event_window_days": 7, "iv_rv_trigger": -0.05},
        ),
        ModelDescriptor(
            model_name="cheap_vol_mean_reversion_timing",
            display_name="Cheap-Vol Mean Reversion Timing",
            hyperparameter_template={"entry_zscore": -1.5, "exit_zscore": -0.25},
        ),
        ModelDescriptor(
            model_name="meta_label_classifier",
            display_name="Meta Label Classifier",
            hyperparameter_template={"confidence_threshold": 0.58, "calibration": "conformal"},
        ),
    ),
    "regime_change_detection": (
        ModelDescriptor(
            model_name="hmm_regime_change",
            display_name="HMM Regime Change",
            hyperparameter_template={"n_states": 3, "em_iterations": 50},
        ),
        ModelDescriptor(
            model_name="markov_regime_switching",
            display_name="Markov Regime Switching",
            hyperparameter_template={"n_states": 3, "em_iterations": 30},
        ),
        ModelDescriptor(
            model_name="changepoint_regime_change",
            display_name="Changepoint Regime Change",
            hyperparameter_template={"penalty": 3.0, "min_segment_length": 10},
        ),
        ModelDescriptor(
            model_name="meta_label_classifier",
            display_name="Meta Label Classifier",
            hyperparameter_template={"confidence_threshold": 0.57, "calibration": "conformal"},
        ),
    ),
}


for _leg_type, _descriptors in _LEG_MODEL_CATALOG.items():
    for _descriptor in _descriptors:
        if _descriptor.model_name not in MODEL_REGISTRY:
            raise RuntimeError(
                f"Catalog model '{_descriptor.model_name}' for leg '{_leg_type}' is not in MODEL_REGISTRY"
            )


def list_models_for_leg(leg_type: str) -> list[ModelDescriptor]:
    """Return dropdown-ready model descriptors for a regime leg type."""
    return list(_LEG_MODEL_CATALOG.get(leg_type, ()))


def is_model_allowed_for_leg(leg_type: str, model_name: str) -> bool:
    return any(item.model_name == model_name for item in _LEG_MODEL_CATALOG.get(leg_type, ()))


def validate_model_leg_pairing(leg_type: str, model_name: str) -> None:
    if not list_models_for_leg(leg_type):
        raise ValueError(f"Unsupported leg_type: {leg_type}")
    if not is_model_allowed_for_leg(leg_type, model_name):
        allowed = ", ".join(item.model_name for item in list_models_for_leg(leg_type))
        raise ValueError(
            f"Model '{model_name}' is not allowed for leg '{leg_type}'. Allowed models: {allowed}"
        )


__all__ = [
    "ModelDescriptor",
    "is_model_allowed_for_leg",
    "list_models_for_leg",
    "validate_model_leg_pairing",
]

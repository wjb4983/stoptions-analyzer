from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from .registry import MODEL_REGISTRY


@dataclass(frozen=True)
class ModelDescriptor:
    model_name: str
    display_name: str
    hyperparameter_template: dict[str, Any]
    capability_tags: frozenset[str] = frozenset()
    catalog_phase: str = "phase_1"


_LEG_MODEL_CATALOG: dict[str, tuple[ModelDescriptor, ...]] = {
    "timeseries_momentum": (
        ModelDescriptor(
            model_name="momentum",
            display_name="Classic Momentum",
            hyperparameter_template={"lookback_days": 126, "smoothing_window": 20},
            catalog_phase="phase_0",
        ),
        ModelDescriptor(
            model_name="momentum_forecasting",
            display_name="Momentum Forecasting",
            hyperparameter_template={"lookback_days": 252, "trend_window": 20, "vol_window": 20},
            catalog_phase="phase_0",
        ),
        ModelDescriptor(
            model_name="meta_label_classifier",
            display_name="Meta Label Classifier",
            hyperparameter_template={"confidence_threshold": 0.55, "calibration": "conformal"},
            catalog_phase="phase_0",
        ),
    ),
    "volatility_risk_premium_selling": (
        ModelDescriptor(
            model_name="volatility_carry",
            display_name="Volatility Carry",
            hyperparameter_template={"carry_threshold": 0.1, "rebalance_days": 5},
            catalog_phase="phase_0",
        ),
        ModelDescriptor(
            model_name="vrp_carry_relative_value",
            display_name="VRP Carry/Relative Value",
            hyperparameter_template={"carry_threshold": 0.08, "relative_value_quantile": 0.7},
            catalog_phase="phase_0",
        ),
        ModelDescriptor(
            model_name="meta_label_classifier",
            display_name="Meta Label Classifier",
            hyperparameter_template={"confidence_threshold": 0.6, "calibration": "conformal"},
            catalog_phase="phase_0",
        ),
    ),
    "cheap_vol_buying": (
        ModelDescriptor(
            model_name="mean_reversion",
            display_name="Mean Reversion",
            hyperparameter_template={"reversion_half_life": 5, "entry_zscore": -1.25},
            catalog_phase="phase_0",
        ),
        ModelDescriptor(
            model_name="cheap_vol_event_timing",
            display_name="Cheap-Vol Event Timing",
            hyperparameter_template={"event_window_days": 7, "iv_rv_trigger": -0.05},
            catalog_phase="phase_0",
        ),
        ModelDescriptor(
            model_name="cheap_vol_mean_reversion_timing",
            display_name="Cheap-Vol Mean Reversion Timing",
            hyperparameter_template={"entry_zscore": -1.5, "exit_zscore": -0.25},
            catalog_phase="phase_0",
        ),
        ModelDescriptor(
            model_name="meta_label_classifier",
            display_name="Meta Label Classifier",
            hyperparameter_template={"confidence_threshold": 0.58, "calibration": "conformal"},
            catalog_phase="phase_0",
        ),
    ),
    "regime_change_detection": (
        ModelDescriptor(
            model_name="hmm_regime_change",
            display_name="HMM Regime Change",
            hyperparameter_template={"n_states": 3, "em_iterations": 50},
            catalog_phase="phase_0",
        ),
        ModelDescriptor(
            model_name="markov_regime_switching",
            display_name="Markov Regime Switching",
            hyperparameter_template={"n_states": 3, "em_iterations": 30},
            catalog_phase="phase_0",
        ),
        ModelDescriptor(
            model_name="changepoint_regime_change",
            display_name="Changepoint Regime Change",
            hyperparameter_template={"penalty": 3.0, "min_segment_length": 10},
            catalog_phase="phase_0",
        ),
        ModelDescriptor(
            model_name="meta_label_classifier",
            display_name="Meta Label Classifier",
            hyperparameter_template={"confidence_threshold": 0.57, "calibration": "conformal"},
            catalog_phase="phase_0",
        ),
    ),
    "volatility_clustering": (
        ModelDescriptor(
            model_name="volatility_carry",
            display_name="GARCH-Style Volatility Clustering",
            hyperparameter_template={"p_order": 1, "q_order": 1, "decay_lambda": 0.94},
            catalog_phase="phase_1",
        ),
        ModelDescriptor(
            model_name="options_volatility",
            display_name="EGARCH Asymmetry Wrapper",
            hyperparameter_template={"asymmetry_term": 0.2, "vol_of_vol_window": 30},
            catalog_phase="phase_1",
        ),
        ModelDescriptor(
            model_name="term_structure_slope",
            display_name="HAR-RV Horizon Blend",
            hyperparameter_template={"daily_window": 1, "weekly_window": 5, "monthly_window": 22},
            catalog_phase="phase_1",
        ),
        ModelDescriptor(
            model_name="meta_label_classifier",
            display_name="Meta Label Classifier",
            hyperparameter_template={"confidence_threshold": 0.61, "calibration": "conformal"},
            catalog_phase="phase_1",
        ),
    ),
    "iv_ev_spread_term_structure": (
        ModelDescriptor(
            model_name="term_structure_slope",
            display_name="IV/EV Term Slope Decomposition",
            hyperparameter_template={"front_tenor_days": 14, "back_tenor_days": 60, "slope_threshold": 0.03},
            catalog_phase="phase_1",
        ),
        ModelDescriptor(
            model_name="vrp_carry_relative_value",
            display_name="IV-EV Carry Relative Value",
            hyperparameter_template={"iv_ev_spread_threshold": 0.05, "cross_tenor_quantile": 0.75},
            catalog_phase="phase_1",
        ),
        ModelDescriptor(
            model_name="dispersion",
            display_name="Cross-Sectional Spread Decomposition",
            hyperparameter_template={"dispersion_window": 20, "cross_asset_decay": 0.9},
            catalog_phase="phase_1",
        ),
        ModelDescriptor(
            model_name="meta_label_classifier",
            display_name="Meta Label Classifier",
            hyperparameter_template={"confidence_threshold": 0.59, "calibration": "conformal"},
            catalog_phase="phase_1",
        ),
    ),
    "self_exciting_event_intensity": (
        ModelDescriptor(
            model_name="event_driven",
            display_name="Hawkes Event Baseline",
            hyperparameter_template={"baseline_intensity": 0.05, "decay_half_life_minutes": 45},
            capability_tags=frozenset({"needs_event_process_spec"}),
            catalog_phase="phase_1",
        ),
        ModelDescriptor(
            model_name="options_flow_driven",
            display_name="Flow-Triggered Self-Excitation",
            hyperparameter_template={"excitation_alpha": 0.35, "event_window_minutes": 60},
            capability_tags=frozenset({"needs_event_process_spec"}),
            catalog_phase="phase_1",
        ),
        ModelDescriptor(
            model_name="microstructure_imbalance",
            display_name="Order-Book Hawkes Proxy",
            hyperparameter_template={"queue_reactivity": 0.4, "imbalance_decay_minutes": 10},
            capability_tags=frozenset({"needs_event_process_spec"}),
            catalog_phase="phase_1",
        ),
        ModelDescriptor(
            model_name="meta_label_classifier",
            display_name="Meta Label Classifier",
            hyperparameter_template={"confidence_threshold": 0.62, "calibration": "conformal"},
            catalog_phase="phase_1",
        ),
    ),
    "vol_surface_calibration": (
        ModelDescriptor(
            model_name="options_volatility",
            display_name="Local Vol Surface Wrapper",
            hyperparameter_template={"surface_grid_points": 25, "smoothing_penalty": 1e-4},
            capability_tags=frozenset({"needs_calibration_spec"}),
            catalog_phase="phase_1",
        ),
        ModelDescriptor(
            model_name="term_structure_slope",
            display_name="SABR Term/Skew Calibration",
            hyperparameter_template={"beta": 0.5, "nu_init": 0.3, "rho_init": -0.2},
            capability_tags=frozenset({"needs_calibration_spec"}),
            catalog_phase="phase_1",
        ),
        ModelDescriptor(
            model_name="volatility_carry",
            display_name="Heston Calibration Proxy",
            hyperparameter_template={"kappa_init": 1.5, "theta_init": 0.04, "xi_init": 0.4},
            capability_tags=frozenset({"needs_calibration_spec"}),
            catalog_phase="phase_1",
        ),
        ModelDescriptor(
            model_name="meta_label_classifier",
            display_name="Meta Label Classifier",
            hyperparameter_template={"confidence_threshold": 0.6, "calibration": "conformal"},
            catalog_phase="phase_1",
        ),
    ),
    "cross_asset_macro_conditioned": (
        ModelDescriptor(
            model_name="macro_regime_conditioned",
            display_name="Macro-Conditioned Regime Model",
            hyperparameter_template={"macro_lookback_days": 126, "shock_sensitivity": 0.8},
            capability_tags=frozenset({"supports_architecture_spec"}),
            catalog_phase="phase_2",
        ),
        ModelDescriptor(
            model_name="factor_neutral_cross_sectional_rank",
            display_name="Cross-Asset Factor-Neutral Ranker",
            hyperparameter_template={"ranking_horizon_days": 63, "neutralization_strength": 0.7},
            catalog_phase="phase_2",
        ),
        ModelDescriptor(
            model_name="dispersion",
            display_name="Cross-Asset Dispersion Blending",
            hyperparameter_template={"dispersion_window": 30, "cross_asset_decay": 0.85},
            catalog_phase="phase_2",
        ),
        ModelDescriptor(
            model_name="meta_label_classifier",
            display_name="Meta Label Classifier",
            hyperparameter_template={"confidence_threshold": 0.63, "calibration": "conformal"},
            catalog_phase="phase_2",
        ),
    ),
    "meta_label_regime_ensemble": (
        ModelDescriptor(
            model_name="meta_label_classifier",
            display_name="Regime Meta-Label Stack",
            hyperparameter_template={"confidence_threshold": 0.64, "calibration": "conformal", "ensemble_depth": 3},
            capability_tags=frozenset({"supports_architecture_spec"}),
            catalog_phase="phase_2",
        ),
        ModelDescriptor(
            model_name="markov_regime_switching",
            display_name="State-Machine Ensemble Gating",
            hyperparameter_template={"n_states": 4, "em_iterations": 35},
            catalog_phase="phase_2",
        ),
        ModelDescriptor(
            model_name="momentum_forecasting",
            display_name="Forecast Layer Base Learner",
            hyperparameter_template={"lookback_days": 126, "trend_window": 15, "vol_window": 15},
            catalog_phase="phase_2",
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


def get_model_descriptor(leg_type: str, model_name: str) -> ModelDescriptor | None:
    for descriptor in _LEG_MODEL_CATALOG.get(leg_type, ()):
        if descriptor.model_name == model_name:
            return descriptor
    return None


__all__ = [
    "ModelDescriptor",
    "is_model_allowed_for_leg",
    "list_models_for_leg",
    "get_model_descriptor",
    "validate_model_leg_pairing",
]

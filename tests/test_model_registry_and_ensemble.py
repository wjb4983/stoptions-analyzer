from __future__ import annotations

import numpy as np

from models.ensemble import ModelEnsembler
from models.registry import MODEL_REGISTRY, ModelActivationConfig, activated_models


ALL_PARADIGMS = {
    "momentum",
    "mean_reversion",
    "volatility_carry",
    "term_structure_slope",
    "dispersion",
    "stat_arb_pair_spread",
    "factor_neutral_cross_sectional_rank",
    "macro_regime_conditioned",
    "event_driven",
    "options_flow_driven",
    "options_directional",
    "options_volatility",
    "microstructure_imbalance",
    "meta_label_classifier",
}


def _build_feature_dict(n: int = 32) -> dict[str, np.ndarray]:
    idx = np.linspace(-1.0, 1.0, n)
    return {
        "returns_1m": idx,
        "returns_3m": idx**2,
        "returns_6m": np.sin(idx),
        "zscore_5d": idx,
        "distance_from_vwap": np.cos(idx),
        "short_term_reversal": -idx,
        "implied_vol": 0.2 + idx * 0.03,
        "realized_vol": 0.18 + idx * 0.02,
        "iv_rv_spread": 0.02 + idx * 0.01,
        "front_month_iv": 0.21 + idx * 0.03,
        "back_month_iv": 0.19 + idx * 0.02,
        "term_slope": idx * 0.01,
        "index_iv": 0.2 + idx * 0.04,
        "single_name_iv": 0.24 + idx * 0.03,
        "corr_risk_premium": idx * 0.02,
        "pair_spread_z": idx,
        "pair_half_life": 5 + idx,
        "pair_cointegration_score": 0.5 + idx * 0.1,
        "residual_momentum_rank": idx,
        "size_neutral_rank": np.tanh(idx),
        "value_neutral_rank": np.tanh(idx * 0.5),
        "inflation_surprise": idx * 0.2,
        "growth_surprise": idx * 0.15,
        "liquidity_regime": idx,
        "earnings_surprise": idx * 0.3,
        "event_sentiment": np.sin(idx * 2.0),
        "post_event_drift": idx * 0.25,
        "call_put_flow_ratio": 1.0 + idx * 0.1,
        "large_trade_intensity": 0.5 + idx * 0.2,
        "dealer_gamma_exposure": idx * 0.3,

        "skew_z": idx * 0.8,
        "put_call_flow_imbalance_z": idx * 0.5,
        "dealer_positioning_proxy_z": -idx * 0.3,
        "gamma_exposure_proxy_rank": np.linspace(0.0, 1.0, n),
        "unusual_volume_signature_z": np.sin(idx),
        "convexity_z": idx * 0.2,
        "term_structure_curvature_z": -idx * 0.4,
        "local_surface_distortion_z": np.abs(idx),
        "oi_changes_z": idx * 0.1,
        "unusual_volume_signature_rank": np.linspace(1.0, 0.0, n),
        "order_book_imbalance": idx,
        "trade_imbalance": idx * 0.5,
        "quote_slope": idx * 0.2,
        "base_signal": np.sign(idx),
        "base_confidence": np.abs(idx),
        "risk_filter_score": 0.5 + idx * 0.1,
    }


def test_registry_contains_all_required_paradigms() -> None:
    assert ALL_PARADIGMS.issubset(set(MODEL_REGISTRY))


def test_config_activation_weighted_and_stacking_ensemble() -> None:
    config = ModelActivationConfig.from_dict(
        {
            "paradigms": [
                {"name": "momentum", "enabled": True, "weight": 2.0},
                {"name": "mean_reversion", "enabled": True, "weight": 1.0},
                {"name": "event_driven", "enabled": False, "weight": 1.0},
            ]
        }
    )
    active = activated_models(config)
    assert len(active) == 2

    features = _build_feature_dict()
    labels = (features["returns_1m"] > 0).astype(float)

    ensembler = ModelEnsembler(active)
    ensembler.fit(features, labels)
    voted = ensembler.weighted_vote(features)

    assert voted.signal.shape == labels.shape
    assert voted.probability.shape == labels.shape
    assert voted.confidence_scores.shape == labels.shape
    assert set(voted.feature_importances.keys()) == {"momentum", "mean_reversion"}
    for model_importances in voted.feature_importances.values():
        assert model_importances

    ensembler.fit_stacking(features, labels)
    stacked = ensembler.stacking_predict(features)
    assert stacked.signal.shape == labels.shape
    assert np.all(stacked.probability >= 0.0)
    assert np.all(stacked.probability <= 1.0)

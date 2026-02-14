from __future__ import annotations

import numpy as np

from .base import BaseParadigmModel
from modeling_nextgen.models.ml.meta_label_conformal import MetaLabelConformalModel


class MomentumModel(BaseParadigmModel):
    name = "momentum"

    def required_feature_names(self) -> tuple[str, ...]:
        return ("returns_1m", "returns_3m", "returns_6m")


class MeanReversionModel(BaseParadigmModel):
    name = "mean_reversion"

    def required_feature_names(self) -> tuple[str, ...]:
        return ("zscore_5d", "distance_from_vwap", "short_term_reversal")


class VolatilityCarryModel(BaseParadigmModel):
    name = "volatility_carry"

    def required_feature_names(self) -> tuple[str, ...]:
        return ("implied_vol", "realized_vol", "iv_rv_spread")


class TermStructureSlopeModel(BaseParadigmModel):
    name = "term_structure_slope"

    def required_feature_names(self) -> tuple[str, ...]:
        return ("front_month_iv", "back_month_iv", "term_slope")


class DispersionModel(BaseParadigmModel):
    name = "dispersion"

    def required_feature_names(self) -> tuple[str, ...]:
        return ("index_iv", "single_name_iv", "corr_risk_premium")


class StatArbPairSpreadModel(BaseParadigmModel):
    name = "stat_arb_pair_spread"

    def required_feature_names(self) -> tuple[str, ...]:
        return ("pair_spread_z", "pair_half_life", "pair_cointegration_score")


class FactorNeutralCrossSectionalRankModel(BaseParadigmModel):
    name = "factor_neutral_cross_sectional_rank"

    def required_feature_names(self) -> tuple[str, ...]:
        return ("residual_momentum_rank", "size_neutral_rank", "value_neutral_rank")


class MacroRegimeConditionedModel(BaseParadigmModel):
    name = "macro_regime_conditioned"

    def required_feature_names(self) -> tuple[str, ...]:
        return ("inflation_surprise", "growth_surprise", "liquidity_regime")


class EventDrivenModel(BaseParadigmModel):
    name = "event_driven"

    def required_feature_names(self) -> tuple[str, ...]:
        return ("earnings_surprise", "event_sentiment", "post_event_drift")


class OptionsFlowDrivenModel(BaseParadigmModel):
    name = "options_flow_driven"

    def required_feature_names(self) -> tuple[str, ...]:
        return ("call_put_flow_ratio", "large_trade_intensity", "dealer_gamma_exposure")


class MicrostructureImbalanceModel(BaseParadigmModel):
    name = "microstructure_imbalance"

    def required_feature_names(self) -> tuple[str, ...]:
        return ("order_book_imbalance", "trade_imbalance", "quote_slope")


class MetaLabelClassifierModel(BaseParadigmModel):
    name = "meta_label_classifier"

    def __init__(self) -> None:
        super().__init__()
        self._adapter = MetaLabelConformalModel()

    def required_feature_names(self) -> tuple[str, ...]:
        return ("base_signal", "base_confidence", "risk_filter_score")

    def fit(self, features: dict[str, np.ndarray], labels: np.ndarray) -> None:
        self._adapter.fit(features, labels)
        self.feature_importances_ = dict(self._adapter.feature_importances_)
        probs = self.predict_proba(features)
        self.confidence_scores_ = np.abs(probs - 0.5) * 2.0

    def predict_proba(self, features: dict[str, np.ndarray]) -> np.ndarray:
        probs = self._adapter.predict_proba(features)
        policy = self._adapter.apply_policy(features)
        return np.where(policy.accepted_mask, probs, 0.5)


class OptionsDirectionalModel(BaseParadigmModel):
    name = "options_directional"

    def required_feature_names(self) -> tuple[str, ...]:
        return (
            "skew_z",
            "put_call_flow_imbalance_z",
            "dealer_positioning_proxy_z",
            "gamma_exposure_proxy_rank",
            "unusual_volume_signature_z",
        )


class OptionsVolatilityModel(BaseParadigmModel):
    name = "options_volatility"

    def required_feature_names(self) -> tuple[str, ...]:
        return (
            "convexity_z",
            "term_structure_curvature_z",
            "local_surface_distortion_z",
            "oi_changes_z",
            "unusual_volume_signature_rank",
        )

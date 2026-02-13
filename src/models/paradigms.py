from __future__ import annotations

from .base import BaseParadigmModel


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

    def required_feature_names(self) -> tuple[str, ...]:
        return ("base_signal", "base_confidence", "risk_filter_score")

from __future__ import annotations

from backtesting.regime_builder import SUPPORTED_LEG_FAMILIES
from models.regime_catalog import list_models_for_leg, validate_model_leg_pairing
from models.registry import MODEL_REGISTRY


def test_regime_catalog_covers_all_supported_leg_families() -> None:
    for leg_family in SUPPORTED_LEG_FAMILIES:
        descriptors = list_models_for_leg(leg_family)
        assert descriptors, f"No models configured for leg family: {leg_family}"
        for descriptor in descriptors:
            assert descriptor.model_name in MODEL_REGISTRY
            assert descriptor.hyperparameter_template


def test_validate_model_leg_pairing_rejects_unsupported_pairing() -> None:
    try:
        validate_model_leg_pairing("volatility_risk_premium_selling", "momentum_forecasting")
    except ValueError as exc:
        assert "not allowed for leg" in str(exc)
    else:
        raise AssertionError("Expected validate_model_leg_pairing to reject unsupported pairing")

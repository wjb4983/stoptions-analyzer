from __future__ import annotations

from backtesting.regime_builder import SUPPORTED_LEG_FAMILIES
from models.regime_catalog import get_model_descriptor, list_models_for_leg, validate_model_leg_pairing
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



def test_phase_expansion_families_and_capability_tags_present() -> None:
    phase_1_families = {
        "volatility_clustering",
        "iv_ev_spread_term_structure",
        "self_exciting_event_intensity",
        "vol_surface_calibration",
    }
    phase_2_families = {"cross_asset_macro_conditioned", "meta_label_regime_ensemble"}

    for family in phase_1_families | phase_2_families:
        descriptors = list_models_for_leg(family)
        assert descriptors, f"Expected descriptors for expanded family: {family}"

    assert any(
        "supports_architecture_spec" in descriptor.capability_tags
        for descriptor in list_models_for_leg("cross_asset_macro_conditioned")
    )
    assert any(
        "needs_calibration_spec" in descriptor.capability_tags
        for descriptor in list_models_for_leg("vol_surface_calibration")
    )
    assert any(
        "needs_event_process_spec" in descriptor.capability_tags
        for descriptor in list_models_for_leg("self_exciting_event_intensity")
    )



def test_get_model_descriptor_roundtrip() -> None:
    descriptor = get_model_descriptor("cross_asset_macro_conditioned", "macro_regime_conditioned")
    assert descriptor is not None
    assert descriptor.model_name == "macro_regime_conditioned"
    assert descriptor.catalog_phase == "phase_2"


def test_catalog_registry_consistency_for_all_descriptors() -> None:
    covered_models: set[str] = set()
    for leg_family in SUPPORTED_LEG_FAMILIES:
        descriptors = list_models_for_leg(leg_family)
        for descriptor in descriptors:
            covered_models.add(descriptor.model_name)
            assert descriptor.model_name in MODEL_REGISTRY
            assert descriptor.display_name.strip()
            assert descriptor.hyperparameter_template

    assert {
        "policy_gradient_allocation",
        "dqn_regime_allocation",
        "ppo_regime_policy",
        "temporal_transformer_regime",
        "tcn_regime",
        "lstm_regime",
    }.issubset(covered_models)


def test_validate_model_leg_pairing_accepts_new_rl_dl_models() -> None:
    supported_pairs = (
        ("regime_change_detection", "policy_gradient_allocation"),
        ("regime_change_detection", "dqn_regime_allocation"),
        ("regime_change_detection", "ppo_regime_policy"),
        ("cross_asset_macro_conditioned", "temporal_transformer_regime"),
        ("cross_asset_macro_conditioned", "tcn_regime"),
        ("cross_asset_macro_conditioned", "lstm_regime"),
    )
    for leg_family, model_id in supported_pairs:
        validate_model_leg_pairing(leg_family, model_id)

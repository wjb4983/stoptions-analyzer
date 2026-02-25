from __future__ import annotations

import pytest

from ui.create_regime_page import LEG_CONTROL_GROUPS
from ui.regime_mapping import supported_ui_legs, to_regime_leg_spec


def test_supported_ui_legs_include_regime_change() -> None:
    assert set(supported_ui_legs()) == {
        "Trend Following",
        "Mean Reversion",
        "Volatility Breakout",
        "Regime Change",
        "Volatility Clustering",
        "IV/EV Spread",
        "Event Intensity",
        "Vol Surface",
        "Cross-Asset Macro",
        "Meta-Label Ensemble",
    }


@pytest.mark.parametrize(
    "ui_leg_type, expected_family",
    [
        ("Trend Following", "timeseries_momentum"),
        ("Mean Reversion", "cheap_vol_buying"),
        ("Volatility Breakout", "volatility_risk_premium_selling"),
        ("Regime Change", "regime_change_detection"),
        ("Volatility Clustering", "volatility_clustering"),
        ("IV/EV Spread", "iv_ev_spread_term_structure"),
        ("Event Intensity", "self_exciting_event_intensity"),
        ("Vol Surface", "vol_surface_calibration"),
        ("Cross-Asset Macro", "cross_asset_macro_conditioned"),
        ("Meta-Label Ensemble", "meta_label_regime_ensemble"),
    ],
)
def test_to_regime_leg_spec_maps_ui_leg_to_family(ui_leg_type: str, expected_family: str) -> None:
    mapped = to_regime_leg_spec(
        {
            "name": f"{ui_leg_type} leg",
            "model_type": ui_leg_type,
            "controls": {
                "lookback_days": 42,
                "entry_zscore": 1.7,
                "model_confidence_min": 0.67,
                "max_position_pct": 0.04,
                "max_drawdown_stop": 0.09,
            },
        }
    )

    assert mapped.leg_spec.leg_family == expected_family
    assert mapped.default_model_candidates


def test_to_regime_leg_spec_translates_mean_reversion_knobs() -> None:
    mapped = to_regime_leg_spec(
        {
            "name": "Mean Reversion leg",
            "model_type": "Mean Reversion",
            "controls": {
                "lookback_days": 25,
                "entry_zscore": -0.2,
                "model_confidence_min": 0.72,
                "max_position_pct": 0.03,
                "max_drawdown_stop": 0.07,
            },
        }
    )

    assert mapped.leg_spec.knobs["lookback_days"] == 25.0
    assert mapped.leg_spec.knobs["carry_threshold"] == -0.2
    assert mapped.leg_spec.knobs["vol_filter_max"] == 0.72
    assert mapped.leg_spec.knobs["sizing_cap"] == 0.03
    assert mapped.leg_spec.knobs["stop_loss_pct"] == 0.07


def test_to_regime_leg_spec_rejects_unknown_ui_leg() -> None:
    with pytest.raises(ValueError, match="Unsupported UI leg"):
        to_regime_leg_spec({"name": "x", "model_type": "Unknown", "controls": {}})


def test_to_regime_leg_spec_translates_regime_change_knobs() -> None:
    mapped = to_regime_leg_spec(
        {
            "name": "Regime Change leg",
            "model_type": "Regime Change",
            "controls": {
                "lookback_days": 80,
                "detection_threshold": 2.1,
                "max_position_pct": 0.02,
                "max_drawdown_stop": 0.05,
            },
        }
    )

    assert mapped.leg_spec.leg_family == "regime_change_detection"
    assert mapped.leg_spec.knobs["lookback_days"] == 80.0
    assert mapped.leg_spec.knobs["detection_threshold"] == 2.1
    assert mapped.leg_spec.knobs["sizing_cap"] == 0.02
    assert mapped.leg_spec.knobs["stop_loss_pct"] == 0.05


def test_to_regime_leg_spec_regime_change_accepts_entry_zscore_alias() -> None:
    mapped = to_regime_leg_spec(
        {
            "name": "Regime Change leg",
            "model_type": "Regime Change",
            "controls": {
                "entry_zscore": 1.4,
            },
        }
    )

    assert mapped.leg_spec.knobs["detection_threshold"] == 1.4


def test_to_regime_leg_spec_translates_event_intensity_knobs() -> None:
    mapped = to_regime_leg_spec(
        {
            "name": "Event Intensity leg",
            "model_type": "Event Intensity",
            "controls": {
                "lookback_days": 20,
                "entry_zscore": 1.8,
                "model_confidence_min": 0.66,
                "max_position_pct": 0.025,
                "max_drawdown_stop": 0.06,
            },
        }
    )

    assert mapped.leg_spec.leg_family == "self_exciting_event_intensity"
    assert mapped.leg_spec.knobs["lookback_days"] == 20.0
    assert mapped.leg_spec.knobs["detection_threshold"] == 1.8
    assert mapped.leg_spec.knobs["vol_filter_min"] == 0.66
    assert mapped.leg_spec.knobs["sizing_cap"] == 0.025
    assert mapped.leg_spec.knobs["stop_loss_pct"] == 0.06



def test_supported_ui_legs_have_control_schema_entries() -> None:
    for ui_leg in supported_ui_legs():
        assert ui_leg in LEG_CONTROL_GROUPS



def test_to_regime_leg_spec_translates_cross_asset_macro_knobs() -> None:
    mapped = to_regime_leg_spec(
        {
            "name": "Cross-Asset Macro leg",
            "model_type": "Cross-Asset Macro",
            "controls": {
                "lookback_days": 96,
                "entry_zscore": 1.5,
                "model_confidence_min": 0.61,
                "max_position_pct": 0.03,
                "max_drawdown_stop": 0.07,
            },
        }
    )

    assert mapped.leg_spec.leg_family == "cross_asset_macro_conditioned"
    assert mapped.leg_spec.knobs["lookback_days"] == 96.0
    assert mapped.leg_spec.knobs["detection_threshold"] == 1.5
    assert mapped.leg_spec.knobs["vol_filter_min"] == 0.61



def test_to_regime_leg_spec_translates_meta_label_ensemble_knobs() -> None:
    mapped = to_regime_leg_spec(
        {
            "name": "Meta-Label Ensemble leg",
            "model_type": "Meta-Label Ensemble",
            "controls": {
                "lookback_days": 70,
                "entry_zscore": 1.45,
                "model_confidence_min": 0.73,
            },
        }
    )

    assert mapped.leg_spec.leg_family == "meta_label_regime_ensemble"
    assert mapped.leg_spec.knobs["detection_threshold"] == 1.45
    assert mapped.leg_spec.knobs["vol_filter_max"] == 0.73

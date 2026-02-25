from pathlib import Path

import pytest

from backtesting.regime_builder import (
    ModelChoiceSpec,
    RegimeLegSpec,
    RegimeSpec,
    RiskSpec,
    TrainingSpec,
    load_regime_spec,
    save_regime_spec,
    validate_leg_spec,
    validate_regime_spec,
)


def _base_spec(*, legs: tuple[RegimeLegSpec, ...]) -> RegimeSpec:
    return RegimeSpec(
        regime_name="My Core Regime",
        training=TrainingSpec(
            train_start="2018-01-01",
            train_end="2022-12-31",
            validation_window_days=180,
            retrain_frequency_days=20,
        ),
        risk=RiskSpec(
            max_gross_exposure=2.0,
            max_net_exposure=1.0,
            max_drawdown_pct=0.25,
            default_sizing_cap=0.2,
            default_stop_loss_pct=0.08,
        ),
        model_choice=ModelChoiceSpec(
            model_name="meta_label_classifier",
            objective="sharpe",
            hyperparameters={"states": 4},
        ),
        legs=legs,
    )


def test_validate_leg_spec_accepts_each_supported_family() -> None:
    legs = (
        RegimeLegSpec(
            name="tsmom",
            leg_family="timeseries_momentum",
            knobs={"lookback_days": 120, "vol_filter_max": 0.3, "sizing_cap": 0.15, "stop_loss_pct": 0.06},
        ),
        RegimeLegSpec(
            name="vrp_sell",
            leg_family="volatility_risk_premium_selling",
            knobs={
                "lookback_days": 30,
                "carry_threshold": 0.12,
                "vol_filter_min": 0.18,
                "sizing_cap": 0.12,
                "stop_loss_pct": 0.05,
            },
        ),
        RegimeLegSpec(
            name="cheap_vol",
            leg_family="cheap_vol_buying",
            knobs={
                "lookback_days": 20,
                "carry_threshold": -0.1,
                "vol_filter_max": 0.22,
                "sizing_cap": 0.1,
                "stop_loss_pct": 0.04,
            },
        ),
        RegimeLegSpec(
            name="regime_change",
            leg_family="regime_change_detection",
            knobs={"lookback_days": 40, "detection_threshold": 1.25, "sizing_cap": 0.08, "stop_loss_pct": 0.03},
        ),
    )

    for leg in legs:
        validate_leg_spec(leg)

    validate_regime_spec(_base_spec(legs=legs))


def test_validate_leg_spec_rejects_missing_required_knob() -> None:
    leg = RegimeLegSpec(
        name="bad",
        leg_family="regime_change_detection",
        knobs={"lookback_days": 30, "sizing_cap": 0.1, "stop_loss_pct": 0.04},
    )

    with pytest.raises(ValueError, match="Missing required knobs"):
        validate_leg_spec(leg)


def test_validate_leg_spec_rejects_invalid_numeric_ranges() -> None:
    leg = RegimeLegSpec(
        name="bad_range",
        leg_family="cheap_vol_buying",
        knobs={
            "lookback_days": 30,
            "carry_threshold": 1.2,
            "vol_filter_max": 0.25,
            "sizing_cap": 1.1,
            "stop_loss_pct": 0.03,
        },
    )

    with pytest.raises(ValueError, match="carry_threshold"):
        validate_leg_spec(leg)


def test_save_load_roundtrip_keeps_schema_compatible(tmp_path: Path) -> None:
    legs = (
        RegimeLegSpec(
            name="tsmom",
            leg_family="timeseries_momentum",
            knobs={"lookback_days": 120, "vol_filter_max": 0.3, "sizing_cap": 0.15, "stop_loss_pct": 0.06},
        ),
    )
    spec = _base_spec(legs=legs)

    saved = save_regime_spec(spec, output_dir=tmp_path)
    loaded_by_name = load_regime_spec("My Core Regime", output_dir=tmp_path)
    loaded_by_path = load_regime_spec(saved)

    assert saved.name == "my_core_regime.json"
    assert loaded_by_name == spec
    assert loaded_by_path == spec


def test_validate_regime_spec_rejects_invalid_model_leg_pairing() -> None:
    spec = _base_spec(
        legs=(
            RegimeLegSpec(
                name="tsmom",
                leg_family="timeseries_momentum",
                knobs={"lookback_days": 120, "vol_filter_max": 0.3, "sizing_cap": 0.15, "stop_loss_pct": 0.06},
            ),
        )
    )
    spec = RegimeSpec(
        regime_name=spec.regime_name,
        training=spec.training,
        risk=spec.risk,
        model_choice=ModelChoiceSpec(model_name="hmm_regime_change", objective="sharpe", hyperparameters={}),
        legs=spec.legs,
        schema_version=spec.schema_version,
    )

    with pytest.raises(ValueError, match="not allowed for leg"):
        validate_regime_spec(spec)

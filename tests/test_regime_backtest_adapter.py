from __future__ import annotations

import json
from pathlib import Path

import pytest

from src.backtesting.regime_backtest_adapter import (
    RegimeBacktestOption,
    RegimeBundleCompatibilityError,
    discover_regime_backtest_options,
    load_regime_backtest_contract,
)


def _write_json(path: Path, payload: dict) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, sort_keys=True), encoding="utf-8")
    return path


def _training_manifest(run_id: str = "run-1") -> dict:
    return {
        "manifest_schema_version": "2.1.0",
        "manifest_schema_min_reader_version": "2.0.0",
        "run_id": run_id,
        "request": {
            "regime_name": "Risk On",
            "model_choice": "auto_model_search",
            "training_window": {"lookback_days": 144, "retrain_frequency_days": 8},
            "legs": [
                {"name": "Trend", "controls": {"feature_columns": ["ret_5", "vol_20"]}},
                {"name": "Carry", "controls": {"feature_columns": ["carry_signal"]}},
            ],
            "risk_limits": {
                "max_gross_exposure": 1.25,
                "max_net_exposure": 0.45,
                "max_position_weight": 0.11,
                "max_sector_weight": 0.33,
                "confidence_min_assignment_confidence": 0.62,
                "confidence_alert_confidence": 0.74,
                "confidence_min_transition_confidence": 0.57,
            },
            "training_data_settings": {
                "scenario_settings": [{"name": "panic_crash"}, {"name": "bull_low_vol"}],
            },
        },
        "metadata": {"champion_by_leg": {"Trend": "meta_label_classifier", "Carry": "ppo_regime_policy"}},
        "artifact_paths": {
            "trend_model_weights": "artifacts/trend/model.pkl",
            "trend_calibration_object": "artifacts/trend/calibration.json",
            "carry_model_weights": "artifacts/carry/model.pkl",
            "carry_calibration_object": "artifacts/carry/calibration.json"
        },
    }


def test_discover_populates_from_runs_and_exports(tmp_path: Path) -> None:
    training_manifest = _write_json(tmp_path / "run" / "manifest.json", _training_manifest())
    bundle_manifest = _write_json(
        tmp_path / "exports" / "bundle-1" / "bundle_manifest.json",
        {
            "bundle_id": "bundle-1",
            "manifest_schema_version": "1.1.0",
            "manifest_schema_min_reader_version": "1.0.0",
            "bundle_version": "1.1.0",
            "run_id": "run-2",
            "contents": {"training_manifest": str(training_manifest)},
        },
    )

    options = discover_regime_backtest_options(
        [{"run_id": "run-1", "artifact_path": str(training_manifest), "summary": "Risk On"}],
        regime_exports_root=tmp_path / "exports",
    )

    labels = [item.label for item in options]
    assert any("training run" in label for label in labels)
    assert any("export bundle" in label for label in labels)
    assert any(item.manifest_path == str(bundle_manifest) for item in options)


def test_contract_hydration_maps_expected_defaults(tmp_path: Path) -> None:
    manifest_path = _write_json(tmp_path / "run" / "manifest.json", _training_manifest("run-hydrate"))
    option = RegimeBacktestOption(
        option_id="training:run-hydrate",
        label="hydrate",
        source="training_run",
        manifest_path=str(manifest_path),
    )

    contract = load_regime_backtest_contract(option)

    assert contract.defaults["strategy"] == "momentum"
    assert contract.defaults["lookback_days"] == "144"
    assert contract.defaults["skip_days"] == "8"
    assert contract.defaults["portfolio_max_gross_exposure"] == "1.25"
    assert contract.defaults["selected_scenario_packs"] == "panic_crash,bull_low_vol"
    assert contract.defaults["hydration_schema_version"] == "1.1.0"
    assert contract.execution_artifacts["champion_model_ids"]["Trend"] == "meta_label_classifier"
    assert contract.execution_artifacts["model_paths"]["trend"] == "artifacts/trend/model.pkl"
    assert contract.execution_artifacts["calibration_paths"]["carry"] == "artifacts/carry/calibration.json"
    assert contract.execution_artifacts["feature_expectations"]["Trend"] == ["ret_5", "vol_20"]


def test_invalid_or_missing_bundle_manifest_raises_actionable_error(tmp_path: Path) -> None:
    missing_option = RegimeBacktestOption(
        option_id="bundle:missing",
        label="missing",
        source="bundle",
        manifest_path=str(tmp_path / "missing" / "bundle_manifest.json"),
    )
    with pytest.raises(RegimeBundleCompatibilityError, match="Manifest not found"):
        load_regime_backtest_contract(missing_option)

    invalid_manifest_path = tmp_path / "exports" / "bad" / "bundle_manifest.json"
    invalid_manifest_path.parent.mkdir(parents=True, exist_ok=True)
    invalid_manifest_path.write_text("{not valid json", encoding="utf-8")
    invalid_option = RegimeBacktestOption(
        option_id="bundle:bad",
        label="bad",
        source="bundle",
        manifest_path=str(invalid_manifest_path),
    )
    with pytest.raises(RegimeBundleCompatibilityError, match="invalid JSON"):
        load_regime_backtest_contract(invalid_option)


def test_contract_hydration_supports_legacy_n_minus_one_training_manifest(tmp_path: Path) -> None:
    legacy_manifest = _write_json(
        tmp_path / "run" / "legacy_manifest.json",
        {
            "run_id": "legacy-run",
            "request": {
                "schema_version": 2,
                "regime_name": "Legacy Risk On",
                "model_choice": "auto_model_search",
                "training_window": {"lookback_days": 120, "retrain_frequency_days": 6},
                "risk_limits": {},
            },
        },
    )
    option = RegimeBacktestOption(
        option_id="training:legacy-run",
        label="legacy",
        source="training_run",
        manifest_path=str(legacy_manifest),
    )

    contract = load_regime_backtest_contract(option)
    assert contract.regime_name == "Legacy Risk On"


def test_bundle_rejects_incompatible_schema_version(tmp_path: Path) -> None:
    training_manifest = _write_json(tmp_path / "run" / "manifest.json", _training_manifest())
    bad_bundle_manifest = _write_json(
        tmp_path / "exports" / "bundle-bad" / "bundle_manifest.json",
        {
            "bundle_id": "bundle-bad",
            "manifest_schema_version": "2.0.0",
            "bundle_version": "2.0.0",
            "run_id": "run-2",
            "contents": {"training_manifest": str(training_manifest)},
        },
    )
    option = RegimeBacktestOption(
        option_id="bundle:bundle-bad",
        label="bad bundle",
        source="bundle",
        manifest_path=str(bad_bundle_manifest),
    )

    with pytest.raises(RegimeBundleCompatibilityError, match="not supported"):
        load_regime_backtest_contract(option)

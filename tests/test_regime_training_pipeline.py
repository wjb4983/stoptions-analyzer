from __future__ import annotations

import json
from datetime import datetime, timedelta, timezone
from pathlib import Path

import numpy as np

from backtesting.regime_training_pipeline import (
    AdapterIssue,
    LegOutOfSampleMetrics,
    RegimeLegTrainingConfig,
    RegimeTrainingAdapterOutput,
    RegimeTrainingRequest,
    TrainedArtifactLocations,
    _load_returns_from_cache,
    RegistryBackedRegimeTrainingAdapter,
    execute_regime_training_pipeline,
)


class _FailingAdapter:
    def fit_and_backtest(self, request: RegimeTrainingRequest, run_dir):  # noqa: ANN001
        raise RuntimeError("adapter exploded")


class _StableAdapter:
    def fit_and_backtest(self, request: RegimeTrainingRequest, run_dir):  # noqa: ANN001
        return RegimeTrainingAdapterOutput(
            per_leg_artifacts={
                "Trend": TrainedArtifactLocations(
                    model_weights=str(run_dir / "trend_model_weights.json"),
                    calibration_object=str(run_dir / "trend_calibration.json"),
                    diagnostics=str(run_dir / "trend_diagnostics.json"),
                ),
            },
            per_leg_oos_metrics={
                "Trend": LegOutOfSampleMetrics(
                    leg_name="Trend",
                    model_id="momentum",
                    metrics={"accuracy": 0.68, "brier_score": 0.19},
                ),
            },
            portfolio_oos_metrics={"portfolio_avg_accuracy": 0.68, "portfolio_avg_brier_score": 0.19},
            issues=(
                AdapterIssue(
                    level="warning",
                    model_id="momentum",
                    message="insufficient history in one bucket",
                ),
            ),
            adapter_name="stable_test_adapter",
        )


def _request(tmp_path, *, legs: tuple[RegimeLegTrainingConfig, ...]) -> RegimeTrainingRequest:
    return RegimeTrainingRequest(
        schema_version=2,
        regime_id="risk_on",
        regime_name="Risk On",
        legs=legs,
        model_choice="hmm_v1",
        training_window={"retrain_frequency_days": 21, "lookback_days": 252},
        risk_limits={"max_drawdown_stop": 0.15, "max_position_pct": 0.08},
        output_dir=str(tmp_path),
        training_data_settings={"allow_synthetic_fallback": True},
    )


def test_pipeline_success_writes_manifest_and_returns_artifacts(tmp_path):
    req = _request(
        tmp_path,
        legs=(
            RegimeLegTrainingConfig(
                name="Trend",
                model_type="timeseries_momentum",
                controls={"model_confidence_min": 0.7, "turnover_limit": 0.2},
            ),
        ),
    )

    result = execute_regime_training_pipeline(req, adapter=_StableAdapter())

    assert result.status == "success"
    assert result.metrics["portfolio_avg_accuracy"] == 0.68
    assert result.warnings == ("[momentum] insufficient history in one bucket",)
    assert "trend_model_weights" in result.artifact_paths
    assert "manifest" in result.artifact_paths

    manifest_path = result.artifact_paths["manifest"]
    payload = json.loads(open(manifest_path, encoding="utf-8").read())
    assert payload["status"] == "success"
    assert payload["manifest_schema_version"] == "2.1.0"
    assert payload["manifest_schema_min_reader_version"] == "2.0.0"
    assert payload["request"]["regime_name"] == "Risk On"
    assert payload["artifact_paths"]["spec"].endswith("regime_spec_snapshot.json")
    assert payload["metadata"]["candidate_leaderboard"] == {}
    assert payload["metadata"]["champion_by_leg"] == {}
    reproducibility = payload["metadata"]["reproducibility"]
    assert reproducibility["request_checksum"]
    leg_fingerprint = reproducibility["legs"]["00:Trend"]
    assert leg_fingerprint["hyperparameters_checksum"]
    assert leg_fingerprint["architecture_spec_checksum"]


def test_pipeline_validation_failure_returns_machine_readable_error(tmp_path):
    req = _request(tmp_path, legs=tuple())

    result = execute_regime_training_pipeline(req)

    assert result.status == "failed"
    assert result.error_payload is not None
    assert result.error_payload["code"] == "INVALID_REGIME_SPEC"
    assert result.error_payload["stage"] == "validate_regime_spec"
    assert any("at least one leg is required" in err for err in result.errors)


def test_pipeline_adapter_failure_returns_machine_readable_error(tmp_path):
    req = _request(
        tmp_path,
        legs=(
            RegimeLegTrainingConfig(
                name="Carry",
                model_type="volatility_risk_premium_selling",
                controls={"model_confidence_min": 0.6},
            ),
        ),
    )

    result = execute_regime_training_pipeline(req, adapter=_FailingAdapter())

    assert result.status == "failed"
    assert result.error_payload is not None
    assert result.error_payload["code"] == "TRAINING_EXECUTION_FAILED"
    assert result.error_payload["stage"] == "fit_and_backtest"
    assert result.error_payload["exception_type"] == "RuntimeError"


def test_pipeline_default_adapter_runs_registry_backed_training(tmp_path):
    req = RegimeTrainingRequest(
        schema_version=2,
        regime_id="risk_off",
        regime_name="Risk Off",
        legs=(
            RegimeLegTrainingConfig(
                name="Regime Detection",
                model_type="regime_change_detection",
                controls={"model_confidence_min": 0.62, "turnover_limit": 0.25},
            ),
        ),
        model_choice="auto",
        training_window={"retrain_frequency_days": 14, "lookback_days": 252},
        risk_limits={"max_drawdown_stop": 0.15, "max_position_pct": 0.08},
        output_dir=str(tmp_path),
        training_data_settings={"allow_synthetic_fallback": True},
    )

    result = execute_regime_training_pipeline(req)

    assert result.status == "success"
    assert result.metrics["legs_trained"] == 1.0
    assert any(key.endswith("_model_weights") for key in result.artifact_paths)
    assert "Regime Detection" in result.metadata["candidate_leaderboard"]
    assert result.metadata["champion_by_leg"]["Regime Detection"]
    governance = result.metadata["governance_by_leg"]["Regime Detection"]
    assert "pass_fail" in governance
    assert "deployment_slot_eligibility" in governance


def test_pipeline_validation_rejects_incomplete_model_specific_specs(tmp_path):
    req = _request(
        tmp_path,
        legs=(
            RegimeLegTrainingConfig(
                name="Vol Surface",
                model_type="vol_surface_calibration",
                controls={"model_confidence_min": 0.6},
                model_id="options_volatility",
            ),
            RegimeLegTrainingConfig(
                name="Event Intensity",
                model_type="self_exciting_event_intensity",
                controls={"model_confidence_min": 0.6},
                model_id="event_driven",
            ),
        ),
    )

    result = execute_regime_training_pipeline(req)

    assert result.status == "failed"
    assert any("legs[0].calibration_spec" in err for err in result.errors)
    assert any("legs[1].event_process_spec" in err for err in result.errors)


def test_pipeline_validation_accepts_complete_model_specific_specs(tmp_path):
    req = _request(
        tmp_path,
        legs=(
            RegimeLegTrainingConfig(
                name="Cross Asset Macro",
                model_type="cross_asset_macro_conditioned",
                controls={"model_confidence_min": 0.6},
                model_id="macro_regime_conditioned",
                architecture_spec={"schema_version": 1, "layers": [{"type": "Dense", "units": 16, "activation": "relu"}], "optimizer": {"name": "adam", "learning_rate": 0.001}, "loss": {"name": "binary_cross_entropy"}, "scheduler": {"name": "none"}, "training": {"batch_size": 16, "epochs": 8, "early_stopping": {"enabled": True, "patience": 3}}},
            ),
            RegimeLegTrainingConfig(
                name="Vol Surface",
                model_type="vol_surface_calibration",
                controls={"model_confidence_min": 0.6},
                model_id="options_volatility",
                calibration_spec={"model": "heston", "parameters": {"kappa": 1.2}},
            ),
            RegimeLegTrainingConfig(
                name="Event Intensity",
                model_type="self_exciting_event_intensity",
                controls={"model_confidence_min": 0.6},
                model_id="event_driven",
                event_process_spec={"process_type": "hawkes", "parameters": {"alpha": 0.3}},
            ),
        ),
    )

    result = execute_regime_training_pipeline(req)

    assert result.status == "success"


def test_pipeline_persists_cache_audit_report_into_run_dir(tmp_path):
    audit_path = tmp_path / "cache_audit_input.json"
    audit_path.write_text(json.dumps({"pass": False, "failing_symbols": ["MSFT"]}), encoding="utf-8")

    req = _request(
        tmp_path,
        legs=(
            RegimeLegTrainingConfig(
                name="Trend",
                model_type="timeseries_momentum",
                controls={"model_confidence_min": 0.7, "turnover_limit": 0.2},
            ),
        ),
    )
    req.training_data_settings["cache_audit_report"] = str(audit_path)

    result = execute_regime_training_pipeline(req, adapter=_StableAdapter())

    assert "cache_audit_report" in result.artifact_paths
    copied = Path(result.artifact_paths["cache_audit_report"])
    assert copied.exists()
    assert copied.name == "cache_audit_report.json"


def _write_symbol_cache(symbol_root: Path, symbol: str, start: datetime, days: int) -> None:
    safe = symbol
    bucket = symbol_root / safe / "1m"
    bucket.mkdir(parents=True, exist_ok=True)
    timestamps = np.array([int((start + timedelta(days=idx)).timestamp() * 1000) for idx in range(days)], dtype=np.int64)
    close = np.linspace(100.0, 105.0, num=days, dtype=float)
    year = start.year
    path = bucket / f"{safe}_1m_{year}.npz"
    np.savez_compressed(path, t=timestamps, c=close)


def test_load_returns_from_cache_exactly_five_year_history(tmp_path):
    now = datetime(2026, 1, 1, tzinfo=timezone.utc)
    start = now - timedelta(days=365 * 5)
    _write_symbol_cache(tmp_path, "AAPL", start, days=365 * 5 + 3)

    bundle = _load_returns_from_cache(
        symbols=["AAPL"],
        cache_root=tmp_path,
        min_usable_history_years=5,
        required_universe_pass_ratio=1.0,
        now=now,
    )

    assert bundle.metadata["symbols_used"] == ["AAPL"]
    assert bundle.metadata["symbols_excluded"] == []
    assert bundle.metadata["pass"] is True
    assert bundle.metadata["effective_training_start"] is not None
    assert bundle.metadata["per_symbol"]["AAPL"]["years_used"] >= 4.99


def test_load_returns_from_cache_uses_deeper_than_five_year_history(tmp_path):
    now = datetime(2026, 1, 1, tzinfo=timezone.utc)
    start = now - timedelta(days=365 * 7)
    _write_symbol_cache(tmp_path, "AAPL", start, days=365 * 7 + 4)

    bundle = _load_returns_from_cache(
        symbols=["AAPL"],
        cache_root=tmp_path,
        min_usable_history_years=5,
        required_universe_pass_ratio=1.0,
        now=now,
    )

    expected_start = start.isoformat()
    assert bundle.metadata["per_symbol"]["AAPL"]["effective_start"] == expected_start
    assert bundle.metadata["per_symbol"]["AAPL"]["years_used"] > 5.0


def test_load_returns_from_cache_mixed_coverage_universe_enforces_pass_ratio(tmp_path):
    now = datetime(2026, 1, 1, tzinfo=timezone.utc)
    _write_symbol_cache(tmp_path, "AAPL", now - timedelta(days=365 * 6), days=365 * 6 + 3)
    _write_symbol_cache(tmp_path, "MSFT", now - timedelta(days=200), days=220)

    bundle = _load_returns_from_cache(
        symbols=["AAPL", "MSFT", "TSLA"],
        cache_root=tmp_path,
        min_usable_history_years=5,
        required_universe_pass_ratio=0.6,
        now=now,
    )

    assert bundle.metadata["symbols_requested"] == ["AAPL", "MSFT", "TSLA"]
    assert bundle.metadata["symbols_used"] == ["AAPL"]
    excluded = {row["symbol"]: row["reason"] for row in bundle.metadata["symbols_excluded"]}
    assert excluded["MSFT"] == "insufficient_usable_history"
    assert excluded["TSLA"] == "missing_or_unreadable_cache"
    assert bundle.metadata["universe_pass_ratio"] == 1 / 3
    assert bundle.metadata["pass"] is False




def test_pipeline_fails_when_synthetic_fallback_disabled_and_real_history_insufficient(tmp_path):
    req = RegimeTrainingRequest(
        schema_version=2,
        regime_id="risk_off",
        regime_name="Risk Off",
        legs=(
            RegimeLegTrainingConfig(
                name="Regime Detection",
                model_type="regime_change_detection",
                controls={"model_confidence_min": 0.62, "turnover_limit": 0.25},
            ),
        ),
        model_choice="auto",
        training_window={"retrain_frequency_days": 14, "lookback_days": 252},
        risk_limits={"max_drawdown_stop": 0.15, "max_position_pct": 0.08},
        output_dir=str(tmp_path),
        training_data_settings={
            "universe_symbols": [],
            "allow_synthetic_fallback": False,
        },
    )

    result = execute_regime_training_pipeline(req)

    assert result.status == "failed"
    assert any("INSUFFICIENT_REAL_HISTORY" in err for err in result.errors)
    assert any("coverage_diagnostics=" in err for err in result.errors)


def test_pipeline_uses_synthetic_fallback_when_explicitly_enabled(tmp_path):
    req = RegimeTrainingRequest(
        schema_version=2,
        regime_id="risk_off",
        regime_name="Risk Off",
        legs=(
            RegimeLegTrainingConfig(
                name="Regime Detection",
                model_type="regime_change_detection",
                controls={"model_confidence_min": 0.62, "turnover_limit": 0.25},
            ),
        ),
        model_choice="auto",
        training_window={"retrain_frequency_days": 14, "lookback_days": 252},
        risk_limits={"max_drawdown_stop": 0.15, "max_position_pct": 0.08},
        output_dir=str(tmp_path),
        training_data_settings={
            "universe_symbols": [],
            "allow_synthetic_fallback": True,
        },
    )

    result = execute_regime_training_pipeline(req)

    assert result.status == "success"
    assert result.metadata["synthetic_fallback_used"] is True
    assert any("synthetic fallback" in warning for warning in result.warnings)
    assert "synthetic_fallback_used=true" in result.logs

def test_pipeline_manifest_contains_training_data_audit(tmp_path):
    now = datetime.now(timezone.utc)
    cache_root = tmp_path / "cache"
    _write_symbol_cache(cache_root, "AAPL", now - timedelta(days=365 * 6), days=365 * 6 + 3)

    req = RegimeTrainingRequest(
        schema_version=2,
        regime_id="risk_off",
        regime_name="Risk Off",
        legs=(
            RegimeLegTrainingConfig(
                name="Regime Detection",
                model_type="regime_change_detection",
                controls={"model_confidence_min": 0.62, "turnover_limit": 0.25},
            ),
        ),
        model_choice="auto",
        training_window={"retrain_frequency_days": 14, "lookback_days": 252},
        risk_limits={"max_drawdown_stop": 0.15, "max_position_pct": 0.08},
        output_dir=str(tmp_path),
        training_data_settings={
            "universe_symbols": ["AAPL"],
            "required_history_years": 5,
            "cache_root": str(cache_root),
            "required_universe_pass_ratio": 1.0,
        },
    )

    result = execute_regime_training_pipeline(req)

    manifest = json.loads(Path(result.artifact_paths["manifest"]).read_text(encoding="utf-8"))
    audit = manifest["metadata"]["training_data_audit"]
    assert audit["symbols_requested"] == ["AAPL"]
    assert audit["symbols_used"] == ["AAPL"]
    assert "effective_training_start" in audit



def test_pipeline_continues_with_degraded_universe_coverage(tmp_path):
    now = datetime(2026, 1, 1, tzinfo=timezone.utc)
    cache_root = tmp_path / "cache"
    _write_symbol_cache(cache_root, "AAPL", now - timedelta(days=365 * 6), days=365 * 6 + 3)

    req = RegimeTrainingRequest(
        schema_version=2,
        regime_id="risk_off",
        regime_name="Risk Off",
        legs=(
            RegimeLegTrainingConfig(
                name="Regime Detection",
                model_type="regime_change_detection",
                controls={"model_confidence_min": 0.62, "turnover_limit": 0.25},
            ),
        ),
        model_choice="auto",
        training_window={"retrain_frequency_days": 14, "lookback_days": 252},
        risk_limits={"max_drawdown_stop": 0.15, "max_position_pct": 0.08},
        output_dir=str(tmp_path),
        training_data_settings={
            "universe_symbols": ["AAPL", "MSFT"],
            "required_history_years": 5,
            "cache_root": str(cache_root),
            "required_universe_pass_ratio": 1.0,
        },
    )

    result = execute_regime_training_pipeline(req)

    assert result.status == "success"
    assert result.metadata["training_data_audit"]["pass"] is False
    assert result.metadata["training_data_audit"]["degraded_universe_coverage"] is True
    assert any("continuing with available symbols" in warning for warning in result.warnings)


def test_pipeline_phase2_families_compatibility_with_default_adapter(tmp_path):
    req = RegimeTrainingRequest(
        schema_version=2,
        regime_id="phase2",
        regime_name="Phase 2",
        legs=(
            RegimeLegTrainingConfig(
                name="Cross-Asset Macro",
                model_type="cross_asset_macro_conditioned",
                controls={"model_confidence_min": 0.62, "turnover_limit": 0.2},
            ),
            RegimeLegTrainingConfig(
                name="Meta Ensemble",
                model_type="meta_label_regime_ensemble",
                controls={"model_confidence_min": 0.65, "turnover_limit": 0.2},
            ),
        ),
        model_choice="auto",
        training_window={"retrain_frequency_days": 14, "lookback_days": 252},
        risk_limits={"max_drawdown_stop": 0.15, "max_position_pct": 0.08},
        output_dir=str(tmp_path),
        training_data_settings={"allow_synthetic_fallback": True},
    )

    result = execute_regime_training_pipeline(req)

    assert result.status == "success"
    assert result.metrics["legs_trained"] == 2.0
    assert "Cross-Asset Macro" in result.metadata["champion_by_leg"]
    assert "Meta Ensemble" in result.metadata["champion_by_leg"]


def test_pipeline_validation_rejects_invalid_architecture_schema(tmp_path):
    req = _request(
        tmp_path,
        legs=(
            RegimeLegTrainingConfig(
                name="Cross Asset Macro",
                model_type="cross_asset_macro_conditioned",
                controls={"model_confidence_min": 0.6},
                model_id="macro_regime_conditioned",
                architecture_spec={
                    "schema_version": 1,
                    "layers": [{"type": "Dense", "units": 0, "activation": "relu"}],
                    "optimizer": {"name": "adam", "learning_rate": 0.0},
                    "loss": {"name": "binary_cross_entropy"},
                    "scheduler": {"name": "none"},
                    "training": {
                        "batch_size": 0,
                        "epochs": 0,
                        "early_stopping": {"enabled": True, "patience": 0},
                    },
                },
            ),
        ),
    )

    result = execute_regime_training_pipeline(req)

    assert result.status == "failed"
    assert any("architecture_spec.layers[0].units" in err for err in result.errors)
    assert any("architecture_spec.optimizer.learning_rate" in err for err in result.errors)
    assert any("architecture_spec.training.batch_size" in err for err in result.errors)


def test_pipeline_consumes_architecture_spec_for_supported_models(tmp_path):
    req = _request(
        tmp_path,
        legs=(
            RegimeLegTrainingConfig(
                name="Cross Asset Macro",
                model_type="cross_asset_macro_conditioned",
                controls={"model_confidence_min": 0.6},
                model_id="macro_regime_conditioned",
                architecture_spec={
                    "schema_version": 1,
                    "layers": [{"type": "Dense", "units": 16, "activation": "relu"}],
                    "optimizer": {"name": "adam", "learning_rate": 0.001},
                    "loss": {"name": "binary_cross_entropy"},
                    "scheduler": {"name": "none"},
                    "training": {
                        "batch_size": 16,
                        "epochs": 8,
                        "early_stopping": {"enabled": True, "patience": 3},
                    },
                },
            ),
        ),
    )

    result = execute_regime_training_pipeline(req)

    assert result.status == "success"
    diagnostics_key = next(key for key in result.artifact_paths if key.endswith("_diagnostics"))
    diagnostics_payload = json.loads(Path(result.artifact_paths[diagnostics_key]).read_text(encoding="utf-8"))
    assert diagnostics_payload["architecture_spec"]["optimizer"]["name"] == "adam"


def test_candidate_selection_modes_include_new_model_ids() -> None:
    leg = RegimeLegTrainingConfig(
        name="Regime Detection",
        model_type="regime_change_detection",
        controls={"model_confidence_min": 0.6},
        model_id="ppo_regime_policy",
    )

    req_single = RegimeTrainingRequest(
        schema_version=2,
        regime_id="mode-single",
        regime_name="Mode Single",
        legs=(leg,),
        model_choice="single_model",
        training_window={"retrain_frequency_days": 14, "lookback_days": 252},
        risk_limits={"max_drawdown_stop": 0.15, "max_position_pct": 0.08},
        output_dir=".",
        training_data_settings={"allow_synthetic_fallback": True},
    )
    assert RegistryBackedRegimeTrainingAdapter._candidate_model_ids(req_single, leg) == ["ppo_regime_policy"]

    req_auto = RegimeTrainingRequest(
        schema_version=2,
        regime_id="mode-auto",
        regime_name="Mode Auto",
        legs=(leg,),
        model_choice="auto",
        training_window={"retrain_frequency_days": 14, "lookback_days": 252},
        risk_limits={"max_drawdown_stop": 0.15, "max_position_pct": 0.08},
        output_dir=".",
        training_data_settings={"allow_synthetic_fallback": True},
    )
    auto_candidates = RegistryBackedRegimeTrainingAdapter._candidate_model_ids(req_auto, leg)
    assert "policy_gradient_allocation" in auto_candidates
    assert "dqn_regime_allocation" in auto_candidates
    assert "ppo_regime_policy" in auto_candidates

    req_ensemble = RegimeTrainingRequest(
        schema_version=2,
        regime_id="mode-ensemble",
        regime_name="Mode Ensemble",
        legs=(leg,),
        model_choice="ensemble",
        training_window={"retrain_frequency_days": 14, "lookback_days": 252},
        risk_limits={"max_drawdown_stop": 0.15, "max_position_pct": 0.08},
        output_dir=".",
        training_data_settings={"allow_synthetic_fallback": True},
    )
    assert RegistryBackedRegimeTrainingAdapter._candidate_model_ids(req_ensemble, leg) == ["meta_label_classifier"]

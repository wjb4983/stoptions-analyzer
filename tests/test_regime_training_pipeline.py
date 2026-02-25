from __future__ import annotations

import json

from backtesting.regime_training_pipeline import (
    AdapterIssue,
    LegOutOfSampleMetrics,
    RegimeLegTrainingConfig,
    RegimeTrainingAdapterOutput,
    RegimeTrainingRequest,
    TrainedArtifactLocations,
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
    assert payload["request"]["regime_name"] == "Risk On"
    assert payload["artifact_paths"]["spec"].endswith("regime_spec_snapshot.json")
    assert payload["metadata"]["candidate_leaderboard"] == {}
    assert payload["metadata"]["champion_by_leg"] == {}


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
                model_type="regime_change_detection",
                controls={"model_confidence_min": 0.6},
                model_id="heston_surface_model",
            ),
            RegimeLegTrainingConfig(
                name="Event Intensity",
                model_type="regime_change_detection",
                controls={"model_confidence_min": 0.6},
                model_id="hawkes_jump_intensity",
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
                name="ANN",
                model_type="regime_change_detection",
                controls={"model_confidence_min": 0.6},
                model_id="ann_classifier",
                architecture_spec={"layers": [{"units": 16, "activation": "relu"}]},
            ),
            RegimeLegTrainingConfig(
                name="Vol Surface",
                model_type="regime_change_detection",
                controls={"model_confidence_min": 0.6},
                model_id="heston_surface_model",
                calibration_spec={"model": "heston", "parameters": {"kappa": 1.2}},
            ),
            RegimeLegTrainingConfig(
                name="Event Intensity",
                model_type="regime_change_detection",
                controls={"model_confidence_min": 0.6},
                model_id="hawkes_jump_intensity",
                event_process_spec={"process_type": "hawkes", "parameters": {"alpha": 0.3}},
            ),
        ),
    )

    result = execute_regime_training_pipeline(req)

    assert result.status == "success"

from __future__ import annotations

import json

from backtesting.regime_training_pipeline import (
    RegimeLegTrainingConfig,
    RegimeTrainingRequest,
    execute_regime_training_pipeline,
)


class _FailingAdapter:
    def fit_and_backtest(self, request: RegimeTrainingRequest, run_dir):  # noqa: ANN001
        raise RuntimeError("adapter exploded")


class _StableAdapter:
    def fit_and_backtest(self, request: RegimeTrainingRequest, run_dir):  # noqa: ANN001
        return {
            "metrics": {"sharpe": 1.23, "max_drawdown": 0.11},
            "warnings": ["insufficient history in one bucket"],
            "artifacts": {"equity_curve": run_dir / "equity.csv"},
            "adapter": "stable_test_adapter",
        }


def _request(tmp_path, *, legs: tuple[RegimeLegTrainingConfig, ...]) -> RegimeTrainingRequest:
    return RegimeTrainingRequest(
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
                model_type="trend",
                controls={"model_confidence_min": 0.7, "turnover_limit": 0.2},
            ),
        ),
    )

    result = execute_regime_training_pipeline(req, adapter=_StableAdapter())

    assert result.status == "success"
    assert result.metrics["sharpe"] == 1.23
    assert result.warnings == ("insufficient history in one bucket",)
    assert "manifest" in result.artifact_paths

    manifest_path = result.artifact_paths["manifest"]
    payload = json.loads(open(manifest_path, encoding="utf-8").read())
    assert payload["status"] == "success"
    assert payload["request"]["regime_name"] == "Risk On"
    assert payload["artifact_paths"]["spec"].endswith("regime_spec_snapshot.json")


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
                model_type="carry",
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

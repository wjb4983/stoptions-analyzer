from __future__ import annotations

import hashlib
import json
from dataclasses import asdict, dataclass, field, replace
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Literal, Protocol

import numpy as np

from modeling_nextgen.calibration.probability import ProbabilityCalibrator
from models.regime_catalog import list_models_for_leg, validate_model_leg_pairing
from models.registry import create_model

DEFAULT_REGIME_TRAINING_OUTPUT_DIR = Path("data/regime_training_runs")


@dataclass(frozen=True)
class RegimeLegTrainingConfig:
    name: str
    model_type: str
    controls: dict[str, float]
    model_id: str = ""
    selected_model_id: str = ""
    hyperparameters: dict[str, Any] = field(default_factory=dict)
    architecture_spec: dict[str, Any] | None = None
    calibration_spec: dict[str, Any] | None = None
    event_process_spec: dict[str, Any] | None = None

    @property
    def resolved_model_id(self) -> str:
        model_id = self.model_id.strip()
        if model_id:
            return model_id
        return self.selected_model_id.strip()


@dataclass(frozen=True)
class RegimeTrainingRequest:
    schema_version: int
    regime_id: str
    regime_name: str
    legs: tuple[RegimeLegTrainingConfig, ...]
    model_choice: str
    training_window: dict[str, int]
    risk_limits: dict[str, float]
    output_dir: str | None = None

    @property
    def regime_label(self) -> str:
        """Backwards-compatible alias used by current UI code."""
        return self.regime_name


@dataclass(frozen=True)
class RegimeTrainingResult:
    run_id: str
    status: str
    metrics: dict[str, float]
    artifact_paths: dict[str, str]
    timestamps: dict[str, str]
    warnings: tuple[str, ...] = field(default_factory=tuple)
    errors: tuple[str, ...] = field(default_factory=tuple)
    error_payload: dict[str, Any] | None = None
    summary: str = ""
    metadata: dict[str, Any] = field(default_factory=dict)
    logs: tuple[str, ...] = field(default_factory=tuple)

    @property
    def started_at(self) -> str:
        return self.timestamps.get("started_at", "")

    @property
    def completed_at(self) -> str:
        return self.timestamps.get("completed_at", "")

    @property
    def artifact_path(self) -> str:
        return self.artifact_paths.get("manifest", "")


class RegimeTrainingAdapter(Protocol):
    def fit_and_backtest(self, request: RegimeTrainingRequest, run_dir: Path) -> "RegimeTrainingAdapterOutput":
        """Run fitting + backtest for a regime and return structured adapter output."""


@dataclass(frozen=True)
class AdapterIssue:
    level: Literal["warning", "error"]
    model_id: str
    message: str
    leg_name: str | None = None


@dataclass(frozen=True)
class TrainedArtifactLocations:
    model_weights: str
    calibration_object: str
    diagnostics: str


@dataclass(frozen=True)
class LegOutOfSampleMetrics:
    leg_name: str
    model_id: str
    metrics: dict[str, float]


@dataclass(frozen=True)
class RegimeTrainingAdapterOutput:
    per_leg_artifacts: dict[str, TrainedArtifactLocations]
    per_leg_oos_metrics: dict[str, LegOutOfSampleMetrics]
    portfolio_oos_metrics: dict[str, float]
    issues: tuple[AdapterIssue, ...] = ()
    adapter_name: str = ""

    def warnings(self) -> tuple[str, ...]:
        return tuple(
            f"[{issue.model_id}] {issue.message}" for issue in self.issues if issue.level == "warning"
        )

    def errors(self) -> tuple[str, ...]:
        return tuple(
            f"[{issue.model_id}] {issue.message}" for issue in self.issues if issue.level == "error"
        )


class PlaceholderRegimeTrainingAdapter:
    """Stable placeholder kept for explicit dev/test mode only."""

    def fit_and_backtest(self, request: RegimeTrainingRequest, run_dir: Path) -> RegimeTrainingAdapterOutput:
        issues = (
            AdapterIssue(
                level="warning",
                model_id="placeholder",
                message="Running placeholder adapter output; use for explicit dev/test mode only.",
            ),
        )
        return RegimeTrainingAdapterOutput(
            per_leg_artifacts={},
            per_leg_oos_metrics={},
            portfolio_oos_metrics={"leg_count": float(len(request.legs))},
            issues=issues,
            adapter_name="placeholder",
        )


class RegistryBackedRegimeTrainingAdapter:
    """Production adapter that dispatches each leg by model id from model registry/catalog."""

    def fit_and_backtest(self, request: RegimeTrainingRequest, run_dir: Path) -> RegimeTrainingAdapterOutput:
        artifacts_dir = run_dir / "artifacts"
        artifacts_dir.mkdir(parents=True, exist_ok=True)
        per_leg_artifacts: dict[str, TrainedArtifactLocations] = {}
        per_leg_oos_metrics: dict[str, LegOutOfSampleMetrics] = {}
        issues: list[AdapterIssue] = []

        for idx, leg in enumerate(request.legs):
            leg_key = f"{idx:02d}_{leg.name.lower().replace(' ', '_')}"
            try:
                model_id = self._select_model_id(request, leg)
                validate_model_leg_pairing(leg.model_type, model_id)
                model = create_model(model_id)

                features, labels = self._build_dataset(model.required_feature_names(), leg.controls, seed=idx)
                split_idx = max(int(len(labels) * 0.7), 12)
                train_x = {name: values[:split_idx] for name, values in features.items()}
                test_x = {name: values[split_idx:] for name, values in features.items()}
                train_y = labels[:split_idx]
                test_y = labels[split_idx:]

                model.fit(train_x, train_y)
                train_probs = model.predict_proba(train_x)
                test_probs = model.predict_proba(test_x)

                calibrator = ProbabilityCalibrator(method="platt", n_bins=10)
                calibrator.fit(train_probs, np.where(train_y > 0, 1.0, 0.0))
                calibrated_test_probs = calibrator.transform(test_probs)
                test_labels_binary = np.where(test_y > 0, 1.0, 0.0)
                report = calibrator.report(test_probs, test_labels_binary, calibrated_probabilities=calibrated_test_probs)

                preds = np.where(calibrated_test_probs >= 0.5, 1.0, 0.0)
                metrics = {
                    "accuracy": float(np.mean(preds == test_labels_binary)),
                    "brier_score": float(report.brier_score),
                    "expected_calibration_error": float(report.expected_calibration_error),
                    "avg_confidence": float(np.mean(np.abs(calibrated_test_probs - 0.5) * 2.0)),
                    "oos_sample_size": float(test_labels_binary.size),
                }

                leg_dir = artifacts_dir / leg_key
                leg_dir.mkdir(parents=True, exist_ok=True)
                weights_path = leg_dir / "model_weights.json"
                calibration_path = leg_dir / "calibration_object.json"
                diagnostics_path = leg_dir / "diagnostics.json"
                weights_path.write_text(
                    json.dumps(
                        {
                            "model_id": model_id,
                            "required_features": list(model.required_feature_names()),
                            "feature_importances": {k: float(v) for k, v in model.feature_importances_.items()},
                        },
                        indent=2,
                        sort_keys=True,
                    ),
                    encoding="utf-8",
                )
                calibration_path.write_text(
                    json.dumps(
                        {
                            "model_id": model_id,
                            "method": report.method,
                            "sample_size": report.sample_size,
                            "expected_calibration_error": report.expected_calibration_error,
                            "brier_score": report.brier_score,
                        },
                        indent=2,
                        sort_keys=True,
                    ),
                    encoding="utf-8",
                )
                diagnostics_path.write_text(
                    json.dumps(
                        {
                            "leg_name": leg.name,
                            "model_id": model_id,
                            "oos_metrics": metrics,
                            "controls": {k: float(v) for k, v in leg.controls.items()},
                        },
                        indent=2,
                        sort_keys=True,
                    ),
                    encoding="utf-8",
                )

                per_leg_artifacts[leg.name] = TrainedArtifactLocations(
                    model_weights=str(weights_path),
                    calibration_object=str(calibration_path),
                    diagnostics=str(diagnostics_path),
                )
                per_leg_oos_metrics[leg.name] = LegOutOfSampleMetrics(
                    leg_name=leg.name,
                    model_id=model_id,
                    metrics=metrics,
                )
            except Exception as exc:
                issues.append(
                    AdapterIssue(
                        level="error",
                        model_id=request.model_choice,
                        leg_name=leg.name,
                        message=f"leg '{leg.name}' failed: {exc}",
                    )
                )

        portfolio_oos_metrics = self._aggregate_portfolio_metrics(per_leg_oos_metrics)
        return RegimeTrainingAdapterOutput(
            per_leg_artifacts=per_leg_artifacts,
            per_leg_oos_metrics=per_leg_oos_metrics,
            portfolio_oos_metrics=portfolio_oos_metrics,
            issues=tuple(issues),
            adapter_name="registry_backed",
        )

    @staticmethod
    def _build_dataset(
        required_features: tuple[str, ...],
        controls: dict[str, float],
        *,
        seed: int,
        sample_size: int = 160,
    ) -> tuple[dict[str, np.ndarray], np.ndarray]:
        rng_seed = int(sum(abs(float(v)) for v in controls.values()) * 10_000) + seed
        rng = np.random.default_rng(rng_seed)
        features: dict[str, np.ndarray] = {}
        for feature_name in required_features:
            features[feature_name] = rng.normal(loc=0.0, scale=1.0, size=sample_size).astype(float)

        latent = np.zeros(sample_size, dtype=float)
        for feature_values in features.values():
            latent += feature_values
        latent += rng.normal(loc=0.0, scale=0.5, size=sample_size)
        labels = np.where(latent >= np.median(latent), 1.0, -1.0)
        return features, labels

    @staticmethod
    def _select_model_id(request: RegimeTrainingRequest, leg: RegimeLegTrainingConfig) -> str:
        descriptors = list_models_for_leg(leg.model_type)
        if not descriptors:
            raise ValueError(f"No catalog entries configured for leg type '{leg.model_type}'")

        mode = request.model_choice.strip().lower()
        selected_model_id = leg.resolved_model_id.lower()
        allowed = {item.model_name for item in descriptors}

        if mode in {"single_model", "auto_model_search", "", "auto"}:
            if selected_model_id and selected_model_id in allowed:
                return selected_model_id
            return descriptors[0].model_name

        if mode == "ensemble":
            for descriptor in descriptors:
                if descriptor.model_name == "meta_label_classifier":
                    return descriptor.model_name
            return descriptors[0].model_name

        if mode in allowed:
            return mode

        return descriptors[0].model_name

    @staticmethod
    def _aggregate_portfolio_metrics(
        per_leg_oos_metrics: dict[str, LegOutOfSampleMetrics],
    ) -> dict[str, float]:
        if not per_leg_oos_metrics:
            return {"legs_trained": 0.0}

        metric_keys = set()
        for leg_metrics in per_leg_oos_metrics.values():
            metric_keys.update(leg_metrics.metrics.keys())

        aggregate: dict[str, float] = {"legs_trained": float(len(per_leg_oos_metrics))}
        for metric_key in sorted(metric_keys):
            values = [float(item.metrics[metric_key]) for item in per_leg_oos_metrics.values() if metric_key in item.metrics]
            if values:
                aggregate[f"portfolio_avg_{metric_key}"] = float(np.mean(values))
        return aggregate


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _deterministic_run_id(request: RegimeTrainingRequest) -> str:
    payload = asdict(request)
    payload["legs"] = [asdict(leg) for leg in request.legs]
    digest = hashlib.sha256(json.dumps(payload, sort_keys=True).encode("utf-8")).hexdigest()
    return digest[:12]


def validate_regime_spec(request: RegimeTrainingRequest) -> list[str]:
    errors: list[str] = []
    if int(request.schema_version) < 2:
        errors.append("schema_version must be >= 2")
    if not request.regime_id.strip():
        errors.append("regime_id is required")
    if not request.regime_name.strip():
        errors.append("regime_name is required")
    if not request.model_choice.strip():
        errors.append("model_choice is required")
    if not request.legs:
        errors.append("at least one leg is required")

    for idx, leg in enumerate(request.legs):
        if not leg.name.strip():
            errors.append(f"legs[{idx}].name is required")
        if not leg.model_type.strip():
            errors.append(f"legs[{idx}].model_type is required")
        if not isinstance(leg.hyperparameters, dict):
            errors.append(f"legs[{idx}].hyperparameters must be an object")

    retrain_days = request.training_window.get("retrain_frequency_days")
    if retrain_days is not None and int(retrain_days) <= 0:
        errors.append("training_window.retrain_frequency_days must be > 0")

    for key, value in request.risk_limits.items():
        if float(value) < 0:
            errors.append(f"risk_limits.{key} must be >= 0")

    errors.extend(validate_model_specific_specs(request))
    return errors


def validate_model_specific_specs(request: RegimeTrainingRequest) -> list[str]:
    errors: list[str] = []
    for idx, leg in enumerate(request.legs):
        model_id = leg.resolved_model_id.lower()

        errors.extend(
            _validate_optional_spec_object(leg.architecture_spec, field_path=f"legs[{idx}].architecture_spec")
        )
        errors.extend(_validate_optional_spec_object(leg.calibration_spec, field_path=f"legs[{idx}].calibration_spec"))
        errors.extend(
            _validate_optional_spec_object(leg.event_process_spec, field_path=f"legs[{idx}].event_process_spec")
        )

        if _requires_architecture_spec(model_id):
            errors.extend(
                _validate_architecture_spec(leg.architecture_spec, field_path=f"legs[{idx}].architecture_spec")
            )
        if _requires_calibration_spec(model_id):
            errors.extend(
                _validate_calibration_spec(leg.calibration_spec, field_path=f"legs[{idx}].calibration_spec")
            )
        if _requires_event_process_spec(model_id):
            errors.extend(
                _validate_event_process_spec(leg.event_process_spec, field_path=f"legs[{idx}].event_process_spec")
            )
    return errors


def _validate_optional_spec_object(payload: dict[str, Any] | None, *, field_path: str) -> list[str]:
    if payload is None:
        return []
    if isinstance(payload, dict):
        return []
    return [f"{field_path} must be an object when provided"]


def _requires_architecture_spec(model_id: str) -> bool:
    return any(token in model_id for token in ("ann", "neural", "transformer", "mlp"))


def _requires_calibration_spec(model_id: str) -> bool:
    return any(token in model_id for token in ("local_vol", "heston", "sabr", "vol_surface", "black_scholes"))


def _requires_event_process_spec(model_id: str) -> bool:
    return any(token in model_id for token in ("hawkes", "jump", "intensity"))


def _validate_architecture_spec(payload: dict[str, Any] | None, *, field_path: str) -> list[str]:
    if not isinstance(payload, dict):
        return [f"{field_path} is required for ANN/neural model legs"]
    layers = payload.get("layers")
    if not isinstance(layers, list) or not layers:
        return [f"{field_path}.layers must be a non-empty list"]
    return []


def _validate_calibration_spec(payload: dict[str, Any] | None, *, field_path: str) -> list[str]:
    if not isinstance(payload, dict):
        return [f"{field_path} is required for volatility/surface calibration models"]
    errors: list[str] = []
    if not str(payload.get("model", "")).strip():
        errors.append(f"{field_path}.model is required")
    parameters = payload.get("parameters")
    if not isinstance(parameters, dict):
        errors.append(f"{field_path}.parameters must be an object")
    return errors


def _validate_event_process_spec(payload: dict[str, Any] | None, *, field_path: str) -> list[str]:
    if not isinstance(payload, dict):
        return [f"{field_path} is required for Hawkes/jump-intensity models"]
    errors: list[str] = []
    if not str(payload.get("process_type", "")).strip():
        errors.append(f"{field_path}.process_type is required")
    parameters = payload.get("parameters")
    if not isinstance(parameters, dict):
        errors.append(f"{field_path}.parameters must be an object")
    return errors


def save_regime_spec(request: RegimeTrainingRequest, run_dir: Path) -> Path:
    spec_path = run_dir / "regime_spec_snapshot.json"
    spec_path.write_text(json.dumps(asdict(request), indent=2, sort_keys=True), encoding="utf-8")
    return spec_path


def compute_summary_metrics(
    adapter_output: RegimeTrainingAdapterOutput,
    request: RegimeTrainingRequest,
) -> tuple[dict[str, float], str]:
    metrics = {key: float(value) for key, value in adapter_output.portfolio_oos_metrics.items()}
    summary = (
        f"Trained {len(request.legs)} leg(s) for regime '{request.regime_name}' "
        f"using model choice '{request.model_choice}'."
    )
    return metrics, summary


def _flatten_artifact_locations(
    output: RegimeTrainingAdapterOutput,
) -> dict[str, str]:
    paths: dict[str, str] = {}
    for leg_name, artifacts in output.per_leg_artifacts.items():
        safe_leg = leg_name.lower().replace(" ", "_")
        paths[f"{safe_leg}_model_weights"] = artifacts.model_weights
        paths[f"{safe_leg}_calibration_object"] = artifacts.calibration_object
        paths[f"{safe_leg}_diagnostics"] = artifacts.diagnostics
    return paths


def write_regime_training_manifest(
    *,
    run_dir: Path,
    request: RegimeTrainingRequest,
    result: RegimeTrainingResult,
) -> Path:
    manifest_path = run_dir / "manifest.json"
    manifest = {
        "run_id": result.run_id,
        "status": result.status,
        "request": asdict(request),
        "metrics": result.metrics,
        "timestamps": result.timestamps,
        "summary": result.summary,
        "warnings": list(result.warnings),
        "errors": list(result.errors),
        "error_payload": result.error_payload,
        "artifact_paths": result.artifact_paths,
        "metadata": result.metadata,
        "logs": list(result.logs),
    }
    manifest_path.write_text(json.dumps(manifest, indent=2, sort_keys=True), encoding="utf-8")
    return manifest_path


def run_regime_training(
    request: RegimeTrainingRequest,
    output_dir: str | Path | None = None,
    adapter: RegimeTrainingAdapter | None = None,
) -> RegimeTrainingResult:
    started_at = _utc_now_iso()
    run_id = _deterministic_run_id(request)
    resolved_output_dir = request.output_dir or output_dir
    output_root = (
        Path(resolved_output_dir) if resolved_output_dir is not None else DEFAULT_REGIME_TRAINING_OUTPUT_DIR
    )
    run_dir = output_root / f"run_{run_id}"
    run_dir.mkdir(parents=True, exist_ok=True)

    spec_path = save_regime_spec(request, run_dir)
    validation_errors = validate_regime_spec(request)
    if validation_errors:
        completed_at = _utc_now_iso()
        error_payload = {
            "code": "INVALID_REGIME_SPEC",
            "stage": "validate_regime_spec",
            "errors": validation_errors,
        }
        result = RegimeTrainingResult(
            run_id=run_id,
            status="failed",
            metrics={},
            artifact_paths={"spec": str(spec_path)},
            timestamps={"started_at": started_at, "completed_at": completed_at},
            warnings=(),
            errors=tuple(validation_errors),
            error_payload=error_payload,
            summary="Regime training request validation failed.",
            metadata={"regime_id": request.regime_id, "regime_name": request.regime_name},
            logs=("saved config snapshot", "validation failed"),
        )
        manifest_path = write_regime_training_manifest(run_dir=run_dir, request=request, result=result)
        return replace(result, artifact_paths={**result.artifact_paths, "manifest": str(manifest_path)})

    runner = adapter or RegistryBackedRegimeTrainingAdapter()
    if request.model_choice.strip().lower() in {"placeholder", "dev", "test"}:
        runner = adapter or PlaceholderRegimeTrainingAdapter()
    try:
        adapter_output = runner.fit_and_backtest(request, run_dir)
        metrics, summary = compute_summary_metrics(adapter_output, request)
        completed_at = _utc_now_iso()
        warnings = adapter_output.warnings()
        adapter_errors = adapter_output.errors()
        artifacts = {"spec": str(spec_path), **_flatten_artifact_locations(adapter_output)}
        oos_metrics_payload = {
            leg_name: {
                "model_id": leg_metrics.model_id,
                "metrics": {k: float(v) for k, v in leg_metrics.metrics.items()},
            }
            for leg_name, leg_metrics in adapter_output.per_leg_oos_metrics.items()
        }
        logs = (
            "saved config snapshot",
            f"adapter={adapter_output.adapter_name or runner.__class__.__name__}",
            "computed summary metrics",
        )
        status = "failed" if adapter_errors else "success"
        result = RegimeTrainingResult(
            run_id=run_id,
            status=status,
            metrics=metrics,
            artifact_paths=artifacts,
            timestamps={"started_at": started_at, "completed_at": completed_at},
            warnings=warnings,
            errors=adapter_errors,
            error_payload=(
                {
                    "code": "PARTIAL_TRAINING_FAILURE",
                    "stage": "fit_and_backtest",
                    "errors": list(adapter_errors),
                }
                if adapter_errors
                else None
            ),
            summary=summary,
            metadata={
                "regime_id": request.regime_id,
                "regime_name": request.regime_name,
                "model_choice": request.model_choice,
                "oos_metrics": oos_metrics_payload,
            },
            logs=logs,
        )
    except Exception as exc:
        completed_at = _utc_now_iso()
        error_payload = {
            "code": "TRAINING_EXECUTION_FAILED",
            "stage": "fit_and_backtest",
            "message": str(exc),
            "exception_type": type(exc).__name__,
        }
        result = RegimeTrainingResult(
            run_id=run_id,
            status="failed",
            metrics={},
            artifact_paths={"spec": str(spec_path)},
            timestamps={"started_at": started_at, "completed_at": completed_at},
            warnings=(),
            errors=(str(exc),),
            error_payload=error_payload,
            summary="Regime training failed during fit/backtest stage.",
            metadata={"regime_id": request.regime_id, "regime_name": request.regime_name},
            logs=("saved config snapshot", "fit_and_backtest failed"),
        )

    manifest_path = write_regime_training_manifest(run_dir=run_dir, request=request, result=result)
    return replace(result, artifact_paths={**result.artifact_paths, "manifest": str(manifest_path)})


def execute_regime_training_pipeline(
    request: RegimeTrainingRequest,
    *,
    adapter: RegimeTrainingAdapter | None = None,
    output_dir: str | Path | None = None,
) -> RegimeTrainingResult:
    """UI seam point for Create Regime and Research Lab orchestration."""
    return run_regime_training(request=request, output_dir=output_dir, adapter=adapter)

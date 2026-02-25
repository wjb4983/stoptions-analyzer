from __future__ import annotations

import hashlib
import json
from dataclasses import asdict, dataclass, field, replace
from datetime import datetime, timezone
from pathlib import Path
from statistics import mean
from typing import Any, Protocol

DEFAULT_REGIME_TRAINING_OUTPUT_DIR = Path("data/regime_training_runs")


@dataclass(frozen=True)
class RegimeLegTrainingConfig:
    name: str
    model_type: str
    controls: dict[str, float]


@dataclass(frozen=True)
class RegimeTrainingRequest:
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
    def fit_and_backtest(self, request: RegimeTrainingRequest, run_dir: Path) -> dict[str, Any]:
        """Run fitting + backtest for a regime and return deterministic metrics payload."""


class PlaceholderRegimeTrainingAdapter:
    """Stable placeholder until richer model/backtest engines are wired in."""

    def fit_and_backtest(self, request: RegimeTrainingRequest, run_dir: Path) -> dict[str, Any]:
        avg_confidence = mean(
            float(leg.controls.get("model_confidence_min", 0.6)) for leg in request.legs
        )
        avg_turnover = mean(float(leg.controls.get("turnover_limit", 0.3)) for leg in request.legs)
        avg_slippage = mean(float(leg.controls.get("slippage_bps", 8.0)) for leg in request.legs)
        expected_retrain_days = float(request.training_window.get("retrain_frequency_days", 21))

        return {
            "metrics": {
                "leg_count": float(len(request.legs)),
                "avg_model_confidence_min": round(avg_confidence, 4),
                "avg_turnover_limit": round(avg_turnover, 4),
                "avg_slippage_bps": round(avg_slippage, 4),
                "expected_retrain_frequency_days": expected_retrain_days,
            },
            "warnings": [],
            "artifacts": {},
            "adapter": "placeholder",
        }


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _deterministic_run_id(request: RegimeTrainingRequest) -> str:
    payload = asdict(request)
    payload["legs"] = [asdict(leg) for leg in request.legs]
    digest = hashlib.sha256(json.dumps(payload, sort_keys=True).encode("utf-8")).hexdigest()
    return digest[:12]


def validate_regime_spec(request: RegimeTrainingRequest) -> list[str]:
    errors: list[str] = []
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

    retrain_days = request.training_window.get("retrain_frequency_days")
    if retrain_days is not None and int(retrain_days) <= 0:
        errors.append("training_window.retrain_frequency_days must be > 0")

    for key, value in request.risk_limits.items():
        if float(value) < 0:
            errors.append(f"risk_limits.{key} must be >= 0")

    return errors


def save_regime_spec(request: RegimeTrainingRequest, run_dir: Path) -> Path:
    spec_path = run_dir / "regime_spec_snapshot.json"
    spec_path.write_text(json.dumps(asdict(request), indent=2, sort_keys=True), encoding="utf-8")
    return spec_path


def compute_summary_metrics(adapter_output: dict[str, Any], request: RegimeTrainingRequest) -> tuple[dict[str, float], str]:
    metrics = {
        key: float(value)
        for key, value in dict(adapter_output.get("metrics", {})).items()
    }
    summary = (
        f"Trained {len(request.legs)} leg(s) for regime '{request.regime_name}' "
        f"using model choice '{request.model_choice}'."
    )
    return metrics, summary


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

    runner = adapter or PlaceholderRegimeTrainingAdapter()
    try:
        adapter_output = runner.fit_and_backtest(request, run_dir)
        metrics, summary = compute_summary_metrics(adapter_output, request)
        completed_at = _utc_now_iso()
        warnings = tuple(str(item) for item in adapter_output.get("warnings", []))
        artifacts = {"spec": str(spec_path), **{k: str(v) for k, v in dict(adapter_output.get("artifacts", {})).items()}}
        logs = (
            "saved config snapshot",
            f"adapter={adapter_output.get('adapter', runner.__class__.__name__)}",
            "computed summary metrics",
        )
        result = RegimeTrainingResult(
            run_id=run_id,
            status="success",
            metrics=metrics,
            artifact_paths=artifacts,
            timestamps={"started_at": started_at, "completed_at": completed_at},
            warnings=warnings,
            errors=(),
            error_payload=None,
            summary=summary,
            metadata={
                "regime_id": request.regime_id,
                "regime_name": request.regime_name,
                "model_choice": request.model_choice,
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

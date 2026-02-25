from __future__ import annotations

import hashlib
import json
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Any

DEFAULT_REGIME_EXPORT_OUTPUT_DIR = Path("data/regime_exports")
EXPORT_SCHEMA_VERSION = "1.0"


@dataclass(frozen=True)
class RegimeExportBundle:
    bundle_id: str
    bundle_dir: str
    bundle_manifest_path: str
    deployment_version: str
    exported_paths: dict[str, str]


def export_regime_training_bundle(
    training_manifest_path: str | Path,
    *,
    output_dir: str | Path | None = None,
) -> RegimeExportBundle:
    manifest_path = Path(training_manifest_path)
    payload = _read_manifest(manifest_path)

    run_id = str(payload.get("run_id", "")).strip()
    if not run_id:
        raise ValueError("manifest missing run_id")

    request = payload.get("request", {}) if isinstance(payload.get("request"), dict) else {}
    artifact_paths = payload.get("artifact_paths", {}) if isinstance(payload.get("artifact_paths"), dict) else {}
    metrics = payload.get("metrics", {}) if isinstance(payload.get("metrics"), dict) else {}
    metadata = payload.get("metadata", {}) if isinstance(payload.get("metadata"), dict) else {}
    _validate_synthetic_fallback_export_guard(payload)

    deployment_version = _deterministic_deployment_version(
        run_id=run_id,
        request=request,
        artifact_paths=artifact_paths,
        metrics=metrics,
        metadata=metadata,
    )
    bundle_id = f"regime_export_{deployment_version}"

    export_root = Path(output_dir) if output_dir is not None else DEFAULT_REGIME_EXPORT_OUTPUT_DIR
    bundle_dir = export_root / bundle_id
    bundle_dir.mkdir(parents=True, exist_ok=True)

    models_dir = bundle_dir / "models"
    calibrations_dir = bundle_dir / "calibration"
    diagnostics_dir = bundle_dir / "evaluation"
    metadata_dir = bundle_dir / "metadata"
    for folder in (models_dir, calibrations_dir, diagnostics_dir, metadata_dir):
        folder.mkdir(parents=True, exist_ok=True)

    copied: dict[str, str] = {}
    copied["training_manifest"] = _copy_if_exists(manifest_path, metadata_dir / "training_manifest.json")

    spec_path_raw = artifact_paths.get("spec")
    if isinstance(spec_path_raw, str) and spec_path_raw.strip():
        copied["regime_spec_snapshot"] = _copy_if_exists(Path(spec_path_raw), metadata_dir / "regime_spec_snapshot.json")

    model_keys = sorted(k for k in artifact_paths if k.endswith("_model_weights"))
    calibration_keys = sorted(k for k in artifact_paths if k.endswith("_calibration_object"))
    diagnostics_keys = sorted(k for k in artifact_paths if k.endswith("_diagnostics"))

    for key in model_keys:
        copied[key] = _copy_if_exists(Path(str(artifact_paths[key])), models_dir / f"{key}.json")

    for key in calibration_keys:
        copied[key] = _copy_if_exists(Path(str(artifact_paths[key])), calibrations_dir / f"{key}.json")

    for key in diagnostics_keys:
        copied[key] = _copy_if_exists(Path(str(artifact_paths[key])), diagnostics_dir / f"{key}.json")

    feature_schema_payload = _build_feature_schema_payload(copied, request)
    feature_schema_path = metadata_dir / "feature_schema.json"
    feature_schema_path.write_text(json.dumps(feature_schema_payload, indent=2, sort_keys=True), encoding="utf-8")
    copied["feature_schema"] = str(feature_schema_path)

    evaluation_report = {
        "run_id": run_id,
        "status": payload.get("status"),
        "summary": payload.get("summary", ""),
        "portfolio_metrics": metrics,
        "oos_metrics": metadata.get("oos_metrics", {}) if isinstance(metadata.get("oos_metrics"), dict) else {},
        "warnings": payload.get("warnings", []),
        "errors": payload.get("errors", []),
    }
    evaluation_report_path = diagnostics_dir / "evaluation_report.json"
    evaluation_report_path.write_text(json.dumps(evaluation_report, indent=2, sort_keys=True), encoding="utf-8")
    copied["evaluation_report"] = str(evaluation_report_path)

    provenance_payload = {
        "export_schema_version": EXPORT_SCHEMA_VERSION,
        "bundle_id": bundle_id,
        "deployment_version": deployment_version,
        "source_run": {
            "run_id": run_id,
            "status": payload.get("status"),
            "timestamps": payload.get("timestamps", {}),
            "summary": payload.get("summary", ""),
        },
        "source_manifest_path": str(manifest_path),
        "artifact_hashes": _artifact_hashes(copied),
        "request": request,
    }
    provenance_path = metadata_dir / "provenance.json"
    provenance_path.write_text(json.dumps(provenance_payload, indent=2, sort_keys=True), encoding="utf-8")
    copied["provenance"] = str(provenance_path)

    bundle_manifest_payload = {
        "bundle_id": bundle_id,
        "bundle_version": EXPORT_SCHEMA_VERSION,
        "deployment_version": deployment_version,
        "run_id": run_id,
        "bundle_dir": str(bundle_dir),
        "contents": copied,
    }
    bundle_manifest_path = bundle_dir / "bundle_manifest.json"
    bundle_manifest_path.write_text(json.dumps(bundle_manifest_payload, indent=2, sort_keys=True), encoding="utf-8")

    return RegimeExportBundle(
        bundle_id=bundle_id,
        bundle_dir=str(bundle_dir),
        bundle_manifest_path=str(bundle_manifest_path),
        deployment_version=deployment_version,
        exported_paths=copied,
    )


def _read_manifest(manifest_path: Path) -> dict[str, Any]:
    if not manifest_path.exists():
        raise FileNotFoundError(f"training manifest not found: {manifest_path}")
    payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError("training manifest must be a JSON object")
    return payload


def _deterministic_deployment_version(
    *,
    run_id: str,
    request: dict[str, Any],
    artifact_paths: dict[str, Any],
    metrics: dict[str, Any],
    metadata: dict[str, Any],
) -> str:
    fingerprint_payload = {
        "run_id": run_id,
        "request": request,
        "artifacts": artifact_paths,
        "metrics": metrics,
        "metadata": metadata,
    }
    digest = hashlib.sha256(json.dumps(fingerprint_payload, sort_keys=True).encode("utf-8")).hexdigest()[:10]
    return f"{run_id}-{digest}"


def _copy_if_exists(source: Path, destination: Path) -> str:
    destination.parent.mkdir(parents=True, exist_ok=True)
    if source.exists():
        shutil.copy2(source, destination)
        return str(destination)
    destination.write_text(
        json.dumps({"missing_source": str(source)}, indent=2, sort_keys=True),
        encoding="utf-8",
    )
    return str(destination)


def _build_feature_schema_payload(copied_paths: dict[str, str], request: dict[str, Any]) -> dict[str, Any]:
    schema_version = str(request.get("schema_version", "")) or "unknown"
    required_features: set[str] = set()

    for key, path in copied_paths.items():
        if not key.endswith("_model_weights"):
            continue
        model_payload = json.loads(Path(path).read_text(encoding="utf-8"))
        if isinstance(model_payload, dict) and isinstance(model_payload.get("required_features"), list):
            for name in model_payload["required_features"]:
                required_features.add(str(name))

    return {
        "feature_schema_version": f"regime-v{schema_version}",
        "required_features": sorted(required_features),
        "feature_count": len(required_features),
    }


def _artifact_hashes(paths: dict[str, str]) -> dict[str, str]:
    hashes: dict[str, str] = {}
    for key, path in sorted(paths.items()):
        body = Path(path).read_bytes()
        hashes[key] = hashlib.sha256(body).hexdigest()
    return hashes


def _validate_synthetic_fallback_export_guard(payload: dict[str, Any]) -> None:
    metadata = payload.get("metadata", {}) if isinstance(payload.get("metadata"), dict) else {}
    request = payload.get("request", {}) if isinstance(payload.get("request"), dict) else {}
    settings = (
        request.get("training_data_settings", {})
        if isinstance(request.get("training_data_settings"), dict)
        else {}
    )
    fallback_used = bool(metadata.get("synthetic_fallback_used", False))
    fallback_allowed = bool(settings.get("allow_synthetic_fallback", False))
    if fallback_used and not fallback_allowed:
        raise ValueError(
            "Export blocked: synthetic fallback was used without "
            "request.training_data_settings.allow_synthetic_fallback=true"
        )

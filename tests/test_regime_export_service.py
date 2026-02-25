from __future__ import annotations

import json

import pytest
from pathlib import Path

from backtesting.regime_export_service import export_regime_training_bundle


def _write_json(path: Path, payload: dict) -> str:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, sort_keys=True), encoding="utf-8")
    return str(path)


def test_export_bundle_collects_expected_artifacts_and_metadata(tmp_path):
    run_dir = tmp_path / "run"
    spec = _write_json(run_dir / "regime_spec_snapshot.json", {"schema_version": 2, "regime_id": "risk_on"})
    weights = _write_json(run_dir / "trend_model_weights.json", {"required_features": ["ret_5d", "vol_20d"]})
    calibration = _write_json(run_dir / "trend_calibration_object.json", {"method": "platt"})
    diagnostics = _write_json(run_dir / "trend_diagnostics.json", {"accuracy": 0.71})

    manifest = {
        "run_id": "abcd1234",
        "status": "success",
        "summary": "ok",
        "timestamps": {"started_at": "2024-01-01T00:00:00+00:00", "completed_at": "2024-01-01T00:01:00+00:00"},
        "request": {"schema_version": 2, "regime_name": "Risk On"},
        "metrics": {"portfolio_avg_accuracy": 0.71},
        "metadata": {"oos_metrics": {"Trend": {"accuracy": 0.71}}},
        "artifact_paths": {
            "spec": spec,
            "trend_model_weights": weights,
            "trend_calibration_object": calibration,
            "trend_diagnostics": diagnostics,
        },
    }
    manifest_path = _write_json(run_dir / "manifest.json", manifest)

    bundle = export_regime_training_bundle(manifest_path, output_dir=tmp_path / "exports")

    assert bundle.bundle_id.startswith("regime_export_abcd1234-")
    assert Path(bundle.bundle_manifest_path).exists()
    assert Path(bundle.exported_paths["regime_spec_snapshot"]).exists()
    assert Path(bundle.exported_paths["trend_model_weights"]).exists()
    assert Path(bundle.exported_paths["trend_calibration_object"]).exists()
    assert Path(bundle.exported_paths["trend_diagnostics"]).exists()

    feature_schema = json.loads(Path(bundle.exported_paths["feature_schema"]).read_text(encoding="utf-8"))
    assert feature_schema["feature_schema_version"] == "regime-v2"
    assert feature_schema["required_features"] == ["ret_5d", "vol_20d"]


def test_export_bundle_versioning_is_deterministic(tmp_path):
    run_dir = tmp_path / "run"
    manifest_path = _write_json(
        run_dir / "manifest.json",
        {
            "run_id": "run-z",
            "status": "success",
            "request": {"schema_version": 3, "regime_name": "Risk Off"},
            "metrics": {"portfolio_avg_accuracy": 0.55},
            "metadata": {},
            "artifact_paths": {},
        },
    )

    first = export_regime_training_bundle(manifest_path, output_dir=tmp_path / "exports")
    second = export_regime_training_bundle(manifest_path, output_dir=tmp_path / "exports")

    assert first.deployment_version == second.deployment_version
    assert first.bundle_dir == second.bundle_dir


def test_export_bundle_blocks_unintentional_synthetic_fallback(tmp_path):
    run_dir = tmp_path / "run"
    manifest_path = _write_json(
        run_dir / "manifest.json",
        {
            "run_id": "run-fallback",
            "status": "success",
            "request": {
                "schema_version": 2,
                "regime_name": "Risk Off",
                "training_data_settings": {"allow_synthetic_fallback": False},
            },
            "metrics": {"portfolio_avg_accuracy": 0.55},
            "metadata": {"synthetic_fallback_used": True},
            "artifact_paths": {},
        },
    )

    with pytest.raises(ValueError, match="synthetic fallback"):
        export_regime_training_bundle(manifest_path, output_dir=tmp_path / "exports")

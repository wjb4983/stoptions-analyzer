from __future__ import annotations

import json
from pathlib import Path

from models.model_profiles import build_model_profile_registry


def test_registry_includes_catalog_preset_and_trained_profiles(tmp_path: Path) -> None:
    manifest = {
        "request": {
            "legs": [
                {
                    "name": "Trend leg",
                    "model_type": "timeseries_momentum",
                    "model_id": "momentum_forecasting",
                    "selected_model_id": "momentum_forecasting",
                    "hyperparameters": {"lookback_days": 300},
                }
            ]
        },
        "metadata": {
            "reproducibility": {
                "legs": {
                    "00:Trend leg": {
                        "hyperparameters_checksum": "abc123",
                        "architecture_spec_checksum": "def456",
                    }
                }
            }
        },
    }
    manifest_path = tmp_path / "manifest.json"
    manifest_path.write_text(json.dumps(manifest), encoding="utf-8")

    presets = {
        "my_preset": {
            "profile_id": "preset:timeseries_momentum:my_preset",
            "display_name": "My preset",
            "leg_family": "timeseries_momentum",
            "base_model_id": "momentum",
            "hyperparameters": {"lookback_days": 77},
        }
    }
    runs = [{"run_id": "run_123", "artifact_path": str(manifest_path)}]

    registry = build_model_profile_registry(
        leg_family="timeseries_momentum",
        presets=presets,
        training_runs=runs,
    )

    assert any(p.source == "catalog" for p in registry.catalog_profiles)
    assert any(p.profile_id == "preset:timeseries_momentum:my_preset" for p in registry.preset_profiles)
    trained = next(p for p in registry.trained_profiles if p.base_model_id == "momentum_forecasting")
    assert trained.artifact_reference is not None
    assert trained.artifact_reference.run_id == "run_123"
    assert trained.artifact_reference.checksum_metadata["hyperparameters_checksum"] == "abc123"

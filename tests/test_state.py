from __future__ import annotations

import json

import state
from state import AppState, REGIME_DEFINITION_SCHEMA_VERSION


def test_app_state_save_and_load_roundtrip_includes_regime_keys(tmp_path, monkeypatch) -> None:
    app_state_path = tmp_path / "app_state.json"
    monkeypatch.setattr(state, "STATE_PATH", app_state_path)

    original = AppState(
        tickers=["AAPL", "MSFT"],
        selected_ticker="AAPL",
        regime_definitions={
            "bull": {
                "label": "Bull",
                "confidence_thresholds": {"entry": 0.7},
                "legs": [
                    {
                        "name": "trend",
                        "model_type": "timeseries_momentum",
                        "controls": {},
                        "selected_model_id": "ann_classifier",
                        "hyperparameters": {"depth": 2},
                    }
                ],
            }
        },
        regime_training_runs=[{"id": "run-1", "regime_id": "bull", "score": 0.83}],
        active_regime_id="bull",
    )
    original.save()

    payload = json.loads(app_state_path.read_text(encoding="utf-8"))
    assert payload["regime_definitions"]["bull"]["schema_version"] == REGIME_DEFINITION_SCHEMA_VERSION
    migrated_leg = payload["regime_definitions"]["bull"]["legs"][0]
    assert migrated_leg["model_id"] == "ann_classifier"
    assert payload["regime_training_runs"] == original.regime_training_runs
    assert payload["active_regime_id"] == original.active_regime_id

    loaded = AppState.load()
    assert loaded.regime_definitions["bull"]["schema_version"] == REGIME_DEFINITION_SCHEMA_VERSION
    assert loaded.regime_definitions["bull"]["legs"][0]["model_id"] == "ann_classifier"
    assert loaded.regime_training_runs == original.regime_training_runs
    assert loaded.active_regime_id == original.active_regime_id


def test_legacy_app_state_without_regime_keys_loads_defaults(tmp_path, monkeypatch) -> None:
    app_state_path = tmp_path / "legacy_state.json"
    monkeypatch.setattr(state, "STATE_PATH", app_state_path)

    legacy_payload = {
        "tickers": ["AAPL"],
        "selected_ticker": "AAPL",
        "analysis_mode": "Stock Analysis",
        "option_strategy": "Naked Call",
        "general_analysis_settings": {"analysis_type": "Cross-Sectional"},
        "backtest_settings": {"strategy": "momentum"},
        "backtest_templates": {},
    }
    app_state_path.write_text(json.dumps(legacy_payload), encoding="utf-8")

    loaded = AppState.load()

    assert "baseline" in loaded.regime_definitions
    assert loaded.regime_definitions["baseline"]["schema_version"] == REGIME_DEFINITION_SCHEMA_VERSION
    assert loaded.regime_training_runs == []
    assert loaded.active_regime_id is None


def test_legacy_regime_definition_payload_migrates_to_versioned_leg_schema(tmp_path, monkeypatch) -> None:
    app_state_path = tmp_path / "legacy_regime_state.json"
    monkeypatch.setattr(state, "STATE_PATH", app_state_path)

    legacy_payload = {
        "regime_definitions": {
            "legacy": {
                "label": "Legacy",
                "legs": [
                    {
                        "name": "legacy leg",
                        "model_type": "regime_change_detection",
                        "selected_model_id": "heston_surface_model",
                        "hyperparameters": [],
                    }
                ],
            }
        }
    }
    app_state_path.write_text(json.dumps(legacy_payload), encoding="utf-8")

    loaded = AppState.load()

    leg = loaded.regime_definitions["legacy"]["legs"][0]
    assert loaded.regime_definitions["legacy"]["schema_version"] == REGIME_DEFINITION_SCHEMA_VERSION
    assert leg["model_id"] == "heston_surface_model"
    assert leg["hyperparameters"] == {}
    assert leg["architecture_spec"] is None
    assert leg["calibration_spec"] is None
    assert leg["event_process_spec"] is None



def test_migrate_regime_leg_model_type_aliases_to_canonical_family(tmp_path, monkeypatch) -> None:
    app_state_path = tmp_path / "legacy_ui_type_state.json"
    monkeypatch.setattr(state, "STATE_PATH", app_state_path)

    legacy_payload = {
        "regime_definitions": {
            "legacy": {
                "label": "Legacy UI mapping",
                "legs": [
                    {
                        "name": "macro leg",
                        "model_type": "Cross-Asset Macro",
                        "selected_model_id": "macro_regime_conditioned",
                        "hyperparameters": {},
                    }
                ],
            }
        }
    }
    app_state_path.write_text(json.dumps(legacy_payload), encoding="utf-8")

    loaded = AppState.load()
    leg = loaded.regime_definitions["legacy"]["legs"][0]

    assert leg["model_type"] == "cross_asset_macro_conditioned"


def test_remote_execution_settings_roundtrip_and_defaults(tmp_path, monkeypatch) -> None:
    app_state_path = tmp_path / "remote_state.json"
    monkeypatch.setattr(state, "STATE_PATH", app_state_path)

    original = AppState(
        remote_execution_settings={
            "mode": "remote",
            "ssh_host": "example.internal",
            "ssh_port": "2222",
            "ssh_user": "quant",
            "remote_project_path": "/srv/stoptions/jobs",
            "remote_python_command": "/usr/bin/python3",
            "api_key_policy": "server_only",
        }
    )
    original.save()
    loaded = AppState.load()
    assert loaded.remote_execution_settings["mode"] == "remote"
    assert loaded.remote_execution_settings["ssh_host"] == "example.internal"
    assert loaded.remote_execution_settings["ssh_port"] == "2222"
    assert loaded.remote_execution_settings["api_key_policy"] == "server_only"


def test_remote_jobs_roundtrip_and_backfill_from_active_jobs(tmp_path, monkeypatch) -> None:
    app_state_path = tmp_path / "remote_jobs_state.json"
    monkeypatch.setattr(state, "STATE_PATH", app_state_path)

    original = AppState(
        remote_jobs={
            "job-1": {
                "job_id": "job-1",
                "job_type": "backtest",
                "submitted_at": "2026-01-01T00:00:00Z",
                "last_known_state": "running",
                "server_host": "quant-host",
                "summary_cache_path": "/tmp/job-1-summary.json",
            }
        }
    )
    original.save()
    loaded = AppState.load()
    assert loaded.remote_jobs["job-1"]["server_host"] == "quant-host"
    assert loaded.remote_jobs["job-1"]["summary_cache_path"] == "/tmp/job-1-summary.json"

    legacy_payload = {
        "active_jobs": {
            "job-9": {
                "job_id": "job-9",
                "job_type": "backtest",
                "status": "completed",
                "submitted_at": "2026-01-02T00:00:00Z",
                "server_hostname": "legacy-host",
            }
        }
    }
    app_state_path.write_text(json.dumps(legacy_payload), encoding="utf-8")
    legacy_loaded = AppState.load()
    assert legacy_loaded.remote_jobs["job-9"]["last_known_state"] == "completed"
    assert legacy_loaded.remote_jobs["job-9"]["server_host"] == "legacy-host"

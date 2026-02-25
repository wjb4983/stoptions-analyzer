from __future__ import annotations

import json

import state
from state import AppState


def test_app_state_save_and_load_roundtrip_includes_regime_keys(tmp_path, monkeypatch) -> None:
    app_state_path = tmp_path / "app_state.json"
    monkeypatch.setattr(state, "STATE_PATH", app_state_path)

    original = AppState(
        tickers=["AAPL", "MSFT"],
        selected_ticker="AAPL",
        regime_definitions={"bull": {"label": "Bull", "confidence_thresholds": {"entry": 0.7}}},
        regime_training_runs=[{"id": "run-1", "regime_id": "bull", "score": 0.83}],
        active_regime_id="bull",
    )
    original.save()

    payload = json.loads(app_state_path.read_text(encoding="utf-8"))
    assert payload["regime_definitions"] == original.regime_definitions
    assert payload["regime_training_runs"] == original.regime_training_runs
    assert payload["active_regime_id"] == original.active_regime_id

    loaded = AppState.load()
    assert loaded.regime_definitions == original.regime_definitions
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
    assert loaded.regime_training_runs == []
    assert loaded.active_regime_id is None

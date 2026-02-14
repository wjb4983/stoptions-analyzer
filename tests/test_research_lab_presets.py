from __future__ import annotations

import json

import ui.research_lab_page as research_lab_page
from ui.research_lab_page import DEFAULT_RESEARCH_WORKFLOW_PRESETS, ResearchLabPage
from ui.workflow_preset_validator import validate_workflow_preset_payload


def test_load_workflow_presets_uses_repo_config(monkeypatch, tmp_path):
    preset_path = tmp_path / "research_lab_presets.json"
    payload = {
        "default_preset": "fast_iteration",
        "presets": {
            "fast_iteration": {
                "entry_signals": ["ts_momentum"],
                "exit_signals": ["none"],
                "optimization": {"n_trials": 5, "sampler": "random"},
                "walk_forward": {
                    "train_fraction": 0.6,
                    "validation_fraction": 0.2,
                    "test_fraction": 0.2,
                    "step_fraction": 0.1,
                    "split_policy": "volatility-regime-stratified",
                },
                "stress_controls": {"enable_historical_replay_regimes": False},
            }
        },
    }
    preset_path.write_text(json.dumps(payload), encoding="utf-8")
    monkeypatch.setattr(research_lab_page, "RESEARCH_LAB_PRESETS_PATH", preset_path)

    page = object.__new__(ResearchLabPage)
    loaded = ResearchLabPage._load_workflow_presets(page)

    assert loaded["default_preset"] == "fast_iteration"
    assert loaded["presets"]["fast_iteration"]["optimization"]["n_trials"] == 5
    assert loaded["presets"]["fast_iteration"]["walk_forward"]["split_policy"] == "volatility-regime-stratified"


def test_load_workflow_presets_falls_back_on_invalid_payload(monkeypatch, tmp_path):
    preset_path = tmp_path / "research_lab_presets.json"
    preset_path.write_text("[]", encoding="utf-8")
    monkeypatch.setattr(research_lab_page, "RESEARCH_LAB_PRESETS_PATH", preset_path)

    page = object.__new__(ResearchLabPage)
    loaded = ResearchLabPage._load_workflow_presets(page)

    assert loaded["default_preset"] == DEFAULT_RESEARCH_WORKFLOW_PRESETS["default_preset"]
    assert "balanced_baseline" in loaded["presets"]
    assert any("JSON object" in warning for warning in page._workflow_preset_warnings)


def test_workflow_preset_validation_migrates_legacy_sampler_and_bounds():
    payload = {
        "default_preset": "legacy",
        "presets": {
            "legacy": {
                "optimization": {"sampler": "bayesian"},
                "walk_forward": {
                    "train_fraction": 1.5,
                    "validation_fraction": 0.2,
                    "test_fraction": 0.2,
                    "step_fraction": 0.1,
                },
                "stress_controls": {
                    "historical_window_fraction": -0.2,
                    "overlay_liquidity_multiplier": 2.0,
                },
            }
        },
    }

    result = validate_workflow_preset_payload(payload, fallback_payload=DEFAULT_RESEARCH_WORKFLOW_PRESETS)

    preset = result.payload["presets"]["legacy"]
    assert preset["optimization"]["sampler"] == "tpe"
    assert preset["walk_forward"]["train_fraction"] == 0.70
    assert preset["stress_controls"]["historical_window_fraction"] == 0.20
    assert preset["stress_controls"]["overlay_liquidity_multiplier"] == 0.4
    assert any("migrated optimization.sampler 'bayesian' -> 'tpe'" in warning for warning in result.warnings)


def test_format_preset_warning_text_summarizes_entries():
    page = object.__new__(ResearchLabPage)
    page._workflow_preset_warnings = [f"w{i}" for i in range(5)]

    text = ResearchLabPage._format_preset_warning_text(page)

    assert "w0" in text
    assert "w2" in text
    assert "and 2 more warning(s)" in text

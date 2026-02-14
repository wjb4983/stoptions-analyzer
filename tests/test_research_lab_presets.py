from __future__ import annotations

import json

import ui.research_lab_page as research_lab_page
from ui.research_lab_page import DEFAULT_RESEARCH_WORKFLOW_PRESETS, ResearchLabPage


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

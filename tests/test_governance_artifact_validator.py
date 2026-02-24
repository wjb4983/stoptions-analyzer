from __future__ import annotations

import json
from pathlib import Path

import pytest

from reports import validate_governance_artifacts


EXPECTED_FILES = {
    "quality_gates_report.json",
    "model_card.json",
    "strategy_card.json",
    "benchmark_scorecard.json",
    "baseline_metrics_summary.json",
    "volatility_strategy_realism_report.json",
    "baseline_metrics_summary.csv",
    "equity_drawdown_plot_data.csv",
    "pass_fail_checklist.md",
}


def test_validator_passes_with_repository_artifacts() -> None:
    validate_governance_artifacts.validate_governance_artifacts()


def test_validator_detects_malformed_json(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    for filename in EXPECTED_FILES:
        src = Path("reports") / filename
        (tmp_path / filename).write_bytes(src.read_bytes())

    (tmp_path / "model_card.json").write_text(json.dumps({}), encoding="utf-8")

    monkeypatch.setattr(validate_governance_artifacts, "REPORTS_DIR", tmp_path)
    monkeypatch.setattr(
        validate_governance_artifacts,
        "JSON_ARTIFACTS",
        [tmp_path / "quality_gates_report.json", tmp_path / "model_card.json", tmp_path / "strategy_card.json", tmp_path / "benchmark_scorecard.json", tmp_path / "baseline_metrics_summary.json"],
    )
    monkeypatch.setattr(
        validate_governance_artifacts,
        "CSV_ARTIFACTS",
        [tmp_path / "baseline_metrics_summary.csv", tmp_path / "equity_drawdown_plot_data.csv"],
    )
    monkeypatch.setattr(
        validate_governance_artifacts,
        "TEXT_ARTIFACTS",
        [tmp_path / "pass_fail_checklist.md"],
    )

    with pytest.raises(RuntimeError, match="empty or malformed"):
        validate_governance_artifacts.validate_governance_artifacts()

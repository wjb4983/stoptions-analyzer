from __future__ import annotations

import csv
import json
from pathlib import Path
from typing import Any

REPORTS_DIR = Path(__file__).resolve().parent

JSON_ARTIFACTS = [
    REPORTS_DIR / "quality_gates_report.json",
    REPORTS_DIR / "model_card.json",
    REPORTS_DIR / "strategy_card.json",
    REPORTS_DIR / "benchmark_scorecard.json",
    REPORTS_DIR / "baseline_metrics_summary.json",
    REPORTS_DIR / "volatility_strategy_realism_report.json",
]

CSV_ARTIFACTS = [
    REPORTS_DIR / "baseline_metrics_summary.csv",
    REPORTS_DIR / "equity_drawdown_plot_data.csv",
]

TEXT_ARTIFACTS = [
    REPORTS_DIR / "pass_fail_checklist.md",
]


def _assert_file_exists(path: Path) -> None:
    if not path.exists() or not path.is_file():
        raise RuntimeError(f"Missing required governance artifact: {path}")


def _load_json(path: Path) -> Any:
    _assert_file_exists(path)
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def _validate_json_artifacts() -> None:
    for path in JSON_ARTIFACTS:
        payload = _load_json(path)
        if payload in ({}, [], None):
            raise RuntimeError(f"Governance artifact is empty or malformed: {path}")


def _validate_csv_artifacts() -> None:
    for path in CSV_ARTIFACTS:
        _assert_file_exists(path)
        with path.open("r", encoding="utf-8", newline="") as handle:
            rows = list(csv.DictReader(handle))
        if not rows:
            raise RuntimeError(f"Benchmark artifact has no rows: {path}")


def _validate_text_artifacts() -> None:
    for path in TEXT_ARTIFACTS:
        _assert_file_exists(path)
        if not path.read_text(encoding="utf-8").strip():
            raise RuntimeError(f"Governance artifact is empty: {path}")


def validate_governance_artifacts() -> None:
    _validate_json_artifacts()
    _validate_csv_artifacts()
    _validate_text_artifacts()


if __name__ == "__main__":
    validate_governance_artifacts()
    print("Governance artifact validation passed.")

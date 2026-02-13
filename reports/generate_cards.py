from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


REPORTS_DIR = Path(__file__).resolve().parent
DEFAULT_MODEL_CARD_PATH = REPORTS_DIR / "model_card.json"
DEFAULT_STRATEGY_CARD_PATH = REPORTS_DIR / "strategy_card.json"


def _load_json(path: Path) -> Any:
    return json.loads(path.read_text())


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def build_model_card(*, scorecard: dict[str, Any], quality_gates: dict[str, Any]) -> dict[str, Any]:
    return {
        "card_version": "1.0",
        "generated_at": _utc_now(),
        "model_name": "stoptions-vectorized-backtest-governance-model",
        "intended_use": "Promotion gating for systematic strategy deployment readiness.",
        "validation_summary": {
            "promotion_gate_pass": bool(scorecard.get("promotion_gate", {}).get("pass", False)),
            "failed_critical_dimensions": scorecard.get("promotion_gate", {}).get("failed_critical_dimensions", []),
        },
        "quality_gates": quality_gates.get("gates", {}),
        "deployment_readiness": bool(quality_gates.get("all_gates_pass", False)),
        "limitations": [
            "Synthetic benchmark data does not replace live production monitoring.",
            "Calibration coverage depends on available historical fills.",
        ],
    }


def build_strategy_card(*, baseline_rows: list[dict[str, Any]], quality_gates: dict[str, Any]) -> dict[str, Any]:
    scenarios = [str(row.get("scenario", "unknown")) for row in baseline_rows]
    return {
        "card_version": "1.0",
        "generated_at": _utc_now(),
        "strategy_name": "time-series-momentum-baseline-suite",
        "scenarios_evaluated": scenarios,
        "performance_snapshot": baseline_rows,
        "friction_adjusted_performance_gate": quality_gates.get("gates", {}).get("friction_adjusted_performance", {}),
        "merge_deploy_gates_passed": bool(quality_gates.get("all_gates_pass", False)),
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate standardized model and strategy card artifacts.")
    parser.add_argument("--scorecard", default=str(REPORTS_DIR / "benchmark_scorecard.json"))
    parser.add_argument("--quality-gates", default=str(REPORTS_DIR / "quality_gates_report.json"))
    parser.add_argument("--baseline", default=str(REPORTS_DIR / "baseline_metrics_summary.json"))
    parser.add_argument("--model-card-out", default=str(DEFAULT_MODEL_CARD_PATH))
    parser.add_argument("--strategy-card-out", default=str(DEFAULT_STRATEGY_CARD_PATH))
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    scorecard = _load_json(Path(args.scorecard))
    quality_gates = _load_json(Path(args.quality_gates))
    baseline_rows = _load_json(Path(args.baseline))

    model_card = build_model_card(scorecard=scorecard, quality_gates=quality_gates)
    strategy_card = build_strategy_card(baseline_rows=baseline_rows, quality_gates=quality_gates)

    Path(args.model_card_out).write_text(json.dumps(model_card, indent=2) + "\n")
    Path(args.strategy_card_out).write_text(json.dumps(strategy_card, indent=2) + "\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

from __future__ import annotations

import argparse
import json
import importlib.util
import sys
from pathlib import Path
from typing import Any

import numpy as np

ROOT_DIR = Path(__file__).resolve().parents[1]
QUALITY_GATES_MODULE_PATH = ROOT_DIR / "src" / "modeling_nextgen" / "validation" / "quality_gates.py"
_spec = importlib.util.spec_from_file_location("modeling_nextgen_quality_gates", QUALITY_GATES_MODULE_PATH)
if _spec is None or _spec.loader is None:
    raise ImportError(f"Unable to load quality gate module from {QUALITY_GATES_MODULE_PATH}")
_quality_gates_module = importlib.util.module_from_spec(_spec)
sys.modules[_spec.name] = _quality_gates_module
_spec.loader.exec_module(_quality_gates_module)
evaluate_modeling_quality_gates = _quality_gates_module.evaluate_modeling_quality_gates
promotion_blocked = _quality_gates_module.promotion_blocked


REPORTS_DIR = Path(__file__).resolve().parent
DEFAULT_SCORECARD_PATH = REPORTS_DIR / "benchmark_scorecard.json"
DEFAULT_CALIBRATION_PATH = REPORTS_DIR / "calibration_report.json"
DEFAULT_BASELINE_PATH = REPORTS_DIR / "baseline_metrics_summary.json"
DEFAULT_NO_ARB_PATH = REPORTS_DIR / "no_arb_diagnostics_report.json"
DEFAULT_OUTPUT_PATH = REPORTS_DIR / "quality_gates_report.json"



def _load_json(path: Path) -> Any:
    return json.loads(path.read_text())



def _bool_gate(name: str, passed: bool, details: dict[str, Any]) -> dict[str, Any]:
    return {"name": name, "pass": bool(passed), "details": details}



def evaluate_quality_gates(
    *,
    scorecard: dict[str, Any],
    calibration_report: dict[str, Any],
    baseline_rows: list[dict[str, Any]],
    no_arb_report: dict[str, Any] | None = None,
) -> dict[str, Any]:
    baseline = list(baseline_rows)
    if not baseline:
        raise ValueError("baseline_rows cannot be empty")

    sharpe_values = np.asarray([float(row.get("sharpe", np.nan)) for row in baseline], dtype=float)
    turnover_values = np.asarray([float(row.get("turnover_total", np.nan)) for row in baseline], dtype=float)

    data_quality = _bool_gate(
        "data_quality",
        bool(np.all(np.isfinite(sharpe_values)) and np.all(np.isfinite(turnover_values))),
        {
            "scenarios": len(baseline),
            "all_sharpe_finite": bool(np.all(np.isfinite(sharpe_values))),
            "all_turnover_finite": bool(np.all(np.isfinite(turnover_values))),
        },
    )

    leakage = scorecard.get("dimensions", {}).get("reproducibility", {})
    leakage_checks = leakage.get("checks", {}) if isinstance(leakage, dict) else {}
    hash_consistency = float(leakage_checks.get("hash_consistency_ratio", {}).get("value", 0.0))
    leakage_tests = _bool_gate(
        "leakage_tests",
        hash_consistency >= 0.99,
        {
            "hash_consistency_ratio": hash_consistency,
            "source_dimension": "reproducibility",
        },
    )

    validation_integrity = _bool_gate(
        "validation_integrity",
        bool(scorecard.get("promotion_gate", {}).get("pass", False)),
        {
            "failed_critical_dimensions": list(scorecard.get("promotion_gate", {}).get("failed_critical_dimensions", [])),
        },
    )

    nextgen_gates = evaluate_modeling_quality_gates(
        scorecard=scorecard,
        calibration_report=calibration_report,
        baseline_rows=baseline,
    )

    no_arb_model_gate = no_arb_report.get("model_gate", {}) if isinstance(no_arb_report, dict) else {}
    no_arb_gate = _bool_gate(
        "no_arbitrage_surface",
        bool(no_arb_model_gate.get("pass", True)),
        {
            "threshold": int(no_arb_model_gate.get("threshold", 0)),
            "diagnostics": no_arb_report.get("diagnostics", {}) if isinstance(no_arb_report, dict) else {},
        },
    )

    gates = [data_quality, leakage_tests, validation_integrity, *nextgen_gates, no_arb_gate]
    all_pass = not promotion_blocked(gates)

    return {
        "all_gates_pass": all_pass,
        "required_for_merge_and_deploy": True,
        "gates": {gate["name"]: gate for gate in gates},
    }



def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Evaluate required quality gates for merge/deploy.")
    parser.add_argument("--scorecard", default=str(DEFAULT_SCORECARD_PATH))
    parser.add_argument("--calibration", default=str(DEFAULT_CALIBRATION_PATH))
    parser.add_argument("--baseline", default=str(DEFAULT_BASELINE_PATH))
    parser.add_argument("--no-arb", default=str(DEFAULT_NO_ARB_PATH))
    parser.add_argument("--out", default=str(DEFAULT_OUTPUT_PATH))
    return parser.parse_args()



def main() -> int:
    args = parse_args()
    result = evaluate_quality_gates(
        scorecard=_load_json(Path(args.scorecard)),
        calibration_report=_load_json(Path(args.calibration)),
        baseline_rows=_load_json(Path(args.baseline)),
        no_arb_report=_load_json(Path(args.no_arb)),
    )
    out_path = Path(args.out)
    out_path.write_text(json.dumps(result, indent=2) + "\n")
    return 0 if result["all_gates_pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())

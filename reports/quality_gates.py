from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

import numpy as np


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

    fit_error = calibration_report.get("fit_error", {}) if isinstance(calibration_report.get("fit_error"), dict) else {}
    stability = calibration_report.get("stability", {}) if isinstance(calibration_report.get("stability"), dict) else {}
    mae_max = float(fit_error.get("mae_bps_max", np.inf))
    coeff_std = float(stability.get("impact_coefficient_std_bps", np.inf))
    calibration = _bool_gate(
        "calibration",
        bool(mae_max <= 15.0 and coeff_std <= 20.0),
        {
            "mae_bps_max": mae_max,
            "impact_coefficient_std_bps": coeff_std,
            "thresholds": {"mae_bps_max": 15.0, "impact_coefficient_std_bps": 20.0},
        },
    )

    friction_adjusted = _bool_gate(
        "friction_adjusted_performance",
        bool(np.all(np.isfinite(sharpe_values)) and np.nanmean(turnover_values) >= 0.0),
        {
            "mean_sharpe": float(np.nanmean(sharpe_values)),
            "mean_turnover": float(np.nanmean(turnover_values)),
        },
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

    gates = [data_quality, leakage_tests, validation_integrity, calibration, friction_adjusted, no_arb_gate]
    all_pass = all(bool(gate["pass"]) for gate in gates)

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

from __future__ import annotations

import json
from pathlib import Path

import pytest

from reports import benchmark_bundle, quality_gates, run_baselines


REPORTS_DIR = Path("reports")
EXPECTED_RANGES_PATH = REPORTS_DIR / "benchmark_expected_ranges.json"
BASELINE_SUMMARY_PATH = REPORTS_DIR / "baseline_metrics_summary.json"
BASELINE_CONFIG_PATH = REPORTS_DIR / "baseline_configs.json"
CALIBRATION_REPORT_PATH = REPORTS_DIR / "calibration_report.json"
CALIBRATION_SNAPSHOTS_PATH = REPORTS_DIR / "slippage_calibration_snapshots.json"
VOLATILITY_STRATEGY_REALISM_PATH = REPORTS_DIR / "volatility_strategy_realism_report.json"

RUNTIME_TOLERANCE_SECONDS = 0.15
THROUGHPUT_FLOOR_RATIO = 0.80
CALIBRATION_PARAM_DRIFT_BOUNDS = {
    "impact_coefficient_bps": 1.0,
    "base_bps": 0.10,
    "participation_exponent": 0.05,
    "max_participation": 0.05,
}


def _load_json(path: Path):
    return json.loads(path.read_text())


def _bound_delta(value: float, lower: float, upper: float) -> float:
    if value < lower:
        return lower - value
    if value > upper:
        return value - upper
    return 0.0


def test_generated_benchmark_metrics_remain_within_configured_ranges() -> None:
    expected_ranges = _load_json(EXPECTED_RANGES_PATH)
    scorecard = benchmark_bundle.build_benchmark_scorecard(expected_ranges=expected_ranges)

    breaches: list[str] = []
    for dimension_name, dimension in scorecard["dimensions"].items():
        for metric_name, check in dimension["checks"].items():
            value = float(check["value"])
            lower, upper = check["expected_range"]
            delta = _bound_delta(value=value, lower=float(lower), upper=float(upper))
            if not bool(check["pass"]):
                breaches.append(
                    (
                        f"{dimension_name}.{metric_name} breached range "
                        f"[{lower:.6f}, {upper:.6f}] with value={value:.6f}; "
                        f"breach_delta={delta:.6f}"
                    )
                )

    assert not breaches, "\n".join(["Benchmark range breaches:", *breaches])


def test_pr_guard_runtime_and_throughput_ceilings_from_baseline_summary() -> None:
    baseline_rows = _load_json(BASELINE_SUMMARY_PATH)
    baseline_config = _load_json(BASELINE_CONFIG_PATH)
    workload_size = int(baseline_config["target_universe_assets"]) * int(baseline_config["target_horizon_periods"])

    runtime_ceiling = float(run_baselines.RUNTIME_THRESHOLD_SECONDS) + RUNTIME_TOLERANCE_SECONDS
    expected_floor_throughput = workload_size / runtime_ceiling

    breaches: list[str] = []
    for row in baseline_rows:
        scenario = row["scenario"]
        runtime_seconds = float(row["runtime_seconds"])
        throughput = workload_size / runtime_seconds if runtime_seconds > 0.0 else 0.0

        if runtime_seconds > runtime_ceiling:
            breaches.append(
                (
                    f"{scenario}.runtime_seconds exceeded ceiling "
                    f"{runtime_ceiling:.6f}s with {runtime_seconds:.6f}s; "
                    f"breach_delta={runtime_seconds - runtime_ceiling:.6f}s"
                )
            )

        min_allowed = expected_floor_throughput * THROUGHPUT_FLOOR_RATIO
        if throughput < min_allowed:
            breaches.append(
                (
                    f"{scenario}.throughput below floor {min_allowed:.2f} period-assets/s "
                    f"with {throughput:.2f}; breach_delta={min_allowed - throughput:.2f}"
                )
            )

    assert not breaches, "\n".join(["Runtime/throughput breaches:", *breaches])


def test_quality_gate_score_thresholds_pass_for_current_artifacts() -> None:
    expected_ranges = _load_json(EXPECTED_RANGES_PATH)
    scorecard = benchmark_bundle.build_benchmark_scorecard(expected_ranges=expected_ranges)
    calibration_report = _load_json(CALIBRATION_REPORT_PATH)
    baseline_rows = _load_json(BASELINE_SUMMARY_PATH)
    volatility_strategy_report = _load_json(VOLATILITY_STRATEGY_REALISM_PATH)

    result = quality_gates.evaluate_quality_gates(
        scorecard=scorecard,
        calibration_report=calibration_report,
        baseline_rows=baseline_rows,
        volatility_strategy_report=volatility_strategy_report,
    )

    failures = [
        f"{name}: {gate['details']}"
        for name, gate in result["gates"].items()
        if not bool(gate["pass"])
    ]
    assert result["all_gates_pass"], "\n".join(["Quality gate failures:", *failures])


def test_calibration_snapshot_drift_within_bounds() -> None:
    payload = _load_json(CALIBRATION_SNAPSHOTS_PATH)
    default_params = payload["default_params"]
    snapshots = payload["snapshots"]

    breaches: list[str] = []
    for snapshot in snapshots:
        date = snapshot["effective_date"]
        params = snapshot["params"]
        for param_name, max_abs_drift in CALIBRATION_PARAM_DRIFT_BOUNDS.items():
            baseline_value = float(default_params[param_name])
            snapshot_value = float(params[param_name])
            abs_drift = abs(snapshot_value - baseline_value)
            if abs_drift > max_abs_drift:
                breaches.append(
                    (
                        f"snapshot[{date}].{param_name} drift {abs_drift:.6f} exceeded "
                        f"bound {max_abs_drift:.6f}; breach_delta={abs_drift - max_abs_drift:.6f}"
                    )
                )

    assert not breaches, "\n".join(["Calibration drift breaches:", *breaches])


@pytest.mark.slow
def test_slow_regression_rebuild_baselines_within_tolerances(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    expected_rows = _load_json(BASELINE_SUMMARY_PATH)
    expected_by_scenario = {row["scenario"]: row for row in expected_rows}

    monkeypatch.setattr(run_baselines, "OUTPUT_DIR", tmp_path)
    run_baselines.main()

    rebuilt_rows = _load_json(tmp_path / "baseline_metrics_summary.json")
    runtime_ceiling = float(run_baselines.RUNTIME_THRESHOLD_SECONDS) + RUNTIME_TOLERANCE_SECONDS

    breaches: list[str] = []
    for row in rebuilt_rows:
        scenario = row["scenario"]
        expected = expected_by_scenario[scenario]

        runtime = float(row["runtime_seconds"])
        if runtime > runtime_ceiling:
            breaches.append(
                (
                    f"{scenario}.runtime_seconds exceeded ceiling {runtime_ceiling:.6f}s "
                    f"with {runtime:.6f}s; breach_delta={runtime - runtime_ceiling:.6f}s"
                )
            )

        for metric_name, tolerance in {
            "sharpe": 0.15,
            "max_drawdown": 0.02,
            "turnover_total": 6.0,
        }.items():
            value = float(row[metric_name])
            baseline_value = float(expected[metric_name])
            abs_diff = abs(value - baseline_value)
            if abs_diff > tolerance:
                breaches.append(
                    (
                        f"{scenario}.{metric_name} drifted from baseline {baseline_value:.6f} "
                        f"to {value:.6f}; abs_diff={abs_diff:.6f} > tolerance={tolerance:.6f}; "
                        f"breach_delta={abs_diff - tolerance:.6f}"
                    )
                )

    assert not breaches, "\n".join(["Slow baseline regression breaches:", *breaches])

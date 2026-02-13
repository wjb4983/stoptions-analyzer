from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any

import numpy as np

REPORTS_DIR = Path(__file__).resolve().parent
DATASET_PATH = REPORTS_DIR / "benchmark_bundle_dataset.json"
RANGES_PATH = REPORTS_DIR / "benchmark_expected_ranges.json"
SCORECARD_PATH = REPORTS_DIR / "benchmark_scorecard.json"


def _load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text())


def _in_range(value: float, bounds: list[float]) -> bool:
    lower, upper = float(bounds[0]), float(bounds[1])
    return lower <= float(value) <= upper


def _evaluate_dimension(metrics: dict[str, float], spec: dict[str, Any]) -> dict[str, Any]:
    checks = {}
    for name, bounds in spec["ranges"].items():
        metric_value = float(metrics[name])
        checks[name] = {
            "value": metric_value,
            "expected_range": [float(bounds[0]), float(bounds[1])],
            "pass": _in_range(metric_value, bounds),
        }
    passed = all(item["pass"] for item in checks.values())
    return {
        "critical": bool(spec.get("critical", False)),
        "pass": passed,
        "checks": checks,
    }


def _compute_robust_oos_performance(dataset: dict[str, Any]) -> dict[str, float]:
    returns = np.asarray(dataset["oos_returns"], dtype=float)
    sharpe = float(np.mean(returns) / np.std(returns, ddof=1) * np.sqrt(252))
    equity = np.cumprod(1.0 + returns)
    running_peak = np.maximum.accumulate(equity)
    drawdown = equity / np.where(running_peak == 0.0, 1.0, running_peak) - 1.0
    max_drawdown = float(np.min(drawdown))
    hit_rate = float(np.mean(returns > 0.0))
    return {
        "annualized_sharpe": sharpe,
        "max_drawdown": max_drawdown,
        "hit_rate": hit_rate,
    }


def _compute_statistical_significance(dataset: dict[str, Any]) -> dict[str, float]:
    returns = np.asarray(dataset["oos_returns"], dtype=float)
    mean_return_tstat = float(np.mean(returns) / (np.std(returns, ddof=1) / np.sqrt(returns.size)))

    candidates = np.asarray(dataset["candidate_returns"], dtype=float)
    candidate_means = candidates.mean(axis=0)
    observed = float(np.max(candidate_means))
    centered = candidates - candidate_means
    rng = np.random.default_rng(123)
    bootstrap_maxima = []
    for _ in range(400):
        sample_idx = rng.integers(0, centered.shape[0], size=centered.shape[0])
        sampled = centered[sample_idx, :]
        bootstrap_maxima.append(float(np.max(sampled.mean(axis=0))))
    bootstrap_maxima_arr = np.asarray(bootstrap_maxima, dtype=float)
    bootstrap_p_value = float(np.mean(bootstrap_maxima_arr >= observed))

    return {
        "mean_return_tstat": mean_return_tstat,
        "bootstrap_p_value": bootstrap_p_value,
    }


def _compute_execution_realism(dataset: dict[str, Any]) -> dict[str, float]:
    execution = dataset["execution"]
    participation = np.asarray(execution["participation_rates"], dtype=float)
    realized = np.asarray(execution["realized_slippage_bps"], dtype=float)
    expected = np.asarray(execution["expected_slippage_bps"], dtype=float)
    fills = np.asarray(execution["fill_ratios"], dtype=float)
    tracking_error = float(np.mean(np.abs(realized - expected)))
    return {
        "avg_participation_rate": float(np.mean(participation)),
        "slippage_tracking_error_bps": tracking_error,
        "avg_fill_ratio": float(np.mean(fills)),
    }


def _compute_stress_resilience(dataset: dict[str, Any]) -> dict[str, float]:
    scenario_returns = np.asarray(list(dataset["stress"]["scenario_returns"].values()), dtype=float)
    recovery = np.asarray(dataset["stress"]["post_stress_recovery"], dtype=float)
    return {
        "worst_stress_return": float(np.min(scenario_returns)),
        "recovery_mean": float(np.mean(recovery)),
    }


def _snapshot_hash(snapshot: dict[str, float]) -> str:
    payload = json.dumps(snapshot, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def _compute_reproducibility(dataset: dict[str, Any]) -> dict[str, float]:
    snapshots = dataset["reproducibility"]["run_metric_snapshots"]
    hashes = [_snapshot_hash(snapshot) for snapshot in snapshots]
    first = hashes[0]
    hash_consistency_ratio = float(sum(1 for value in hashes if value == first) / len(hashes))

    values = np.asarray([[float(snapshot[k]) for k in sorted(snapshot.keys())] for snapshot in snapshots], dtype=float)
    metric_variance = float(np.max(np.var(values, axis=0)))
    return {
        "hash_consistency_ratio": hash_consistency_ratio,
        "metric_variance": metric_variance,
    }


def build_benchmark_scorecard(
    dataset: dict[str, Any] | None = None,
    expected_ranges: dict[str, Any] | None = None,
) -> dict[str, Any]:
    dataset_payload = dataset or _load_json(DATASET_PATH)
    ranges_payload = expected_ranges or _load_json(RANGES_PATH)

    computed_metrics = {
        "robust_oos_performance": _compute_robust_oos_performance(dataset_payload),
        "statistical_significance": _compute_statistical_significance(dataset_payload),
        "execution_realism": _compute_execution_realism(dataset_payload),
        "stress_resilience": _compute_stress_resilience(dataset_payload),
        "reproducibility": _compute_reproducibility(dataset_payload),
    }

    dimensions = {
        name: _evaluate_dimension(computed_metrics[name], ranges_payload[name])
        for name in [
            "robust_oos_performance",
            "statistical_significance",
            "execution_realism",
            "stress_resilience",
            "reproducibility",
        ]
    }

    critical_failures = [name for name, result in dimensions.items() if result["critical"] and not result["pass"]]
    promotion_gate = {
        "pass": len(critical_failures) == 0,
        "failed_critical_dimensions": critical_failures,
    }

    return {
        "dataset_path": str(DATASET_PATH),
        "expected_ranges_path": str(RANGES_PATH),
        "dimensions": dimensions,
        "promotion_gate": promotion_gate,
    }


def write_scorecard(scorecard: dict[str, Any], output_path: Path = SCORECARD_PATH) -> Path:
    output_path.write_text(json.dumps(scorecard, indent=2) + "\n")
    return output_path


def main() -> int:
    scorecard = build_benchmark_scorecard()
    write_scorecard(scorecard)
    return 0 if scorecard["promotion_gate"]["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())

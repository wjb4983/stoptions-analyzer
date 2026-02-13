from __future__ import annotations

import csv
import itertools
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

import numpy as np


@dataclass(frozen=True)
class WalkForwardFold:
    fold_id: int
    train_start: int
    train_end: int
    validation_start: int
    validation_end: int
    test_start: int
    test_end: int
    purge_window_bars: int = 0
    embargo_window_bars: int = 0
    label_horizon_bars: int = 1

    @property
    def excluded_ranges(self) -> dict[str, list[int]]:
        return {
            "train_validation": [self.train_end, self.validation_start],
            "validation_test": [self.validation_end, self.test_start],
        }

    @property
    def leakage_checks(self) -> dict[str, bool]:
        label_gap = max(1, int(self.label_horizon_bars))
        train_label_end = self.train_end + label_gap
        validation_label_end = self.validation_end + label_gap
        return {
            "train_validation_non_overlap": train_label_end <= self.validation_start,
            "validation_test_non_overlap": validation_label_end <= self.test_start,
            "train_test_non_overlap": train_label_end <= self.test_start,
        }


@dataclass
class WalkForwardResult:
    folds: list[dict[str, Any]]
    aggregate_metrics: dict[str, float]
    stability: dict[str, Any]
    audit: dict[str, Any] | None = None
    validation_report: dict[str, Any] | None = None


def build_walk_forward_folds(
    *,
    total_bars: int,
    train_bars: int,
    validation_bars: int,
    test_bars: int,
    step_bars: int | None = None,
    purge_window_bars: int = 0,
    embargo_window_bars: int = 0,
    label_horizon_bars: int = 1,
) -> list[WalkForwardFold]:
    if min(total_bars, train_bars, validation_bars, test_bars) <= 0:
        raise ValueError("window sizes and total_bars must be positive")
    if purge_window_bars < 0 or embargo_window_bars < 0:
        raise ValueError("purge_window_bars and embargo_window_bars must be non-negative")
    if label_horizon_bars <= 0:
        raise ValueError("label_horizon_bars must be positive")

    step = step_bars if step_bars is not None else test_bars
    if step <= 0:
        raise ValueError("step_bars must be positive")

    effective_purge = max(int(purge_window_bars), int(label_horizon_bars))
    effective_embargo = max(int(embargo_window_bars), int(label_horizon_bars))

    folds: list[WalkForwardFold] = []
    cursor = 0
    fold_id = 0
    while cursor + train_bars + effective_purge + validation_bars + effective_embargo + test_bars <= total_bars:
        train_start = cursor
        train_end = train_start + train_bars
        validation_start = train_end + effective_purge
        validation_end = validation_start + validation_bars
        test_start = validation_end + effective_embargo
        test_end = test_start + test_bars
        fold = WalkForwardFold(
            fold_id=fold_id,
            train_start=train_start,
            train_end=train_end,
            validation_start=validation_start,
            validation_end=validation_end,
            test_start=test_start,
            test_end=test_end,
            purge_window_bars=int(purge_window_bars),
            embargo_window_bars=int(embargo_window_bars),
            label_horizon_bars=int(label_horizon_bars),
        )
        if not all(fold.leakage_checks.values()):
            raise ValueError(f"Fold {fold_id} violates leakage constraints: {fold.leakage_checks}")
        folds.append(fold)
        fold_id += 1
        cursor += step
    return folds


def build_cpcv_walk_forward_folds(
    *,
    total_bars: int,
    n_groups: int,
    n_test_groups: int,
    purge_window_bars: int = 0,
    embargo_window_bars: int = 0,
    label_horizon_bars: int = 1,
) -> list[WalkForwardFold]:
    if total_bars <= 0:
        raise ValueError("total_bars must be positive")
    if n_groups < 3:
        raise ValueError("n_groups must be at least 3 for CPCV")
    if n_test_groups <= 0 or n_test_groups >= n_groups:
        raise ValueError("n_test_groups must be in [1, n_groups-1]")

    boundaries = np.linspace(0, total_bars, n_groups + 1, dtype=int)
    folds: list[WalkForwardFold] = []
    fold_id = 0

    for held_out in itertools.combinations(range(n_groups), n_test_groups):
        train_groups = [idx for idx in range(n_groups) if idx not in held_out]
        if not train_groups:
            continue
        for test_group in held_out:
            train_start = int(boundaries[train_groups[0]])
            train_end = int(boundaries[train_groups[-1] + 1])
            validation_start = int(boundaries[test_group])
            validation_end = int(boundaries[test_group + 1])
            test_start = validation_start
            test_end = validation_end

            fold = WalkForwardFold(
                fold_id=fold_id,
                train_start=train_start,
                train_end=train_end,
                validation_start=validation_start,
                validation_end=validation_end,
                test_start=test_start,
                test_end=test_end,
                purge_window_bars=int(purge_window_bars),
                embargo_window_bars=int(embargo_window_bars),
                label_horizon_bars=int(label_horizon_bars),
            )
            folds.append(fold)
            fold_id += 1
    return folds


def run_walk_forward_optimization(
    *,
    folds: list[WalkForwardFold],
    parameter_candidates: list[dict[str, Any]],
    evaluate_segment: Callable[[dict[str, Any], int, int], dict[str, Any]],
    score_metric: str = "sharpe",
    nested_inner_folds: dict[int, list[WalkForwardFold]] | None = None,
) -> WalkForwardResult:
    if not folds:
        raise ValueError("At least one fold is required")
    if not parameter_candidates:
        raise ValueError("At least one parameter candidate is required")

    fold_rows: list[dict[str, Any]] = []
    selected_keys: list[str] = []

    for fold in folds:
        diagnostics: list[dict[str, Any]] = []
        best_candidate: dict[str, Any] | None = None
        best_score = float("-inf")

        for candidate in parameter_candidates:
            train_eval = evaluate_segment(candidate, fold.train_start, fold.train_end)
            validation_eval = evaluate_segment(candidate, fold.validation_start, fold.validation_end)

            inner_scores: list[float] = []
            inner_details: list[dict[str, Any]] = []
            if nested_inner_folds and fold.fold_id in nested_inner_folds:
                for inner_fold in nested_inner_folds[fold.fold_id]:
                    evaluate_segment(candidate, inner_fold.train_start, inner_fold.train_end)
                    inner_validation_eval = evaluate_segment(
                        candidate,
                        inner_fold.validation_start,
                        inner_fold.validation_end,
                    )
                    inner_score = float(
                        inner_validation_eval.get("metrics", {}).get(score_metric, float("-inf"))
                    )
                    inner_scores.append(inner_score)
                    inner_details.append(
                        {
                            "indices": {
                                "train": [inner_fold.train_start, inner_fold.train_end],
                                "validation": [inner_fold.validation_start, inner_fold.validation_end],
                            },
                            "validation_score": inner_score,
                        }
                    )

            score = (
                float(np.mean(np.asarray(inner_scores, dtype=float)))
                if inner_scores
                else float(validation_eval.get("metrics", {}).get(score_metric, float("-inf")))
            )
            row = {
                "params": dict(candidate),
                "train_metrics": dict(train_eval.get("metrics", {})),
                "validation_metrics": dict(validation_eval.get("metrics", {})),
                "validation_score": score,
                "segments": {
                    "train": {
                        "start": fold.train_start,
                        "end": fold.train_end,
                        "output": dict(train_eval),
                    },
                    "validation": {
                        "start": fold.validation_start,
                        "end": fold.validation_end,
                        "output": dict(validation_eval),
                    },
                },
            }
            if inner_details:
                row["inner_diagnostics"] = inner_details
            diagnostics.append(row)

            tie_breaker = json.dumps(candidate, sort_keys=True)
            current_best = json.dumps(best_candidate, sort_keys=True) if best_candidate is not None else ""
            if score > best_score or (score == best_score and tie_breaker < current_best):
                best_score = score
                best_candidate = dict(candidate)

        if best_candidate is None:
            raise RuntimeError("No candidate selected for fold")

        oos_eval = evaluate_segment(best_candidate, fold.test_start, fold.test_end)
        selected_key = json.dumps(best_candidate, sort_keys=True)
        selected_keys.append(selected_key)

        fold_rows.append(
            {
                "fold_id": fold.fold_id,
                "indices": {
                    "train": [fold.train_start, fold.train_end],
                    "validation": [fold.validation_start, fold.validation_end],
                    "test": [fold.test_start, fold.test_end],
                },
                "excluded_ranges": fold.excluded_ranges,
                "leakage_checks": fold.leakage_checks,
                "windows": {
                    "purge_window_bars": fold.purge_window_bars,
                    "embargo_window_bars": fold.embargo_window_bars,
                    "label_horizon_bars": fold.label_horizon_bars,
                },
                "selected_params": best_candidate,
                "validation_score": best_score,
                "oos_metrics": dict(oos_eval.get("metrics", {})),
                "oos_equity": list(oos_eval.get("equity", [])),
                "oos_output": dict(oos_eval),
                "diagnostics": diagnostics,
            }
        )

    aggregate_metrics = _aggregate_numeric_metrics([row["oos_metrics"] for row in fold_rows])
    stability = _build_stability_summary(selected_keys, fold_rows)
    validation_report = _build_validation_report(fold_rows=fold_rows, stability=stability, score_metric=score_metric)
    return WalkForwardResult(
        folds=fold_rows,
        aggregate_metrics=aggregate_metrics,
        stability=stability,
        audit={"n_candidates": len(parameter_candidates), "n_folds": len(folds), "score_metric": score_metric},
        validation_report=validation_report,
    )


def persist_walk_forward_outputs(*, run_dir: Path, result: WalkForwardResult) -> None:
    run_dir.mkdir(parents=True, exist_ok=True)
    folds_dir = run_dir / "folds"
    folds_dir.mkdir(parents=True, exist_ok=True)

    for fold_row in result.folds:
        fold_id = int(fold_row["fold_id"])
        fold_dir = folds_dir / f"fold_{fold_id:03d}"
        fold_dir.mkdir(parents=True, exist_ok=True)

        (fold_dir / "selected_params.json").write_text(json.dumps(fold_row["selected_params"], indent=2))
        (fold_dir / "oos_metrics.json").write_text(json.dumps(fold_row["oos_metrics"], indent=2))
        (fold_dir / "oos_output.json").write_text(json.dumps(fold_row.get("oos_output", {}), indent=2))
        (fold_dir / "diagnostics.json").write_text(json.dumps(fold_row["diagnostics"], indent=2))
        (fold_dir / "fold_full_record.json").write_text(json.dumps(fold_row, indent=2))

        with (fold_dir / "oos_equity.csv").open("w", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=["timestamp", "equity"])
            writer.writeheader()
            for row in fold_row["oos_equity"]:
                writer.writerow({"timestamp": row.get("timestamp"), "equity": row.get("equity")})

    (run_dir / "fold_summary.json").write_text(json.dumps(result.folds, indent=2))
    (run_dir / "aggregate_metrics.json").write_text(json.dumps(result.aggregate_metrics, indent=2))
    (run_dir / "stability.json").write_text(json.dumps(result.stability, indent=2))
    (run_dir / "audit.json").write_text(json.dumps(result.audit or {}, indent=2))
    validation_dir = run_dir / "validation"
    validation_dir.mkdir(parents=True, exist_ok=True)
    validation_payload = result.validation_report or _build_validation_report(
        fold_rows=result.folds,
        stability=result.stability,
        score_metric=str((result.audit or {}).get("score_metric", "sharpe")),
    )
    (validation_dir / "report.json").write_text(json.dumps(validation_payload, indent=2))
    (validation_dir / "report.txt").write_text(_format_validation_report_text(validation_payload))


def _aggregate_numeric_metrics(metrics_rows: list[dict[str, Any]]) -> dict[str, float]:
    all_keys = sorted({key for row in metrics_rows for key, value in row.items() if isinstance(value, (int, float))})
    aggregated: dict[str, float] = {}
    for key in all_keys:
        values = np.asarray([float(row[key]) for row in metrics_rows if key in row], dtype=float)
        if values.size == 0:
            continue
        aggregated[f"{key}_mean"] = float(np.mean(values))
        aggregated[f"{key}_std"] = float(np.std(values, ddof=0))
        aggregated[f"{key}_min"] = float(np.min(values))
        aggregated[f"{key}_max"] = float(np.max(values))
    return aggregated


def _build_stability_summary(selected_keys: list[str], fold_rows: list[dict[str, Any]]) -> dict[str, Any]:
    selection_counts: dict[str, int] = {}
    for key in selected_keys:
        selection_counts[key] = selection_counts.get(key, 0) + 1

    validation_scores = [float(row["validation_score"]) for row in fold_rows]
    return {
        "selection_counts": selection_counts,
        "unique_selected_params": len(selection_counts),
        "validation_score_mean": float(np.mean(validation_scores)) if validation_scores else 0.0,
        "validation_score_std": float(np.std(validation_scores, ddof=0)) if validation_scores else 0.0,
        "fold_reuse": _build_fold_reuse_diagnostics(fold_rows),
    }


def _build_validation_report(
    *,
    fold_rows: list[dict[str, Any]],
    stability: dict[str, Any],
    score_metric: str,
) -> dict[str, Any]:
    fold_scores: list[dict[str, float]] = []
    train_val_gaps: list[float] = []
    val_test_gaps: list[float] = []
    for row in fold_rows:
        oos_metrics = row.get("oos_metrics", {})
        diagnostics = row.get("diagnostics", [])
        best_diag = max(
            diagnostics,
            key=lambda item: float(item.get("validation_score", float("-inf"))),
            default=None,
        )
        train_metric = 0.0
        val_metric = 0.0
        if isinstance(best_diag, dict):
            train_metric = float(best_diag.get("train_metrics", {}).get(score_metric, 0.0))
            val_metric = float(best_diag.get("validation_metrics", {}).get(score_metric, 0.0))
        oos_metric = float(oos_metrics.get(score_metric, 0.0))
        fold_scores.append(
            {
                "fold_id": int(row.get("fold_id", -1)),
                "train": train_metric,
                "validation": val_metric,
                "test": oos_metric,
            }
        )
        train_val_gaps.append(train_metric - val_metric)
        val_test_gaps.append(val_metric - oos_metric)

    instability_flags: list[str] = []
    overfit_flags: list[str] = []
    validation_std = float(stability.get("validation_score_std", 0.0))
    if validation_std > 0.75:
        instability_flags.append("validation_score_variance_high")
    unique_params = int(stability.get("unique_selected_params", 0))
    if unique_params > max(2, len(fold_rows) // 2):
        instability_flags.append("parameter_churn_high")
    if train_val_gaps and float(np.mean(np.asarray(train_val_gaps, dtype=float))) > 0.2:
        overfit_flags.append("train_validation_gap_large")
    if val_test_gaps and float(np.mean(np.asarray(val_test_gaps, dtype=float))) > 0.2:
        overfit_flags.append("validation_test_decay_large")

    return {
        "summary": {
            "score_metric": score_metric,
            "fold_count": len(fold_rows),
            "validation_score_std": validation_std,
            "mean_train_validation_gap": float(np.mean(np.asarray(train_val_gaps, dtype=float))) if train_val_gaps else 0.0,
            "mean_validation_test_gap": float(np.mean(np.asarray(val_test_gaps, dtype=float))) if val_test_gaps else 0.0,
        },
        "flags": {
            "instability": instability_flags,
            "overfitting": overfit_flags,
        },
        "per_fold_score_ladder": fold_scores,
    }


def _format_validation_report_text(report: dict[str, Any]) -> str:
    summary = report.get("summary", {})
    flags = report.get("flags", {})
    instability = flags.get("instability", []) if isinstance(flags, dict) else []
    overfitting = flags.get("overfitting", []) if isinstance(flags, dict) else []
    lines = [
        "Walk-Forward Validation Report",
        "==============================",
        f"Score metric: {summary.get('score_metric', 'sharpe')}",
        f"Fold count: {int(summary.get('fold_count', 0))}",
        f"Validation std: {float(summary.get('validation_score_std', 0.0)):.6f}",
        f"Mean train-validation gap: {float(summary.get('mean_train_validation_gap', 0.0)):.6f}",
        f"Mean validation-test gap: {float(summary.get('mean_validation_test_gap', 0.0)):.6f}",
        "",
        "Instability flags:",
    ]
    lines.extend([f"- {item}" for item in instability] or ["- none"])
    lines.append("Overfitting flags:")
    lines.extend([f"- {item}" for item in overfitting] or ["- none"])
    return "\n".join(lines)


def _build_fold_reuse_diagnostics(fold_rows: list[dict[str, Any]]) -> dict[str, float]:
    usage: dict[str, list[tuple[int, int]]] = {"train": [], "validation": [], "test": []}
    for row in fold_rows:
        idx = row.get("indices", {})
        for name in usage:
            bounds = idx.get(name)
            if isinstance(bounds, list) and len(bounds) == 2:
                usage[name].append((int(bounds[0]), int(bounds[1])))

    diagnostics: dict[str, float] = {}
    for name, spans in usage.items():
        if not spans:
            diagnostics[f"{name}_avg_reuse"] = 0.0
            diagnostics[f"{name}_max_reuse"] = 0.0
            continue
        max_end = max(end for _, end in spans)
        counts = np.zeros(max_end, dtype=int)
        for start, end in spans:
            counts[max(0, start): max(0, end)] += 1
        used = counts[counts > 0]
        diagnostics[f"{name}_avg_reuse"] = float(np.mean(used)) if used.size else 0.0
        diagnostics[f"{name}_max_reuse"] = float(np.max(used)) if used.size else 0.0
    return diagnostics

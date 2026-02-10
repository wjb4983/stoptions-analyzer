from __future__ import annotations

import csv
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


@dataclass
class WalkForwardResult:
    folds: list[dict[str, Any]]
    aggregate_metrics: dict[str, float]
    stability: dict[str, Any]


def build_walk_forward_folds(
    *,
    total_bars: int,
    train_bars: int,
    validation_bars: int,
    test_bars: int,
    step_bars: int | None = None,
) -> list[WalkForwardFold]:
    if min(total_bars, train_bars, validation_bars, test_bars) <= 0:
        raise ValueError("window sizes and total_bars must be positive")

    step = step_bars if step_bars is not None else test_bars
    if step <= 0:
        raise ValueError("step_bars must be positive")

    folds: list[WalkForwardFold] = []
    cursor = 0
    fold_id = 0
    while cursor + train_bars + validation_bars + test_bars <= total_bars:
        train_start = cursor
        train_end = train_start + train_bars
        validation_start = train_end
        validation_end = validation_start + validation_bars
        test_start = validation_end
        test_end = test_start + test_bars
        folds.append(
            WalkForwardFold(
                fold_id=fold_id,
                train_start=train_start,
                train_end=train_end,
                validation_start=validation_start,
                validation_end=validation_end,
                test_start=test_start,
                test_end=test_end,
            )
        )
        fold_id += 1
        cursor += step
    return folds


def run_walk_forward_optimization(
    *,
    folds: list[WalkForwardFold],
    parameter_candidates: list[dict[str, Any]],
    evaluate_segment: Callable[[dict[str, Any], int, int], dict[str, Any]],
    score_metric: str = "sharpe",
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
            score = float(validation_eval.get("metrics", {}).get(score_metric, float("-inf")))
            row = {
                "params": dict(candidate),
                "train_metrics": dict(train_eval.get("metrics", {})),
                "validation_metrics": dict(validation_eval.get("metrics", {})),
                "validation_score": score,
            }
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
                "selected_params": best_candidate,
                "validation_score": best_score,
                "oos_metrics": dict(oos_eval.get("metrics", {})),
                "oos_equity": list(oos_eval.get("equity", [])),
                "diagnostics": diagnostics,
            }
        )

    aggregate_metrics = _aggregate_numeric_metrics([row["oos_metrics"] for row in fold_rows])
    stability = _build_stability_summary(selected_keys, fold_rows)
    return WalkForwardResult(folds=fold_rows, aggregate_metrics=aggregate_metrics, stability=stability)


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
        (fold_dir / "diagnostics.json").write_text(json.dumps(fold_row["diagnostics"], indent=2))

        with (fold_dir / "oos_equity.csv").open("w", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=["timestamp", "equity"])
            writer.writeheader()
            for row in fold_row["oos_equity"]:
                writer.writerow({"timestamp": row.get("timestamp"), "equity": row.get("equity")})

    (run_dir / "fold_summary.json").write_text(json.dumps(result.folds, indent=2))
    (run_dir / "aggregate_metrics.json").write_text(json.dumps(result.aggregate_metrics, indent=2))
    (run_dir / "stability.json").write_text(json.dumps(result.stability, indent=2))


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
    }

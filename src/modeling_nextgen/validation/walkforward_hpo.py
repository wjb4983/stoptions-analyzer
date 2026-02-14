from __future__ import annotations

import csv
import json
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Mapping, Sequence

import numpy as np


@dataclass(frozen=True)
class WalkForwardWindow:
    fold_id: int
    train_start: int
    train_end: int
    test_start: int
    test_end: int


@dataclass
class WalkForwardHPOSummary:
    folds: list[dict[str, Any]]
    aggregate_metrics: dict[str, float]
    parameter_stability: dict[str, Any]
    model_card: dict[str, Any]


def build_walkforward_windows(
    *,
    total_samples: int,
    train_size: int,
    test_size: int,
    step_size: int | None = None,
) -> list[WalkForwardWindow]:
    if min(total_samples, train_size, test_size) <= 0:
        raise ValueError("total_samples, train_size, and test_size must be positive")

    step = test_size if step_size is None else step_size
    if step <= 0:
        raise ValueError("step_size must be positive")

    windows: list[WalkForwardWindow] = []
    cursor = 0
    fold_id = 0
    while cursor + train_size + test_size <= total_samples:
        train_start = cursor
        train_end = cursor + train_size
        test_start = train_end
        test_end = test_start + test_size
        windows.append(
            WalkForwardWindow(
                fold_id=fold_id,
                train_start=train_start,
                train_end=train_end,
                test_start=test_start,
                test_end=test_end,
            )
        )
        cursor += step
        fold_id += 1

    if not windows:
        raise ValueError("No walk-forward windows could be generated for the requested geometry")
    return windows


def run_walkforward_hpo(
    *,
    windows: Sequence[WalkForwardWindow],
    param_candidates: Sequence[Mapping[str, Any]],
    inner_objective: Callable[[Mapping[str, Any], int, int], float],
    outer_evaluator: Callable[[Mapping[str, Any], int, int], Mapping[str, Any]],
    primary_metric: str,
) -> WalkForwardHPOSummary:
    if not windows:
        raise ValueError("At least one walk-forward window is required")
    if not param_candidates:
        raise ValueError("At least one parameter candidate is required")

    folds: list[dict[str, Any]] = []

    for window in windows:
        candidate_scores: list[dict[str, Any]] = []
        best_score = float("-inf")
        best_candidate: dict[str, Any] | None = None

        for candidate in param_candidates:
            score = float(inner_objective(candidate, window.train_start, window.train_end))
            candidate_payload = {
                "params": dict(candidate),
                "inner_score": score,
            }
            candidate_scores.append(candidate_payload)

            tie_breaker = json.dumps(candidate_payload["params"], sort_keys=True)
            current_best = json.dumps(best_candidate, sort_keys=True) if best_candidate is not None else ""
            if score > best_score or (score == best_score and tie_breaker < current_best):
                best_score = score
                best_candidate = dict(candidate)

        if best_candidate is None:
            raise RuntimeError(f"No candidate selected for fold {window.fold_id}")

        outer_eval = dict(outer_evaluator(best_candidate, window.test_start, window.test_end))
        folds.append(
            {
                "fold_id": window.fold_id,
                "indices": {
                    "train": [window.train_start, window.train_end],
                    "test": [window.test_start, window.test_end],
                },
                "selected_params": best_candidate,
                "inner_best_score": best_score,
                "inner_candidates": candidate_scores,
                "outer_metrics": dict(outer_eval.get("metrics", {})),
                "outer_output": outer_eval,
            }
        )

    aggregate_metrics = _aggregate_numeric_metrics([row["outer_metrics"] for row in folds])
    parameter_stability = _build_parameter_stability(folds)
    model_card = _build_model_card(
        folds=folds,
        aggregate_metrics=aggregate_metrics,
        parameter_stability=parameter_stability,
        primary_metric=primary_metric,
    )

    return WalkForwardHPOSummary(
        folds=folds,
        aggregate_metrics=aggregate_metrics,
        parameter_stability=parameter_stability,
        model_card=model_card,
    )


def export_walkforward_hpo_reports(
    *,
    summary: WalkForwardHPOSummary,
    reports_dir: str | Path = "reports",
    artifact_prefix: str = "walkforward_hpo",
) -> dict[str, Path]:
    base = Path(reports_dir)
    base.mkdir(parents=True, exist_ok=True)

    summary_path = base / f"{artifact_prefix}_summary.json"
    model_card_path = base / f"{artifact_prefix}_model_card.json"
    folds_csv_path = base / f"{artifact_prefix}_folds.csv"
    stability_path = base / f"{artifact_prefix}_stability.json"

    summary_payload = {
        "card_version": "1.0",
        "generated_at": _utc_now(),
        "artifact_type": "walkforward_hpo_summary",
        "folds": summary.folds,
        "aggregate_metrics": summary.aggregate_metrics,
        "parameter_stability": summary.parameter_stability,
    }

    summary_path.write_text(json.dumps(summary_payload, indent=2) + "\n")
    model_card_path.write_text(json.dumps(summary.model_card, indent=2) + "\n")
    stability_path.write_text(json.dumps(summary.parameter_stability, indent=2) + "\n")

    _write_folds_csv(summary.folds, folds_csv_path)

    return {
        "summary": summary_path,
        "model_card": model_card_path,
        "folds_csv": folds_csv_path,
        "stability": stability_path,
    }


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _aggregate_numeric_metrics(rows: Sequence[Mapping[str, Any]]) -> dict[str, float]:
    keys = sorted({k for row in rows for k, value in row.items() if isinstance(value, (int, float))})
    out: dict[str, float] = {}
    for key in keys:
        values = np.asarray([float(row[key]) for row in rows if key in row], dtype=float)
        if values.size == 0:
            continue
        out[f"{key}_mean"] = float(np.mean(values))
        out[f"{key}_std"] = float(np.std(values, ddof=0))
        out[f"{key}_min"] = float(np.min(values))
        out[f"{key}_max"] = float(np.max(values))
    return out


def _build_parameter_stability(folds: Sequence[Mapping[str, Any]]) -> dict[str, Any]:
    selected = [dict(row.get("selected_params", {})) for row in folds]
    selected_keys = [json.dumps(params, sort_keys=True) for params in selected]

    selection_counts: dict[str, int] = {}
    for key in selected_keys:
        selection_counts[key] = selection_counts.get(key, 0) + 1

    transitions = 0
    per_fold: list[dict[str, Any]] = []
    for idx, params in enumerate(selected):
        changed_from_prev = idx > 0 and params != selected[idx - 1]
        if changed_from_prev:
            transitions += 1
        per_fold.append(
            {
                "fold_id": int(folds[idx].get("fold_id", idx)),
                "selected_params": params,
                "changed_from_previous": bool(changed_from_prev),
            }
        )

    numeric_drift = _numeric_param_drift(selected)
    total_transitions = max(1, len(selected) - 1)

    return {
        "selection_counts": selection_counts,
        "unique_param_sets": len(selection_counts),
        "transition_rate": float(transitions / total_transitions),
        "numeric_parameter_drift": numeric_drift,
        "per_fold": per_fold,
    }


def _numeric_param_drift(selected: Sequence[Mapping[str, Any]]) -> dict[str, float]:
    if len(selected) < 2:
        return {}

    all_keys = sorted({key for params in selected for key, value in params.items() if isinstance(value, (int, float))})
    drift: dict[str, float] = {}
    for key in all_keys:
        series: list[float] = []
        for params in selected:
            value = params.get(key)
            if isinstance(value, (int, float)):
                series.append(float(value))
        if len(series) < 2:
            continue
        deltas = np.abs(np.diff(np.asarray(series, dtype=float)))
        drift[f"{key}_mean_abs_delta"] = float(np.mean(deltas))
        drift[f"{key}_max_abs_delta"] = float(np.max(deltas))
    return drift


def _build_model_card(
    *,
    folds: Sequence[Mapping[str, Any]],
    aggregate_metrics: Mapping[str, float],
    parameter_stability: Mapping[str, Any],
    primary_metric: str,
) -> dict[str, Any]:
    primary_mean_key = f"{primary_metric}_mean"
    primary_mean = float(aggregate_metrics.get(primary_mean_key, 0.0))
    transition_rate = float(parameter_stability.get("transition_rate", 0.0))
    deployment_ready = bool(np.isfinite(primary_mean) and primary_mean > 0.0 and transition_rate <= 0.8)

    return {
        "card_version": "1.0",
        "generated_at": _utc_now(),
        "model_name": "walkforward-hpo-nextgen-validator",
        "intended_use": "Nested walk-forward tuning and out-of-sample validation for next-gen models.",
        "validation_summary": {
            "fold_count": len(folds),
            "primary_metric": primary_metric,
            "primary_metric_mean": primary_mean,
        },
        "quality_gates": {
            "outer_loop_coverage": {
                "name": "outer_loop_coverage",
                "pass": len(folds) >= 2,
                "details": {"fold_count": len(folds)},
            },
            "parameter_stability": {
                "name": "parameter_stability",
                "pass": transition_rate <= 0.8,
                "details": {
                    "transition_rate": transition_rate,
                    "unique_param_sets": int(parameter_stability.get("unique_param_sets", 0)),
                },
            },
            "primary_metric": {
                "name": "primary_metric",
                "pass": bool(np.isfinite(primary_mean)),
                "details": {"mean": primary_mean, "metric": primary_metric},
            },
        },
        "deployment_readiness": bool(deployment_ready),
        "limitations": [
            "Walk-forward geometry sensitivity can materially alter fold-level outcomes.",
            "Parameter stability is historical and does not guarantee future stationarity.",
        ],
    }


def _write_folds_csv(folds: Sequence[Mapping[str, Any]], path: Path) -> None:
    fieldnames = [
        "fold_id",
        "train_start",
        "train_end",
        "test_start",
        "test_end",
        "selected_params",
        "inner_best_score",
        "outer_metrics",
    ]
    with path.open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for fold in folds:
            indices = fold.get("indices", {})
            train = indices.get("train", [None, None])
            test = indices.get("test", [None, None])
            writer.writerow(
                {
                    "fold_id": fold.get("fold_id"),
                    "train_start": train[0],
                    "train_end": train[1],
                    "test_start": test[0],
                    "test_end": test[1],
                    "selected_params": json.dumps(fold.get("selected_params", {}), sort_keys=True),
                    "inner_best_score": fold.get("inner_best_score"),
                    "outer_metrics": json.dumps(fold.get("outer_metrics", {}), sort_keys=True),
                }
            )


__all__ = [
    "WalkForwardWindow",
    "WalkForwardHPOSummary",
    "build_walkforward_windows",
    "run_walkforward_hpo",
    "export_walkforward_hpo_reports",
]

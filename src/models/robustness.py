from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import numpy as np

from modeling_nextgen.validation.stress_scenarios import build_stress_template_scenarios

from .base import ModelInterface


@dataclass(frozen=True)
class RobustnessThresholds:
    min_model_score: float = 0.7
    min_feature_group_score: float = 0.65
    max_brittle_features: int = 1


def _as_array(features: dict[str, np.ndarray]) -> dict[str, np.ndarray]:
    return {name: np.asarray(values, dtype=float).copy() for name, values in features.items()}


def _safe_predict_proba(model: ModelInterface, features: dict[str, np.ndarray]) -> np.ndarray:
    preds = np.asarray(model.predict_proba(features), dtype=float)
    return np.nan_to_num(preds, nan=0.5, posinf=1.0, neginf=0.0)


def _shift_with_delay(values: np.ndarray, delay: int) -> np.ndarray:
    if delay <= 0 or values.size == 0:
        return values.copy()
    shifted = np.empty_like(values)
    lead = values[0]
    shifted[:delay] = lead
    shifted[delay:] = values[:-delay]
    return shifted


def _perturb_features(
    *,
    features: dict[str, np.ndarray],
    rng: np.random.Generator,
    noise_scale: float,
    delay_bars: int,
    missing_fraction: float,
    distribution_shift_scale: float,
) -> dict[str, dict[str, np.ndarray]]:
    base = _as_array(features)

    noisy: dict[str, np.ndarray] = {}
    delayed: dict[str, np.ndarray] = {}
    missing: dict[str, np.ndarray] = {}
    shifted: dict[str, np.ndarray] = {}

    for name, values in base.items():
        sigma = float(np.nanstd(values))
        sigma = sigma if sigma > 0.0 else 1.0
        noisy[name] = values + rng.normal(0.0, sigma * noise_scale, size=values.shape)

        delayed[name] = _shift_with_delay(values, delay_bars)

        missing_arr = values.copy()
        if missing_fraction > 0.0:
            mask = rng.random(size=values.shape) < missing_fraction
            fill = float(np.nanmean(values)) if np.isfinite(np.nanmean(values)) else 0.0
            missing_arr[mask] = fill
        missing[name] = missing_arr

        center = float(np.nanmean(values)) if np.isfinite(np.nanmean(values)) else 0.0
        shifted[name] = (values - center) * (1.0 + distribution_shift_scale) + center + sigma * distribution_shift_scale

    perturbations = {
        "noise": noisy,
        "delayed_data": delayed,
        "missing_fields": missing,
        "shifted_distribution": shifted,
    }
    perturbations.update(build_stress_template_scenarios(features=base, rng=rng))
    return perturbations


def _prediction_degradation(baseline: np.ndarray, perturbed: np.ndarray) -> float:
    baseline = np.asarray(baseline, dtype=float)
    perturbed = np.asarray(perturbed, dtype=float)
    delta = float(np.mean(np.abs(perturbed - baseline)))
    norm = float(np.mean(np.abs(baseline - 0.5)))
    denom = max(0.05, norm)
    return min(1.0, delta / denom)


def build_robustness_scorecards(
    *,
    models: dict[str, ModelInterface],
    feature_payloads: dict[str, dict[str, np.ndarray]],
    feature_groups: dict[str, tuple[str, ...]] | None = None,
    thresholds: RobustnessThresholds | None = None,
    noise_scale: float = 0.25,
    delay_bars: int = 1,
    missing_fraction: float = 0.2,
    distribution_shift_scale: float = 0.3,
    feature_brittleness_threshold: float = 0.45,
    seed: int = 13,
) -> dict[str, Any]:
    """Build robustness scorecards by model and feature group.

    Degradation is measured from prediction changes under perturbations.
    Higher scores indicate stronger resilience.
    """

    thresholds = thresholds or RobustnessThresholds()
    rng = np.random.default_rng(seed)
    group_defs = feature_groups or {}

    model_rows: list[dict[str, Any]] = []
    for model_id, model in models.items():
        features = feature_payloads.get(model_id)
        if not isinstance(features, dict) or not features:
            continue

        baseline_features = _as_array(features)
        baseline = _safe_predict_proba(model, baseline_features)
        perturbations = _perturb_features(
            features=baseline_features,
            rng=rng,
            noise_scale=noise_scale,
            delay_bars=delay_bars,
            missing_fraction=missing_fraction,
            distribution_shift_scale=distribution_shift_scale,
        )

        scenario_degradation: dict[str, float] = {}
        for scenario, scenario_features in perturbations.items():
            scenario_preds = _safe_predict_proba(model, scenario_features)
            scenario_degradation[scenario] = _prediction_degradation(baseline, scenario_preds)

        model_score = float(max(0.0, 1.0 - np.mean(list(scenario_degradation.values()))))

        feature_rows: list[dict[str, Any]] = []
        brittle_features: list[str] = []
        for feature_name in baseline_features:
            ablation = _as_array(baseline_features)
            fill = float(np.nanmean(ablation[feature_name])) if np.isfinite(np.nanmean(ablation[feature_name])) else 0.0
            ablation[feature_name] = np.full_like(ablation[feature_name], fill)
            ablated_pred = _safe_predict_proba(model, ablation)
            degradation = _prediction_degradation(baseline, ablated_pred)
            score = float(max(0.0, 1.0 - degradation))
            is_brittle = degradation >= feature_brittleness_threshold
            if is_brittle:
                brittle_features.append(feature_name)
            feature_rows.append(
                {
                    "feature": feature_name,
                    "degradation": float(degradation),
                    "score": score,
                    "brittle": bool(is_brittle),
                    "recommended_action": "redesign_or_remove" if is_brittle else "keep",
                }
            )

        group_scores: dict[str, dict[str, Any]] = {}
        for group_name, group_features in group_defs.items():
            selected = [row for row in feature_rows if row["feature"] in set(group_features)]
            if not selected:
                continue
            score = float(np.mean([float(row["score"]) for row in selected]))
            group_scores[group_name] = {
                "score": score,
                "pass": score >= thresholds.min_feature_group_score,
                "features": [row["feature"] for row in selected],
            }

        meets_threshold = model_score >= thresholds.min_model_score and len(brittle_features) <= thresholds.max_brittle_features
        if group_scores:
            meets_threshold = meets_threshold and all(bool(row["pass"]) for row in group_scores.values())

        model_rows.append(
            {
                "model_id": model_id,
                "robustness_score": model_score,
                "meets_minimum_threshold": bool(meets_threshold),
                "scenario_degradation": scenario_degradation,
                "feature_groups": group_scores,
                "brittle_features": brittle_features,
                "feature_breakdown": feature_rows,
            }
        )

    production_ready = bool(model_rows) and all(bool(row["meets_minimum_threshold"]) for row in model_rows)
    return {
        "thresholds": {
            "min_model_score": thresholds.min_model_score,
            "min_feature_group_score": thresholds.min_feature_group_score,
            "max_brittle_features": thresholds.max_brittle_features,
        },
        "production_ready": production_ready,
        "models": model_rows,
        "recommended_feature_actions": [
            {"model_id": row["model_id"], "feature": feat, "action": "redesign_or_remove"}
            for row in model_rows
            for feat in row["brittle_features"]
        ],
    }

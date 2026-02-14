from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable

import numpy as np


@dataclass(frozen=True)
class AdversarialValidationConfig:
    flip_fraction: float = 0.2
    jitter_probability: float = 0.25
    jitter_max_lag: int = 2
    missing_block_fraction: float = 0.15
    missing_block_min_length: int = 2
    missing_block_max_length: int = 8
    max_fragility_score: float = 0.35
    max_worst_case_degradation: float = 0.55


def _as_array(features: dict[str, np.ndarray]) -> dict[str, np.ndarray]:
    return {name: np.asarray(values, dtype=float).copy() for name, values in features.items()}


def _safe_predict(
    model: Any,
    features: dict[str, np.ndarray],
    predict_fn: Callable[[Any, dict[str, np.ndarray]], np.ndarray] | None,
) -> np.ndarray:
    if predict_fn is not None:
        preds = np.asarray(predict_fn(model, features), dtype=float)
    elif hasattr(model, "predict_proba"):
        preds = np.asarray(model.predict_proba(features), dtype=float)
    elif hasattr(model, "predict"):
        preds = np.asarray(model.predict(features), dtype=float)
    else:
        raise TypeError("Model must expose predict_proba/predict or a custom predict_fn must be supplied.")
    return np.nan_to_num(preds, nan=0.5, posinf=1.0, neginf=0.0)


def _prediction_degradation(baseline: np.ndarray, perturbed: np.ndarray) -> float:
    baseline = np.asarray(baseline, dtype=float)
    perturbed = np.asarray(perturbed, dtype=float)
    delta = float(np.mean(np.abs(perturbed - baseline)))
    norm = float(np.mean(np.abs(baseline - 0.5)))
    return float(min(1.0, delta / max(0.05, norm)))


def _feature_sign_flip_corruption(
    *,
    features: dict[str, np.ndarray],
    rng: np.random.Generator,
    flip_fraction: float,
) -> dict[str, np.ndarray]:
    corrupted = _as_array(features)
    p = float(np.clip(flip_fraction, 0.0, 1.0))
    for name, values in corrupted.items():
        if values.size == 0:
            continue
        mask = rng.random(size=values.shape) < p
        values[mask] *= -1.0
        corrupted[name] = values
    return corrupted


def _temporal_jitter_corruption(
    *,
    features: dict[str, np.ndarray],
    rng: np.random.Generator,
    jitter_probability: float,
    jitter_max_lag: int,
) -> dict[str, np.ndarray]:
    corrupted = _as_array(features)
    p = float(np.clip(jitter_probability, 0.0, 1.0))
    lag = max(0, int(jitter_max_lag))
    if lag == 0:
        return corrupted

    for name, values in corrupted.items():
        n = values.size
        if n == 0:
            continue
        base_idx = np.arange(n)
        jitter_mask = rng.random(size=n) < p
        shift = rng.integers(-lag, lag + 1, size=n)
        idx = np.where(jitter_mask, np.clip(base_idx + shift, 0, n - 1), base_idx)
        corrupted[name] = values[idx]
    return corrupted


def _sparse_missing_block_corruption(
    *,
    features: dict[str, np.ndarray],
    rng: np.random.Generator,
    missing_block_fraction: float,
    min_block_length: int,
    max_block_length: int,
) -> dict[str, np.ndarray]:
    corrupted = _as_array(features)
    target_fraction = float(np.clip(missing_block_fraction, 0.0, 1.0))
    min_len = max(1, int(min_block_length))
    max_len = max(min_len, int(max_block_length))

    for name, values in corrupted.items():
        n = values.size
        if n == 0 or target_fraction <= 0.0:
            continue
        missing_mask = np.zeros(n, dtype=bool)
        target_points = int(np.ceil(target_fraction * n))
        missing_points = 0
        while missing_points < target_points:
            length = int(rng.integers(min_len, max_len + 1))
            start = int(rng.integers(0, max(1, n - length + 1)))
            end = min(n, start + length)
            missing_mask[start:end] = True
            missing_points = int(missing_mask.sum())
            if missing_mask.all():
                break

        fill = float(np.nanmedian(values)) if np.isfinite(np.nanmedian(values)) else 0.0
        values[missing_mask] = fill
        corrupted[name] = values
    return corrupted


def build_adversarial_fragility_scorecards(
    *,
    models: dict[str, Any],
    feature_payloads: dict[str, dict[str, np.ndarray]],
    config: AdversarialValidationConfig | None = None,
    predict_fn: Callable[[Any, dict[str, np.ndarray]], np.ndarray] | None = None,
    seed: int = 23,
) -> dict[str, Any]:
    """Evaluate model brittleness under adversarial feature corruptions.

    Returns per-model fragility scorecards and auto-fail gate outcomes.
    """

    cfg = config or AdversarialValidationConfig()
    rng = np.random.default_rng(seed)

    scorecards: list[dict[str, Any]] = []
    for model_id, model in models.items():
        features = feature_payloads.get(model_id)
        if not isinstance(features, dict) or not features:
            continue

        baseline_features = _as_array(features)
        baseline_preds = _safe_predict(model, baseline_features, predict_fn)

        scenarios = {
            "feature_sign_flips": _feature_sign_flip_corruption(
                features=baseline_features,
                rng=rng,
                flip_fraction=cfg.flip_fraction,
            ),
            "temporal_jitter": _temporal_jitter_corruption(
                features=baseline_features,
                rng=rng,
                jitter_probability=cfg.jitter_probability,
                jitter_max_lag=cfg.jitter_max_lag,
            ),
            "sparse_missing_block_corruption": _sparse_missing_block_corruption(
                features=baseline_features,
                rng=rng,
                missing_block_fraction=cfg.missing_block_fraction,
                min_block_length=cfg.missing_block_min_length,
                max_block_length=cfg.missing_block_max_length,
            ),
        }

        scenario_degradation: dict[str, float] = {}
        for scenario_name, scenario_features in scenarios.items():
            perturbed_preds = _safe_predict(model, scenario_features, predict_fn)
            scenario_degradation[scenario_name] = _prediction_degradation(baseline_preds, perturbed_preds)

        fragility_score = float(np.mean(list(scenario_degradation.values()))) if scenario_degradation else 1.0
        worst_case = float(max(scenario_degradation.values())) if scenario_degradation else 1.0
        auto_fail_reasons: list[str] = []
        if fragility_score > cfg.max_fragility_score:
            auto_fail_reasons.append("fragility_score_exceeds_limit")
        if worst_case > cfg.max_worst_case_degradation:
            auto_fail_reasons.append("worst_case_degradation_exceeds_limit")

        scorecards.append(
            {
                "model_id": model_id,
                "fragility_score": fragility_score,
                "resilience_score": float(max(0.0, 1.0 - fragility_score)),
                "worst_case_degradation": worst_case,
                "scenario_degradation": scenario_degradation,
                "auto_fail_gate": bool(auto_fail_reasons),
                "auto_fail_reasons": auto_fail_reasons,
            }
        )

    pass_rate = float(np.mean([0.0 if row["auto_fail_gate"] else 1.0 for row in scorecards])) if scorecards else 0.0
    return {
        "config": {
            "flip_fraction": cfg.flip_fraction,
            "jitter_probability": cfg.jitter_probability,
            "jitter_max_lag": cfg.jitter_max_lag,
            "missing_block_fraction": cfg.missing_block_fraction,
            "missing_block_min_length": cfg.missing_block_min_length,
            "missing_block_max_length": cfg.missing_block_max_length,
            "max_fragility_score": cfg.max_fragility_score,
            "max_worst_case_degradation": cfg.max_worst_case_degradation,
        },
        "models": scorecards,
        "pass_rate": pass_rate,
        "all_models_pass": bool(scorecards) and all(not row["auto_fail_gate"] for row in scorecards),
    }

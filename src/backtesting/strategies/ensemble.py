from __future__ import annotations

from dataclasses import dataclass

import numpy as np


def weighted_voting(signals: np.ndarray, weights: np.ndarray) -> np.ndarray:
    normalized = np.asarray(weights, dtype=float)
    gross = float(np.sum(np.abs(normalized)))
    if gross <= 0.0:
        raise ValueError("weights must have non-zero gross")
    normalized = normalized / gross
    return np.tensordot(signals, normalized, axes=([-1], [0]))


def risk_budgeted_blend(signals: np.ndarray, risk_budgets: np.ndarray, vol_estimates: np.ndarray) -> np.ndarray:
    budgets = np.asarray(risk_budgets, dtype=float)
    vol = np.asarray(vol_estimates, dtype=float)
    if np.any(vol <= 0.0):
        raise ValueError("vol_estimates must be > 0")
    scaled = budgets / vol
    gross = float(np.sum(np.abs(scaled)))
    if gross <= 0.0:
        raise ValueError("risk_budgets must have non-zero gross")
    scaled = scaled / gross
    return np.tensordot(signals, scaled, axes=([-1], [0]))


def meta_model_weighting(signals: np.ndarray, feature_weights: np.ndarray, bias: float = 0.0) -> np.ndarray:
    stacked = np.asarray(signals, dtype=float)
    linear = np.tensordot(stacked, np.asarray(feature_weights, dtype=float), axes=([-1], [0])) + float(bias)
    return np.tanh(linear)


def dynamic_model_weights(
    quality_history: np.ndarray,
    *,
    lookback: int = 20,
    min_weight: float = 0.0,
    eps: float = 1e-12,
) -> np.ndarray:
    """Compute dynamic ensemble weights from rolling out-of-sample quality metrics.

    quality_history is expected to be shape (n_periods, n_models).
    """

    quality = np.asarray(quality_history, dtype=float)
    if quality.ndim != 2:
        raise ValueError("quality_history must be a 2D array")
    if quality.shape[0] == 0:
        raise ValueError("quality_history must contain at least one row")

    window = max(1, min(int(lookback), quality.shape[0]))
    recent = quality[-window:, :]
    score = np.nanmean(recent, axis=0)
    score = np.where(np.isfinite(score), score, 0.0)
    score = np.maximum(score, float(min_weight))
    gross = float(np.sum(np.abs(score)))
    if gross <= eps:
        return np.full(score.shape, 1.0 / max(1, score.size), dtype=float)
    return score / gross


def rolling_dynamic_ensemble(
    signals: np.ndarray,
    quality_history: np.ndarray,
    *,
    lookback: int = 20,
    min_weight: float = 0.0,
) -> np.ndarray:
    """Blend model signals with period-by-period dynamic weights."""

    sig = np.asarray(signals, dtype=float)
    quality = np.asarray(quality_history, dtype=float)
    if sig.ndim != 2 or quality.ndim != 2:
        raise ValueError("signals and quality_history must be 2D arrays")
    if sig.shape != quality.shape:
        raise ValueError("signals and quality_history must have matching shape")

    blended = np.zeros(sig.shape[0], dtype=float)
    for idx in range(sig.shape[0]):
        weights = dynamic_model_weights(
            quality[: idx + 1, :],
            lookback=lookback,
            min_weight=min_weight,
        )
        blended[idx] = float(np.dot(sig[idx, :], weights))
    return blended


@dataclass(frozen=True)
class RegimeMetaPolicyConfig:
    strategy_names: tuple[str, ...]
    regime_weight_map: dict[str, tuple[float, ...]]
    default_weights: tuple[float, ...] | None = None
    turnover_limit: float = 0.35
    max_weight: float = 0.75
    min_weight: float = 0.0


def _normalize_with_constraints(
    weights: np.ndarray,
    *,
    min_weight: float,
    max_weight: float,
    eps: float = 1e-12,
) -> np.ndarray:
    clipped = np.clip(np.asarray(weights, dtype=float), float(min_weight), float(max_weight))
    total = float(np.sum(clipped))
    if total <= eps:
        return np.full(clipped.shape, 1.0 / max(1, clipped.size), dtype=float)
    return clipped / total


def build_regime_weight_schedule(
    *,
    regime_labels: np.ndarray,
    config: RegimeMetaPolicyConfig,
) -> tuple[np.ndarray, dict[str, np.ndarray]]:
    labels = np.asarray(regime_labels, dtype=object)
    n_periods = labels.size
    n_strategies = len(config.strategy_names)
    if n_strategies <= 0:
        raise ValueError("config.strategy_names must be non-empty")

    if config.default_weights is None:
        base = np.full(n_strategies, 1.0 / n_strategies, dtype=float)
    else:
        if len(config.default_weights) != n_strategies:
            raise ValueError("default_weights must match strategy_names length")
        base = np.asarray(config.default_weights, dtype=float)

    max_turnover = max(0.0, float(config.turnover_limit))
    prev = _normalize_with_constraints(base, min_weight=config.min_weight, max_weight=config.max_weight)
    schedule = np.zeros((n_periods, n_strategies), dtype=float)
    turnover = np.zeros(n_periods, dtype=float)

    for idx in range(n_periods):
        label = str(labels[idx])
        proposed = config.regime_weight_map.get(label)
        raw = np.asarray(proposed if proposed is not None else prev, dtype=float)
        if raw.size != n_strategies:
            raise ValueError(f"regime vector length mismatch for {label}")
        target = _normalize_with_constraints(raw, min_weight=config.min_weight, max_weight=config.max_weight)

        diff = target - prev
        step_turnover = float(np.sum(np.abs(diff)))
        if step_turnover > max_turnover and step_turnover > 0.0:
            target = prev + diff * (max_turnover / step_turnover)
            target = _normalize_with_constraints(target, min_weight=config.min_weight, max_weight=config.max_weight)
            step_turnover = float(np.sum(np.abs(target - prev)))

        turnover[idx] = step_turnover
        schedule[idx] = target
        prev = target

    diagnostics = {
        "meta_turnover": turnover,
        "meta_concentration": np.max(schedule, axis=1) if n_periods else np.array([], dtype=float),
    }
    return schedule, diagnostics


def apply_weight_schedule(strategy_sleeves: np.ndarray, weight_schedule: np.ndarray) -> np.ndarray:
    sleeves = np.asarray(strategy_sleeves, dtype=float)
    weights = np.asarray(weight_schedule, dtype=float)
    if sleeves.ndim != 2 or weights.ndim != 2:
        raise ValueError("strategy_sleeves and weight_schedule must be 2D arrays")
    if sleeves.shape != weights.shape:
        raise ValueError("strategy_sleeves and weight_schedule must have matching shape")
    return np.sum(sleeves * weights, axis=1)

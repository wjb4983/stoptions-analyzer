from __future__ import annotations

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

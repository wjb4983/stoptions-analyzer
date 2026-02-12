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

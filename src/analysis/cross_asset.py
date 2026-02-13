from __future__ import annotations

from dataclasses import dataclass
from typing import Mapping, Sequence

import numpy as np


@dataclass(frozen=True)
class ModelEvaluation:
    name: str
    mse: float
    r2: float


@dataclass(frozen=True)
class CrossAssetModelComparison:
    equity_only: ModelEvaluation
    options_only: ModelEvaluation
    cross_asset_conditioned: ModelEvaluation
    outperformance_vs_best_isolated_mse: float
    outperformance_vs_best_isolated_r2: float


@dataclass(frozen=True)
class LeadLagDiagnostics:
    lag_grid: np.ndarray
    correlations: dict[str, np.ndarray]
    strongest_driver: str | None
    strongest_lag: int | None
    strongest_correlation: float
    transmission_betas: dict[str, float]


def build_release_aware_macro_features(
    *,
    target_timestamps: Sequence[float],
    events: Sequence[Mapping[str, float]],
    value_key: str = "value",
    release_time_key: str = "release_ts",
    event_time_key: str = "event_ts",
    max_lag_events: int = 2,
) -> dict[str, np.ndarray]:
    """Join macro events to target timestamps using release-aware (no-lookahead) logic.

    Each row in ``events`` must contain ``event_ts`` (economic period), ``release_ts``
    (publication time), and a numeric value field (default ``value``). For each target
    timestamp we expose the latest released event and lagged values.
    """

    ts = np.asarray(target_timestamps, dtype=float)
    if ts.ndim != 1:
        raise ValueError("target_timestamps must be 1D")

    normalized: list[tuple[float, float, float]] = []
    for ev in events:
        if release_time_key not in ev or value_key not in ev or event_time_key not in ev:
            continue
        release_ts = float(ev[release_time_key])
        event_ts = float(ev[event_time_key])
        value = float(ev[value_key])
        if np.isfinite(release_ts) and np.isfinite(event_ts) and np.isfinite(value):
            normalized.append((release_ts, event_ts, value))

    normalized.sort(key=lambda row: row[0])
    release_ts = np.asarray([r for r, _, _ in normalized], dtype=float)
    event_ts = np.asarray([e for _, e, _ in normalized], dtype=float)
    values = np.asarray([v for _, _, v in normalized], dtype=float)

    n = ts.shape[0]
    latest_value = np.full(n, np.nan, dtype=float)
    latest_event_age = np.full(n, np.nan, dtype=float)
    latest_release_lag = np.full(n, np.nan, dtype=float)
    lag_matrix = np.full((n, max(0, int(max_lag_events))), np.nan, dtype=float)

    for i, t in enumerate(ts):
        idx = int(np.searchsorted(release_ts, t, side="right") - 1)
        if idx < 0:
            continue
        latest_value[i] = values[idx]
        latest_event_age[i] = t - event_ts[idx]
        latest_release_lag[i] = t - release_ts[idx]
        for lag in range(lag_matrix.shape[1]):
            src = idx - (lag + 1)
            if src >= 0:
                lag_matrix[i, lag] = values[src]

    out = {
        "macro_latest": latest_value,
        "macro_event_age": latest_event_age,
        "macro_release_lag": latest_release_lag,
    }
    for lag in range(lag_matrix.shape[1]):
        out[f"macro_lag_{lag+1}"] = lag_matrix[:, lag]
    return out


def join_cross_asset_indicators(
    *,
    target_timestamps: Sequence[float],
    indicators: Mapping[str, Mapping[str, Sequence[float]]],
) -> np.ndarray:
    """As-of join for cross-asset time series (rates, credit, FX, commodities, indices)."""

    ts = np.asarray(target_timestamps, dtype=float)
    if ts.ndim != 1:
        raise ValueError("target_timestamps must be 1D")

    columns: list[np.ndarray] = []
    for _, payload in indicators.items():
        ind_ts = np.asarray(payload.get("timestamps", []), dtype=float)
        ind_values = np.asarray(payload.get("values", []), dtype=float)
        if ind_ts.size != ind_values.size:
            raise ValueError("indicator timestamps and values must have same length")
        order = np.argsort(ind_ts)
        ind_ts = ind_ts[order]
        ind_values = ind_values[order]

        col = np.full(ts.shape[0], np.nan, dtype=float)
        for i, t in enumerate(ts):
            idx = int(np.searchsorted(ind_ts, t, side="right") - 1)
            if idx >= 0:
                col[i] = ind_values[idx]
        columns.append(col)

    if not columns:
        return np.empty((ts.shape[0], 0), dtype=float)
    return np.column_stack(columns)


def compare_conditioned_models(
    *,
    equity_features: np.ndarray,
    options_features: np.ndarray,
    cross_asset_features: np.ndarray,
    target: np.ndarray,
    train_fraction: float = 0.7,
    ridge: float = 1e-6,
) -> CrossAssetModelComparison:
    """Train/evaluate isolated equity/options models versus cross-asset-conditioned model."""

    y = np.asarray(target, dtype=float).reshape(-1)
    x_eq = _ensure_2d(equity_features)
    x_opt = _ensure_2d(options_features)
    x_xa = _ensure_2d(cross_asset_features)

    n = y.shape[0]
    if x_eq.shape[0] != n or x_opt.shape[0] != n or x_xa.shape[0] != n:
        raise ValueError("all feature matrices and target must align on first dimension")

    split = int(max(2, min(n - 1, round(n * float(train_fraction)))))
    eval_eq = _fit_eval_linear("equity_only", x_eq, y, split=split, ridge=ridge)
    eval_opt = _fit_eval_linear("options_only", x_opt, y, split=split, ridge=ridge)
    conditioned_x = np.column_stack([x_eq, x_opt, x_xa])
    eval_cond = _fit_eval_linear("cross_asset_conditioned", conditioned_x, y, split=split, ridge=ridge)

    best_isolated = min((eval_eq, eval_opt), key=lambda m: m.mse)
    return CrossAssetModelComparison(
        equity_only=eval_eq,
        options_only=eval_opt,
        cross_asset_conditioned=eval_cond,
        outperformance_vs_best_isolated_mse=best_isolated.mse - eval_cond.mse,
        outperformance_vs_best_isolated_r2=eval_cond.r2 - max(eval_eq.r2, eval_opt.r2),
    )


def compute_cross_market_transmission(
    *,
    target_series: np.ndarray,
    driver_series: Mapping[str, np.ndarray],
    max_lag: int = 5,
) -> LeadLagDiagnostics:
    """Estimate lead-lag correlations and simple transmission betas across markets."""

    y = np.asarray(target_series, dtype=float).reshape(-1)
    lags = np.arange(-int(max_lag), int(max_lag) + 1, dtype=int)

    correlations: dict[str, np.ndarray] = {}
    betas: dict[str, float] = {}
    strongest_driver: str | None = None
    strongest_lag: int | None = None
    strongest_corr = 0.0

    for name, values in driver_series.items():
        x = np.asarray(values, dtype=float).reshape(-1)
        if x.shape[0] != y.shape[0]:
            raise ValueError("driver series and target series must have same length")

        corr_vec = np.array([_lagged_corr(y, x, lag=int(lag)) for lag in lags], dtype=float)
        correlations[name] = corr_vec

        idx = int(np.nanargmax(np.abs(corr_vec))) if corr_vec.size else 0
        corr = float(corr_vec[idx]) if corr_vec.size else 0.0
        lag = int(lags[idx]) if lags.size else 0
        if abs(corr) > abs(strongest_corr):
            strongest_corr = corr
            strongest_driver = str(name)
            strongest_lag = lag

        betas[name] = _transmission_beta(y, x, lag)

    return LeadLagDiagnostics(
        lag_grid=lags,
        correlations=correlations,
        strongest_driver=strongest_driver,
        strongest_lag=strongest_lag,
        strongest_correlation=float(strongest_corr),
        transmission_betas=betas,
    )


def _ensure_2d(values: np.ndarray) -> np.ndarray:
    arr = np.asarray(values, dtype=float)
    if arr.ndim == 1:
        return arr.reshape(-1, 1)
    if arr.ndim != 2:
        raise ValueError("features must be 1D or 2D")
    return arr


def _fit_eval_linear(name: str, x: np.ndarray, y: np.ndarray, *, split: int, ridge: float) -> ModelEvaluation:
    x_train = x[:split]
    y_train = y[:split]
    x_test = x[split:]
    y_test = y[split:]
    coef = np.linalg.pinv(x_train.T @ x_train + np.eye(x_train.shape[1]) * ridge) @ x_train.T @ y_train
    preds = x_test @ coef
    residual = y_test - preds
    mse = float(np.mean(residual**2)) if residual.size else float("nan")
    var = float(np.var(y_test)) if y_test.size else 0.0
    r2 = float(1.0 - mse / var) if var > 0 else 0.0
    return ModelEvaluation(name=name, mse=mse, r2=r2)


def _lagged_corr(y: np.ndarray, x: np.ndarray, *, lag: int) -> float:
    if lag > 0:
        yv = y[lag:]
        xv = x[:-lag]
    elif lag < 0:
        yv = y[:lag]
        xv = x[-lag:]
    else:
        yv = y
        xv = x
    if yv.size < 3 or xv.size < 3:
        return 0.0
    y_std = float(np.std(yv))
    x_std = float(np.std(xv))
    if y_std == 0.0 or x_std == 0.0:
        return 0.0
    return float(np.corrcoef(yv, xv)[0, 1])


def _transmission_beta(y: np.ndarray, x: np.ndarray, lag: int) -> float:
    if lag > 0:
        yv = y[lag:]
        xv = x[:-lag]
    elif lag < 0:
        yv = y[:lag]
        xv = x[-lag:]
    else:
        yv = y
        xv = x
    if yv.size < 3:
        return 0.0
    denom = float(np.dot(xv, xv))
    if denom <= 0:
        return 0.0
    return float(np.dot(xv, yv) / denom)

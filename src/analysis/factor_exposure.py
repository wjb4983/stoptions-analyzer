from __future__ import annotations

from dataclasses import dataclass

import numpy as np


FACTOR_NAMES: tuple[str, ...] = (
    "market",
    "size",
    "value",
    "momentum",
    "quality",
    "volatility",
    "rates_credit_proxy",
)


@dataclass(frozen=True)
class FactorExposureResult:
    factor_names: tuple[str, ...]
    factor_returns: np.ndarray
    exposures_by_asset: np.ndarray


def build_factor_exposure_model(
    *,
    prices: np.ndarray,
    lookback: int = 20,
) -> FactorExposureResult:
    px = np.asarray(prices, dtype=float)
    if px.ndim != 2:
        raise ValueError("prices must be 2D")

    returns = np.zeros_like(px)
    returns[1:] = np.divide(px[1:], px[:-1], out=np.ones_like(px[1:]), where=px[:-1] != 0.0) - 1.0
    returns = np.nan_to_num(returns, nan=0.0, posinf=0.0, neginf=0.0)

    factors = _factor_proxy_returns(returns)
    exposures = _rolling_factor_exposures(asset_returns=returns, factor_returns=factors, lookback=lookback)
    return FactorExposureResult(
        factor_names=FACTOR_NAMES,
        factor_returns=factors,
        exposures_by_asset=exposures,
    )


def residualize_alpha_signals(
    *,
    raw_signals: np.ndarray,
    factor_exposures: np.ndarray,
) -> np.ndarray:
    signals = np.asarray(raw_signals, dtype=float)
    expo = np.asarray(factor_exposures, dtype=float)
    if signals.ndim != 2:
        raise ValueError("raw_signals must be 2D")
    if expo.ndim != 3 or expo.shape[:2] != signals.shape:
        raise ValueError("factor_exposures must have shape (periods, assets, factors)")

    out = np.zeros_like(signals)
    for t in range(signals.shape[0]):
        x = expo[t]
        y = signals[t]
        if x.size == 0 or y.size == 0:
            out[t] = y
            continue
        xtx = x.T @ x
        ridge = np.eye(xtx.shape[0]) * 1e-6
        beta = np.linalg.pinv(xtx + ridge) @ x.T @ y
        out[t] = y - x @ beta
    return out


def decompose_alpha(
    *,
    weights: np.ndarray,
    asset_returns: np.ndarray,
    factor_exposures: np.ndarray,
    factor_returns: np.ndarray,
) -> dict[str, np.ndarray]:
    w = np.asarray(weights, dtype=float)
    r = np.asarray(asset_returns, dtype=float)
    x = np.asarray(factor_exposures, dtype=float)
    f = np.asarray(factor_returns, dtype=float)
    if w.shape != r.shape:
        raise ValueError("weights and asset_returns must have same shape")
    if x.shape[:2] != w.shape:
        raise ValueError("factor_exposures must have shape (periods, assets, factors)")
    if f.shape[0] != w.shape[0] or f.shape[1] != x.shape[2]:
        raise ValueError("factor_returns must have shape (periods, factors)")

    gross_alpha = np.sum(w * r, axis=1)
    realized_factor_leg = np.einsum("ta,tak,tk->t", w, x, f)
    residual_alpha = gross_alpha - realized_factor_leg
    return {
        "gross_alpha": gross_alpha,
        "factor_beta_contribution": realized_factor_leg,
        "residual_alpha": residual_alpha,
    }


def _factor_proxy_returns(asset_returns: np.ndarray) -> np.ndarray:
    arr = np.asarray(asset_returns, dtype=float)
    periods, assets = arr.shape
    out = np.zeros((periods, len(FACTOR_NAMES)), dtype=float)
    if assets == 0:
        return out

    market = np.mean(arr, axis=1)
    cross_rank = np.argsort(np.argsort(arr, axis=1), axis=1)
    n = max(1, assets - 1)
    size_proxy = 1.0 - (cross_rank / n)
    value_proxy = -arr

    trailing_5 = np.zeros_like(arr)
    for t in range(periods):
        lo = max(0, t - 4)
        trailing_5[t] = np.mean(arr[lo : t + 1], axis=0)

    vol_proxy = np.zeros_like(arr)
    for t in range(periods):
        lo = max(0, t - 9)
        vol_proxy[t] = np.std(arr[lo : t + 1], axis=0)

    quality_proxy = trailing_5 - vol_proxy
    rates_credit_proxy = -market

    out[:, 0] = market
    out[:, 1] = np.mean(size_proxy * arr, axis=1)
    out[:, 2] = np.mean(value_proxy * arr, axis=1)
    out[:, 3] = np.mean(trailing_5 * arr, axis=1)
    out[:, 4] = np.mean(quality_proxy * arr, axis=1)
    out[:, 5] = np.mean(vol_proxy * arr, axis=1)
    out[:, 6] = rates_credit_proxy
    return np.nan_to_num(out, nan=0.0, posinf=0.0, neginf=0.0)


def _rolling_factor_exposures(*, asset_returns: np.ndarray, factor_returns: np.ndarray, lookback: int) -> np.ndarray:
    r = np.asarray(asset_returns, dtype=float)
    f = np.asarray(factor_returns, dtype=float)
    periods, assets = r.shape
    factors = f.shape[1]
    out = np.zeros((periods, assets, factors), dtype=float)
    lb = max(2, int(lookback))
    for t in range(periods):
        lo = max(0, t - lb + 1)
        x = np.nan_to_num(f[lo : t + 1], nan=0.0, posinf=0.0, neginf=0.0)
        xtx = x.T @ x
        ridge = np.eye(factors) * 1e-6
        inv = np.linalg.pinv(xtx + ridge)
        for a in range(assets):
            y = np.nan_to_num(r[lo : t + 1, a], nan=0.0, posinf=0.0, neginf=0.0)
            out[t, a] = inv @ x.T @ y
    return out

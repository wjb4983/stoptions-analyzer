from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import numpy as np


@dataclass(frozen=True)
class OptionExposureSnapshot:
    delta: np.ndarray
    gamma: np.ndarray
    vega: np.ndarray
    theta: np.ndarray


@dataclass(frozen=True)
class OptionsFeaturePipelineResult:
    features: dict[str, np.ndarray]
    sanity_checks: dict[str, bool]


def aggregate_option_exposures(
    *,
    positions: np.ndarray,
    portfolio_weights: np.ndarray,
    delta: np.ndarray | None = None,
    gamma: np.ndarray | None = None,
    vega: np.ndarray | None = None,
    theta: np.ndarray | None = None,
    multipliers: np.ndarray | None = None,
    contract_adjustments: np.ndarray | None = None,
) -> OptionExposureSnapshot:
    pos = np.asarray(positions, dtype=float)
    if pos.ndim != 2:
        raise ValueError("positions must be 2D [periods, assets]")

    rows, assets = pos.shape
    w = np.asarray(portfolio_weights, dtype=float)
    if w.shape != (assets,):
        raise ValueError("portfolio_weights must have shape (assets,)")

    multi = np.ones((rows, assets), dtype=float)
    if multipliers is not None:
        m = np.asarray(multipliers, dtype=float)
        if m.shape == (assets,):
            multi *= m[None, :]
        elif m.shape == (rows, assets):
            multi *= m

    if contract_adjustments is not None:
        adj = np.asarray(contract_adjustments, dtype=float)
        if adj.shape == (assets,):
            multi *= adj[None, :]
        elif adj.shape == (rows, assets):
            multi *= adj

    signed_notional = pos * w[None, :] * multi

    def _resolve(values: np.ndarray | None) -> np.ndarray:
        if values is None:
            return np.zeros_like(signed_notional)
        arr = np.asarray(values, dtype=float)
        if arr.shape != signed_notional.shape:
            return np.zeros_like(signed_notional)
        return arr

    return OptionExposureSnapshot(
        delta=np.sum(signed_notional * _resolve(delta), axis=1),
        gamma=np.sum(signed_notional * _resolve(gamma), axis=1),
        vega=np.sum(signed_notional * _resolve(vega), axis=1),
        theta=np.sum(signed_notional * _resolve(theta), axis=1),
    )


def _coerce_feature_shape(raw: np.ndarray, reference: np.ndarray) -> np.ndarray:
    arr = np.asarray(raw, dtype=float)
    if arr.shape == reference.shape:
        return arr
    if reference.ndim == 2:
        n_time, n_assets = reference.shape
        if arr.ndim == 1:
            if arr.shape[0] == n_time:
                return np.repeat(arr[:, None], n_assets, axis=1)
            if arr.shape[0] == n_assets:
                return np.repeat(arr[None, :], n_time, axis=0)
        if arr.ndim == 2:
            if arr.shape == (1, n_assets):
                return np.repeat(arr, n_time, axis=0)
            if arr.shape == (n_time, 1):
                return np.repeat(arr, n_assets, axis=1)
    raise ValueError(f"Invalid shape {arr.shape}; expected {reference.shape}")


def _safe_divide(a: np.ndarray, b: np.ndarray) -> np.ndarray:
    return np.divide(a, b, out=np.zeros_like(a, dtype=float), where=np.abs(b) > 1e-12)


def _lag(arr: np.ndarray, periods: int = 1) -> np.ndarray:
    if arr.ndim == 1:
        shifted = np.empty_like(arr)
        shifted[:periods] = arr[0]
        shifted[periods:] = arr[:-periods]
        return shifted
    shifted = np.empty_like(arr)
    shifted[:periods, :] = arr[0:1, :]
    shifted[periods:, :] = arr[:-periods, :]
    return shifted


def _rolling_zscore(arr: np.ndarray, window: int) -> np.ndarray:
    vals = np.asarray(arr, dtype=float)
    z = np.zeros_like(vals, dtype=float)
    if vals.ndim == 1:
        for i in range(vals.shape[0]):
            lo = max(0, i - window + 1)
            sample = vals[lo : i + 1]
            mu = float(np.mean(sample))
            sigma = float(np.std(sample))
            z[i] = 0.0 if sigma <= 1e-12 else (vals[i] - mu) / sigma
        return z

    for i in range(vals.shape[0]):
        lo = max(0, i - window + 1)
        sample = vals[lo : i + 1, :]
        mu = np.mean(sample, axis=0)
        sigma = np.std(sample, axis=0)
        z[i, :] = np.divide(vals[i, :] - mu, sigma, out=np.zeros_like(mu), where=sigma > 1e-12)
    return z


def _cross_sectional_rank(arr: np.ndarray) -> np.ndarray:
    vals = np.asarray(arr, dtype=float)
    if vals.ndim == 1:
        order = np.argsort(vals)
        rank = np.empty_like(vals)
        rank[order] = np.linspace(0.0, 1.0, vals.shape[0])
        return rank

    out = np.zeros_like(vals, dtype=float)
    n_assets = vals.shape[1]
    if n_assets <= 1:
        out[:] = 0.5
        return out
    for i in range(vals.shape[0]):
        order = np.argsort(vals[i, :], kind="mergesort")
        row_rank = np.empty(n_assets, dtype=float)
        row_rank[order] = np.linspace(0.0, 1.0, n_assets)
        out[i, :] = row_rank
    return out


def _validate_market_behavior(features: dict[str, np.ndarray]) -> dict[str, bool]:
    skew = np.asarray(features["skew"], dtype=float)
    flow = np.asarray(features["put_call_flow_imbalance"], dtype=float)
    unusual = np.asarray(features["unusual_volume_signature"], dtype=float)
    gex = np.asarray(features["gamma_exposure_proxy"], dtype=float)

    checks = {
        "skew_mostly_put_rich": float(np.mean(skew > 0.0)) >= 0.5,
        "flow_bounded": bool(np.all((flow >= -1.0) & (flow <= 1.0))),
        "unusual_volume_positive": bool(np.all(unusual >= 0.0)),
        "dealer_proxy_inverse_to_gex": float(np.corrcoef(gex.reshape(-1), np.asarray(features["dealer_positioning_proxy"]).reshape(-1))[0, 1]) <= 0.0,
    }
    return checks


def compute_options_feature_pipeline(
    raw_inputs: dict[str, np.ndarray],
    *,
    rolling_window: int = 20,
) -> OptionsFeaturePipelineResult:
    """Build options-derived features with rolling z-scores and cross-sectional ranks.

    Raw inputs expected: `iv_10d_put`, `iv_25d_put`, `iv_atm`, `iv_25d_call`,
    `iv_10d_call`, `iv_1m`, `iv_3m`, `iv_6m`, `put_volume`, `call_volume`,
    `open_interest`, `net_gamma_notional`, `underlying_market_cap`, `spot_return`.
    Arrays can be shape `(time,)` or `(time, assets)`.
    """

    iv_atm = np.asarray(raw_inputs["iv_atm"], dtype=float)
    iv25p = _coerce_feature_shape(raw_inputs["iv_25d_put"], iv_atm)
    iv25c = _coerce_feature_shape(raw_inputs["iv_25d_call"], iv_atm)
    iv10p = _coerce_feature_shape(raw_inputs["iv_10d_put"], iv_atm)
    iv10c = _coerce_feature_shape(raw_inputs["iv_10d_call"], iv_atm)

    iv1m = _coerce_feature_shape(raw_inputs["iv_1m"], iv_atm)
    iv3m = _coerce_feature_shape(raw_inputs["iv_3m"], iv_atm)
    iv6m = _coerce_feature_shape(raw_inputs["iv_6m"], iv_atm)

    put_volume = _coerce_feature_shape(raw_inputs["put_volume"], iv_atm)
    call_volume = _coerce_feature_shape(raw_inputs["call_volume"], iv_atm)
    open_interest = _coerce_feature_shape(raw_inputs["open_interest"], iv_atm)
    net_gamma_notional = _coerce_feature_shape(raw_inputs["net_gamma_notional"], iv_atm)
    mkt_cap = _coerce_feature_shape(raw_inputs["underlying_market_cap"], iv_atm)
    spot_return = _coerce_feature_shape(raw_inputs["spot_return"], iv_atm)

    total_volume = put_volume + call_volume
    lag_oi = _lag(open_interest)

    skew = iv25p - iv25c
    convexity = 0.5 * (iv25p + iv25c) - iv_atm
    term_structure_curvature = iv6m - 2.0 * iv3m + iv1m
    local_surface_distortion = np.abs(iv10p - 2.0 * iv25p + iv_atm) + np.abs(iv_atm - 2.0 * iv25c + iv10c)
    put_call_flow_imbalance = _safe_divide(put_volume - call_volume, np.maximum(total_volume, 1.0))
    oi_changes = _safe_divide(open_interest - lag_oi, np.maximum(lag_oi, 1.0))
    gamma_exposure_proxy = _safe_divide(net_gamma_notional, np.maximum(mkt_cap, 1.0))
    dealer_positioning_proxy = -(gamma_exposure_proxy * (1.0 + spot_return))
    unusual_volume_signature = _safe_divide(total_volume, np.maximum(_lag(total_volume, periods=max(1, rolling_window // 2)), 1.0))

    feature_map = {
        "skew": skew,
        "convexity": convexity,
        "term_structure_curvature": term_structure_curvature,
        "local_surface_distortion": local_surface_distortion,
        "put_call_flow_imbalance": put_call_flow_imbalance,
        "oi_changes": oi_changes,
        "gamma_exposure_proxy": gamma_exposure_proxy,
        "dealer_positioning_proxy": dealer_positioning_proxy,
        "unusual_volume_signature": unusual_volume_signature,
    }

    enriched: dict[str, np.ndarray] = {}
    for name, values in feature_map.items():
        enriched[name] = values
        enriched[f"{name}_z"] = _rolling_zscore(values, window=max(2, int(rolling_window)))
        enriched[f"{name}_rank"] = _cross_sectional_rank(values)

    return OptionsFeaturePipelineResult(features=enriched, sanity_checks=_validate_market_behavior(enriched))


def summarize_lifecycle_events(events: list[dict[str, Any]]) -> dict[str, int]:
    counts: dict[str, int] = {
        "expiry_cash_settlement": 0,
        "assignment_exercise": 0,
        "contract_roll": 0,
        "corporate_action_adjustment": 0,
    }
    for row in events:
        name = str(row.get("event", "")).strip()
        if name in counts:
            counts[name] += 1
    return counts

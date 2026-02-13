from __future__ import annotations

from dataclasses import dataclass

import numpy as np


@dataclass(frozen=True)
class RegimeClassifierConfig:
    n_states: int = 4
    hmm_iterations: int = 12
    cluster_iterations: int = 20
    eps: float = 1e-12


@dataclass(frozen=True)
class RegimeFeatureInputs:
    vix_front: np.ndarray
    vix_back: np.ndarray
    realized_vol: np.ndarray
    yield_curve_slope: np.ndarray
    breadth: np.ndarray
    liquidity_proxy: np.ndarray
    credit_spread: np.ndarray


def _as_1d(values: np.ndarray, n_obs: int) -> np.ndarray:
    arr = np.asarray(values, dtype=float).reshape(-1)
    if arr.size == n_obs:
        return arr
    if arr.size == 0:
        return np.zeros(n_obs, dtype=float)
    if arr.size > n_obs:
        return arr[:n_obs]
    out = np.zeros(n_obs, dtype=float)
    out[: arr.size] = arr
    out[arr.size :] = arr[-1]
    return out


def _zscore(arr: np.ndarray) -> np.ndarray:
    x = np.asarray(arr, dtype=float)
    mean = float(np.nanmean(x)) if x.size else 0.0
    std = float(np.nanstd(x)) if x.size else 1.0
    if std <= 1e-8:
        std = 1.0
    return (np.where(np.isfinite(x), x, mean) - mean) / std


def build_regime_feature_pipeline(inputs: RegimeFeatureInputs) -> tuple[np.ndarray, tuple[str, ...], dict[str, np.ndarray]]:
    n_obs = int(np.asarray(inputs.realized_vol).reshape(-1).size)
    vix_front = _as_1d(inputs.vix_front, n_obs)
    vix_back = _as_1d(inputs.vix_back, n_obs)
    realized_vol = _as_1d(inputs.realized_vol, n_obs)
    slope = _as_1d(inputs.yield_curve_slope, n_obs)
    breadth = _as_1d(inputs.breadth, n_obs)
    liquidity = _as_1d(inputs.liquidity_proxy, n_obs)
    credit = _as_1d(inputs.credit_spread, n_obs)

    vix_term = vix_back - vix_front
    state = {
        "vix_term_structure": vix_term,
        "realized_vol": realized_vol,
        "yield_curve_slope": slope,
        "breadth": breadth,
        "liquidity_proxy": liquidity,
        "credit_spread": credit,
    }
    features = np.column_stack(
        [
            _zscore(vix_term),
            _zscore(realized_vol),
            _zscore(slope),
            _zscore(breadth),
            _zscore(liquidity),
            _zscore(credit),
        ]
    )
    names = (
        "vix_term_structure",
        "realized_vol",
        "yield_curve_slope",
        "breadth",
        "liquidity_proxy",
        "credit_spread",
    )
    return features, names, state


def _fit_cluster_posteriors(features: np.ndarray, n_states: int, n_iter: int, eps: float) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    x = np.asarray(features, dtype=float)
    n_obs, n_features = x.shape
    scores = np.mean(x, axis=1)
    quantiles = np.linspace(0.0, 1.0, n_states + 2)[1:-1]
    seeds = np.quantile(scores, quantiles)
    centers = np.column_stack([seeds for _ in range(n_features)])
    variances = np.full((n_states, n_features), 1.0, dtype=float)
    weights = np.full(n_states, 1.0 / n_states, dtype=float)

    resp = np.full((n_obs, n_states), 1.0 / n_states, dtype=float)
    for _ in range(max(1, int(n_iter))):
        logp = np.zeros((n_obs, n_states), dtype=float)
        for k in range(n_states):
            var = np.maximum(variances[k], 1e-4)
            diff = x - centers[k]
            ll = -0.5 * np.sum((diff * diff) / var + np.log(2.0 * np.pi * var), axis=1)
            logp[:, k] = np.log(max(weights[k], eps)) + ll
        logp -= np.max(logp, axis=1, keepdims=True)
        p = np.exp(logp)
        resp = p / np.maximum(np.sum(p, axis=1, keepdims=True), eps)

        nk = np.sum(resp, axis=0)
        weights = np.maximum(nk / max(float(n_obs), 1.0), eps)
        weights /= np.sum(weights)
        for k in range(n_states):
            denom = max(float(nk[k]), eps)
            centers[k] = np.sum(resp[:, [k]] * x, axis=0) / denom
            diff = x - centers[k]
            variances[k] = np.maximum(np.sum(resp[:, [k]] * (diff * diff), axis=0) / denom, 1e-4)
    return resp, centers, variances


def _smooth_posteriors_hmm(resp: np.ndarray, iterations: int, eps: float) -> np.ndarray:
    posterior = np.array(resp, copy=True)
    n_obs, n_states = posterior.shape
    transition = np.full((n_states, n_states), 1.0 / n_states, dtype=float)
    if n_obs > 1:
        for _ in range(max(1, int(iterations))):
            transition = np.full((n_states, n_states), eps, dtype=float)
            for t in range(1, n_obs):
                transition += np.outer(posterior[t - 1], posterior[t])
            transition /= np.maximum(np.sum(transition, axis=1, keepdims=True), eps)

            forward = np.array(resp, copy=True)
            forward[0] /= np.maximum(np.sum(forward[0]), eps)
            for t in range(1, n_obs):
                forward[t] *= forward[t - 1] @ transition
                forward[t] /= np.maximum(np.sum(forward[t]), eps)

            backward = np.ones_like(resp)
            for t in range(n_obs - 2, -1, -1):
                backward[t] = transition @ (resp[t + 1] * backward[t + 1])
                backward[t] /= np.maximum(np.sum(backward[t]), eps)

            posterior = forward * backward
            posterior /= np.maximum(np.sum(posterior, axis=1, keepdims=True), eps)
    return posterior


def classify_regimes(
    *,
    features: np.ndarray,
    feature_names: tuple[str, ...],
    raw_feature_map: dict[str, np.ndarray],
    config: RegimeClassifierConfig | None = None,
) -> dict[str, np.ndarray]:
    cfg = config or RegimeClassifierConfig()
    x = np.asarray(features, dtype=float)
    if x.ndim != 2:
        raise ValueError("features must be 2D")
    n_obs = x.shape[0]
    n_states = max(2, int(cfg.n_states))
    if n_obs == 0:
        return {
            "regime_probabilities": np.zeros((0, n_states), dtype=float),
            "regime_confidence": np.zeros(0, dtype=float),
            "regime_state_argmax": np.zeros(0, dtype=int),
            "regime_states": np.asarray([f"state_{i}" for i in range(n_states)], dtype=object),
            "regime_labels": np.asarray([], dtype=object),
            "macro_state": np.asarray([], dtype=object),
            "volatility_state": np.asarray([], dtype=object),
            "state_macro_labels": np.asarray(["neutral"] * n_states, dtype=object),
            "state_volatility_labels": np.asarray(["moderate"] * n_states, dtype=object),
            "feature_names": np.asarray(feature_names, dtype=object),
            "feature_matrix": x,
        }

    resp, centers, _ = _fit_cluster_posteriors(x, n_states=n_states, n_iter=cfg.cluster_iterations, eps=cfg.eps)
    posterior = _smooth_posteriors_hmm(resp, iterations=cfg.hmm_iterations, eps=cfg.eps)
    confidence = np.max(posterior, axis=1)
    argmax = np.argmax(posterior, axis=1)

    vix = _zscore(raw_feature_map.get("vix_term_structure", np.zeros(n_obs, dtype=float)))
    rv = _zscore(raw_feature_map.get("realized_vol", np.zeros(n_obs, dtype=float)))
    slope = _zscore(raw_feature_map.get("yield_curve_slope", np.zeros(n_obs, dtype=float)))
    breadth = _zscore(raw_feature_map.get("breadth", np.zeros(n_obs, dtype=float)))
    liquidity = _zscore(raw_feature_map.get("liquidity_proxy", np.zeros(n_obs, dtype=float)))
    credit = _zscore(raw_feature_map.get("credit_spread", np.zeros(n_obs, dtype=float)))

    state_macro_labels: list[str] = []
    state_vol_labels: list[str] = []
    for k in range(n_states):
        w = posterior[:, k]
        denom = max(float(np.sum(w)), cfg.eps)
        macro_score = float(np.sum(w * (0.45 * slope + 0.35 * breadth - 0.20 * credit)) / denom)
        vol_score = float(np.sum(w * (0.40 * rv - 0.30 * vix + 0.30 * liquidity + 0.20 * credit)) / denom)
        state_macro_labels.append("risk_on" if macro_score >= 0.0 else "risk_off")
        if vol_score > 0.45:
            state_vol_labels.append("high")
        elif vol_score < -0.25:
            state_vol_labels.append("low")
        else:
            state_vol_labels.append("mid")

    macro_state = np.asarray([state_macro_labels[idx] for idx in argmax], dtype=object)
    vol_state = np.asarray([state_vol_labels[idx] for idx in argmax], dtype=object)
    labels = np.asarray(
        [f"state_{idx}|macro_{state_macro_labels[idx]}|vol_{state_vol_labels[idx]}" for idx in argmax],
        dtype=object,
    )

    return {
        "regime_probabilities": posterior,
        "regime_confidence": confidence,
        "regime_state_argmax": argmax,
        "regime_states": np.asarray([f"state_{i}" for i in range(n_states)], dtype=object),
        "regime_labels": labels,
        "macro_state": macro_state,
        "volatility_state": vol_state,
        "state_macro_labels": np.asarray(state_macro_labels, dtype=object),
        "state_volatility_labels": np.asarray(state_vol_labels, dtype=object),
        "state_centers": centers,
        "feature_names": np.asarray(feature_names, dtype=object),
        "feature_matrix": x,
    }

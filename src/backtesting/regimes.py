from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import numpy as np


@dataclass(frozen=True)
class RegimeFeatureConfig:
    trend_fast_window: int = 10
    trend_slow_window: int = 40
    vol_window: int = 20
    liquidity_window: int = 20
    macro_window: int = 30
    term_structure_short_window: int = 10
    term_structure_long_window: int = 60
    dispersion_window: int = 20
    correlation_window: int = 20
    breadth_persistence_window: int = 15
    n_states: int = 4
    em_iterations: int = 25
    enabled_features: tuple[str, ...] = (
        "trend_spread",
        "market_vol",
        "term_vol_spread",
        "cross_sectional_dispersion",
        "correlation_stress",
        "breadth_persistence",
    )


def _rolling_mean(values: np.ndarray, window: int) -> np.ndarray:
    arr = np.asarray(values, dtype=float)
    out = np.zeros_like(arr, dtype=float)
    if arr.size == 0:
        return out

    win = max(1, int(window))
    csum = np.cumsum(arr, dtype=float)
    counts = np.minimum(np.arange(1, arr.size + 1, dtype=int), win).astype(float)

    out[:] = csum
    if win < arr.size:
        out[win:] = csum[win:] - csum[:-win]
    out /= counts
    return out


def _rolling_std(values: np.ndarray, window: int) -> np.ndarray:
    arr = np.asarray(values, dtype=float)
    out = np.zeros_like(arr, dtype=float)
    if arr.size == 0:
        return out

    win = max(1, int(window))
    csum = np.cumsum(arr, dtype=float)
    csum_sq = np.cumsum(arr * arr, dtype=float)
    counts = np.minimum(np.arange(1, arr.size + 1, dtype=int), win)

    sum_x = csum.copy()
    sum_x_sq = csum_sq.copy()
    if win < arr.size:
        sum_x[win:] = csum[win:] - csum[:-win]
        sum_x_sq[win:] = csum_sq[win:] - csum_sq[:-win]

    valid = counts > 1
    numer = np.zeros_like(arr, dtype=float)
    numer[valid] = sum_x_sq[valid] - (sum_x[valid] * sum_x[valid] / counts[valid].astype(float))
    denom = np.maximum(counts.astype(float) - 1.0, 1.0)
    out[valid] = np.sqrt(np.maximum(numer[valid] / denom[valid], 0.0))
    return out


def compute_regime_labels(
    prices: np.ndarray,
    *,
    config: RegimeFeatureConfig | None = None,
) -> dict[str, np.ndarray]:
    cfg = config or RegimeFeatureConfig()
    px = np.asarray(prices, dtype=float)
    if px.ndim != 2:
        raise ValueError("prices must be 2D")

    # Observable proxies only: cross-sectional average close-to-close return profile.
    rets = np.zeros(px.shape, dtype=float)
    rets[1:] = px[1:] / np.where(px[:-1] == 0.0, 1.0, px[:-1]) - 1.0
    market_ret = np.nanmean(np.where(np.isfinite(rets), rets, 0.0), axis=1)

    trend_fast = _rolling_mean(market_ret, max(2, int(cfg.trend_fast_window)))
    trend_slow = _rolling_mean(market_ret, max(3, int(cfg.trend_slow_window)))
    trend_spread = trend_fast - trend_slow
    trend = np.where(trend_fast > trend_slow, "up", "down")

    vol = _rolling_std(market_ret, max(2, int(cfg.vol_window)))
    vol_anchor = _rolling_mean(vol, max(3, int(cfg.vol_window * 2)))
    vol_state = np.where(vol > vol_anchor, "high", "low")

    term_short = _rolling_std(market_ret, max(2, int(cfg.term_structure_short_window)))
    term_long = _rolling_std(market_ret, max(3, int(cfg.term_structure_long_window)))
    term_vol_spread = term_short - term_long

    abs_ret = np.abs(market_ret)
    liq_proxy = _rolling_mean(abs_ret, max(2, int(cfg.liquidity_window)))
    liq_anchor = _rolling_mean(liq_proxy, max(3, int(cfg.liquidity_window * 2)))
    liquidity = np.where(liq_proxy > liq_anchor, "thin", "normal")

    breadth = np.nanmean((rets > 0).astype(float), axis=1)
    macro_breadth = _rolling_mean(breadth, max(2, int(cfg.macro_window)))
    macro = np.where(macro_breadth >= 0.5, "risk_on", "risk_off")
    breadth_persistence = np.abs(_rolling_mean((breadth >= 0.5).astype(float), max(2, int(cfg.breadth_persistence_window))) - 0.5) * 2.0

    dispersion = _rolling_mean(np.nanstd(np.where(np.isfinite(rets), rets, 0.0), axis=1), max(2, int(cfg.dispersion_window)))
    correlation_stress = _rolling_correlation_stress(np.where(np.isfinite(rets), rets, 0.0), max(2, int(cfg.correlation_window)))

    feature_map = {
        "trend_spread": trend_spread,
        "market_vol": vol,
        "term_vol_spread": term_vol_spread,
        "cross_sectional_dispersion": dispersion,
        "correlation_stress": correlation_stress,
        "breadth_persistence": breadth_persistence,
        "liquidity_proxy": liq_proxy,
        "macro_breadth": macro_breadth,
    }
    feature_names = [name for name in cfg.enabled_features if name in feature_map]
    if not feature_names:
        feature_names = ["trend_spread", "market_vol", "term_vol_spread", "cross_sectional_dispersion"]

    feature_matrix = np.column_stack([feature_map[name] for name in feature_names]) if px.shape[0] else np.zeros((0, len(feature_names)), dtype=float)
    posterior, state_names = _fit_probabilistic_regime_model(
        features=feature_matrix,
        n_states=max(2, int(cfg.n_states)),
        em_iterations=max(1, int(cfg.em_iterations)),
    )

    component_labels = np.array(
        [
            f"trend_{trend[idx]}|vol_{vol_state[idx]}|liq_{liquidity[idx]}|macro_{macro[idx]}"
            for idx in range(px.shape[0])
        ],
        dtype=object,
    )

    argmax_states = np.argmax(posterior, axis=1) if posterior.size else np.zeros(px.shape[0], dtype=int)
    mapped_labels = _map_states_to_legacy_labels(
        posterior=posterior,
        state_names=state_names,
        component_labels=component_labels,
    )
    labels = np.array([mapped_labels[state_names[idx]] for idx in argmax_states], dtype=object)
    return {
        "trend": trend,
        "volatility": vol_state,
        "liquidity": liquidity,
        "macro": macro,
        "labels": labels,
        "legacy_component_labels": component_labels,
        "regime_probabilities": posterior,
        "regime_states": np.asarray(state_names, dtype=object),
        "regime_state_argmax": np.asarray(argmax_states, dtype=int),
        "regime_state_to_legacy_label": np.asarray([mapped_labels[name] for name in state_names], dtype=object),
        "feature_names": np.asarray(feature_names, dtype=object),
        "feature_matrix": feature_matrix,
    }


def _rolling_correlation_stress(returns: np.ndarray, window: int) -> np.ndarray:
    arr = np.asarray(returns, dtype=float)
    out = np.zeros(arr.shape[0], dtype=float)
    if arr.ndim != 2 or arr.shape[0] == 0:
        return out
    win = max(2, int(window))
    for idx in range(arr.shape[0]):
        start = max(0, idx - win + 1)
        chunk = arr[start : idx + 1]
        if chunk.shape[0] < 2 or chunk.shape[1] < 2:
            continue
        corr = np.corrcoef(chunk, rowvar=False)
        if corr.ndim != 2:
            continue
        triu = np.triu_indices(corr.shape[0], k=1)
        vals = corr[triu]
        vals = vals[np.isfinite(vals)]
        if vals.size:
            out[idx] = float(np.mean(np.abs(vals)))
    return out


def _fit_probabilistic_regime_model(*, features: np.ndarray, n_states: int, em_iterations: int) -> tuple[np.ndarray, list[str]]:
    x = np.asarray(features, dtype=float)
    if x.ndim != 2:
        raise ValueError("features must be 2D")
    n_obs = x.shape[0]
    if n_obs == 0:
        return np.zeros((0, n_states), dtype=float), [f"state_{i}" for i in range(n_states)]
    mean = np.nanmean(x, axis=0)
    std = np.nanstd(x, axis=0)
    std = np.where(std <= 1e-8, 1.0, std)
    z = (np.where(np.isfinite(x), x, mean) - mean) / std

    qs = np.linspace(0.0, 1.0, n_states + 2)[1:-1]
    scores = np.mean(z, axis=1)
    centers_1d = np.quantile(scores, qs)
    centers = np.column_stack([centers_1d for _ in range(z.shape[1])])
    variances = np.full((n_states, z.shape[1]), 1.0, dtype=float)
    weights = np.full(n_states, 1.0 / n_states, dtype=float)

    resp = np.full((n_obs, n_states), 1.0 / n_states, dtype=float)
    for _ in range(em_iterations):
        logp = np.zeros_like(resp)
        for k in range(n_states):
            var = np.maximum(variances[k], 1e-4)
            diff = z - centers[k]
            ll = -0.5 * np.sum((diff * diff) / var + np.log(2.0 * np.pi * var), axis=1)
            logp[:, k] = np.log(max(weights[k], 1e-9)) + ll
        logp -= np.max(logp, axis=1, keepdims=True)
        p = np.exp(logp)
        resp = p / np.maximum(np.sum(p, axis=1, keepdims=True), 1e-12)

        nk = np.sum(resp, axis=0)
        weights = np.maximum(nk / max(float(n_obs), 1.0), 1e-6)
        weights /= np.sum(weights)
        for k in range(n_states):
            denom = max(nk[k], 1e-6)
            centers[k] = np.sum(resp[:, [k]] * z, axis=0) / denom
            diff = z - centers[k]
            variances[k] = np.maximum(np.sum(resp[:, [k]] * (diff * diff), axis=0) / denom, 1e-4)

    transition = np.full((n_states, n_states), 1.0 / n_states, dtype=float)
    if n_obs > 1:
        for t in range(1, n_obs):
            transition += np.outer(resp[t - 1], resp[t])
        transition /= np.maximum(np.sum(transition, axis=1, keepdims=True), 1e-12)

    posterior = np.array(resp, copy=True)
    for t in range(1, n_obs):
        posterior[t] = posterior[t] * (posterior[t - 1] @ transition)
        posterior[t] /= np.maximum(np.sum(posterior[t]), 1e-12)

    state_names = [f"state_{i}" for i in range(n_states)]
    return posterior, state_names


def _map_states_to_legacy_labels(
    *,
    posterior: np.ndarray,
    state_names: list[str],
    component_labels: np.ndarray,
) -> dict[str, str]:
    label_space = sorted(set(str(x) for x in component_labels.tolist()))
    if not label_space:
        return {name: "" for name in state_names}
    mapping: dict[str, str] = {}
    for idx, name in enumerate(state_names):
        best_label = label_space[0]
        best_score = -np.inf
        for label in label_space:
            mask = component_labels == label
            score = float(np.sum(posterior[mask, idx])) if posterior.size else 0.0
            if score > best_score:
                best_score = score
                best_label = label
        mapping[name] = best_label
    return mapping


def resolve_regime_parameters(
    *,
    base_params: dict[str, Any],
    regime_label: str,
    parameter_map: dict[str, dict[str, Any]] | None,
) -> dict[str, Any]:
    resolved = dict(base_params)
    if not parameter_map:
        return resolved
    if "default" in parameter_map:
        resolved.update(parameter_map["default"])
    if regime_label in parameter_map:
        resolved.update(parameter_map[regime_label])
    return resolved


def apply_regime_risk_overlays(
    *,
    weights: np.ndarray,
    regime_labels: np.ndarray,
    risk_map: dict[str, dict[str, float]] | None,
) -> tuple[np.ndarray, dict[str, np.ndarray]]:
    arr = np.asarray(weights, dtype=float)
    labels = np.asarray(regime_labels, dtype=object)
    adjusted = np.array(arr, copy=True)
    leverage_caps = np.full(arr.shape[0], np.nan, dtype=float)
    multipliers = np.ones(arr.shape[0], dtype=float)

    if not risk_map:
        return adjusted, {"regime_leverage_cap": leverage_caps, "regime_risk_multiplier": multipliers}

    default_cfg = risk_map.get("default", {})
    for idx in range(arr.shape[0]):
        cfg = dict(default_cfg)
        if labels[idx] in risk_map:
            cfg.update(risk_map[str(labels[idx])])

        leverage_cap = float(cfg.get("max_gross_exposure", np.inf))
        risk_mult = float(cfg.get("risk_multiplier", 1.0))
        multipliers[idx] = risk_mult
        row = adjusted[idx] * risk_mult
        gross = float(np.sum(np.abs(row)))
        if np.isfinite(leverage_cap):
            leverage_caps[idx] = leverage_cap
            if gross > leverage_cap and gross > 0.0:
                row *= leverage_cap / gross
        adjusted[idx] = row

    return adjusted, {"regime_leverage_cap": leverage_caps, "regime_risk_multiplier": multipliers}


class RegimeScaledSlippageModel:
    def __init__(
        self,
        base_model: Any,
        regime_labels: np.ndarray,
        multipliers: dict[str, float] | None,
    ) -> None:
        self.base_model = base_model
        self.labels = np.asarray(regime_labels, dtype=object)
        self.multipliers = dict(multipliers or {})
        self.default_multiplier = float(self.multipliers.get("default", 1.0))

    def apply(self, price: float, size: float, liquidity_context: Any | None = None) -> float:
        ctx = liquidity_context
        bar_index = int(getattr(ctx, "bar_index", 0)) if ctx is not None else 0
        if bar_index < 0 or bar_index >= self.labels.size:
            scale = self.default_multiplier
        else:
            scale = float(self.multipliers.get(str(self.labels[bar_index]), self.default_multiplier))
        scaled_size = float(size) * max(scale, 0.0)
        return float(self.base_model.apply(price, scaled_size, liquidity_context))


def attribute_pnl_by_regime(*, pnl: np.ndarray, regime_labels: np.ndarray) -> list[dict[str, float | str | int]]:
    pnl_arr = np.asarray(pnl, dtype=float)
    labels = np.asarray(regime_labels, dtype=object)
    grouped: dict[str, list[float]] = {}
    for idx in range(min(pnl_arr.size, labels.size)):
        grouped.setdefault(str(labels[idx]), []).append(float(pnl_arr[idx]))

    rows: list[dict[str, float | str | int]] = []
    for label in sorted(grouped):
        vals = np.asarray(grouped[label], dtype=float)
        rows.append(
            {
                "regime": label,
                "bars": int(vals.size),
                "pnl_total": float(np.sum(vals)),
                "pnl_mean": float(np.mean(vals)) if vals.size else 0.0,
            }
        )
    return rows

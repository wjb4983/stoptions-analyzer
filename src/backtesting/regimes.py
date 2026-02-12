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


def _rolling_mean(values: np.ndarray, window: int) -> np.ndarray:
    out = np.zeros_like(values, dtype=float)
    for idx in range(values.size):
        start = max(0, idx - window + 1)
        out[idx] = float(np.mean(values[start : idx + 1]))
    return out


def _rolling_std(values: np.ndarray, window: int) -> np.ndarray:
    out = np.zeros_like(values, dtype=float)
    for idx in range(values.size):
        start = max(0, idx - window + 1)
        segment = values[start : idx + 1]
        out[idx] = float(np.std(segment, ddof=1)) if segment.size > 1 else 0.0
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
    trend = np.where(trend_fast > trend_slow, "up", "down")

    vol = _rolling_std(market_ret, max(2, int(cfg.vol_window)))
    vol_anchor = _rolling_mean(vol, max(3, int(cfg.vol_window * 2)))
    vol_state = np.where(vol > vol_anchor, "high", "low")

    abs_ret = np.abs(market_ret)
    liq_proxy = _rolling_mean(abs_ret, max(2, int(cfg.liquidity_window)))
    liq_anchor = _rolling_mean(liq_proxy, max(3, int(cfg.liquidity_window * 2)))
    liquidity = np.where(liq_proxy > liq_anchor, "thin", "normal")

    breadth = np.nanmean((rets > 0).astype(float), axis=1)
    macro_breadth = _rolling_mean(breadth, max(2, int(cfg.macro_window)))
    macro = np.where(macro_breadth >= 0.5, "risk_on", "risk_off")

    labels = np.array(
        [
            f"trend_{trend[idx]}|vol_{vol_state[idx]}|liq_{liquidity[idx]}|macro_{macro[idx]}"
            for idx in range(px.shape[0])
        ],
        dtype=object,
    )
    return {
        "trend": trend,
        "volatility": vol_state,
        "liquidity": liquidity,
        "macro": macro,
        "labels": labels,
    }


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

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import numpy as np


PortfolioMethod = Literal["equal_weight", "vol_target", "inverse_vol", "capped_optimization"]


@dataclass(frozen=True)
class PortfolioConstructionConfig:
    method: PortfolioMethod = "equal_weight"
    vol_lookback_bars: int = 20
    target_volatility: float = 0.10
    max_symbol_weight: float = 0.25
    max_sector_weight: float = 0.60
    max_gross_exposure: float = 1.0
    min_net_exposure: float = -1.0
    max_net_exposure: float = 1.0
    sector_map: dict[str, str] | None = None


@dataclass(frozen=True)
class PortfolioConstructionResult:
    target_weights: np.ndarray
    diagnostics: dict[str, np.ndarray]


def construct_target_weights(
    *,
    raw_signals: np.ndarray,
    prices: np.ndarray,
    symbol_order: list[str],
    config: PortfolioConstructionConfig,
) -> PortfolioConstructionResult:
    signals = np.asarray(raw_signals, dtype=float)
    px = np.asarray(prices, dtype=float)
    if signals.shape != px.shape:
        raise ValueError("raw_signals and prices must have same shape")

    periods, assets = signals.shape
    weights = np.zeros_like(signals, dtype=float)
    turnover_by_symbol = np.zeros_like(signals, dtype=float)
    prev = np.zeros(assets, dtype=float)
    sector_map = config.sector_map or {}

    for idx in range(periods):
        row = np.nan_to_num(signals[idx], nan=0.0, posinf=0.0, neginf=0.0)
        tradable = np.isfinite(px[idx]) & (px[idx] > 0.0)
        row[~tradable] = 0.0

        base = _build_base_weights(
            row=row,
            prices=px,
            idx=idx,
            config=config,
            tradable=tradable,
        )
        constrained = _apply_constraints(
            weights=base,
            symbol_order=symbol_order,
            sector_map=sector_map,
            config=config,
        )
        constrained[~tradable] = 0.0
        weights[idx] = constrained
        turnover_by_symbol[idx] = np.abs(constrained - prev)
        prev = constrained

    gross = np.sum(np.abs(weights), axis=1)
    net = np.sum(weights, axis=1)
    concentration = np.zeros(periods, dtype=float)
    nonzero = gross > 0
    concentration[nonzero] = np.sum((np.abs(weights[nonzero]) / gross[nonzero, None]) ** 2, axis=1)
    leverage_usage = np.divide(
        gross,
        max(config.max_gross_exposure, 1e-12),
        out=np.zeros_like(gross),
        where=np.isfinite(gross),
    )

    diagnostics = {
        "gross_exposure": gross,
        "net_exposure": net,
        "concentration": concentration,
        "leverage_usage": leverage_usage,
        "turnover": np.sum(turnover_by_symbol, axis=1),
        "turnover_by_symbol": turnover_by_symbol,
    }
    return PortfolioConstructionResult(target_weights=weights, diagnostics=diagnostics)


def _build_base_weights(
    *,
    row: np.ndarray,
    prices: np.ndarray,
    idx: int,
    config: PortfolioConstructionConfig,
    tradable: np.ndarray,
) -> np.ndarray:
    active = np.flatnonzero((row != 0.0) & tradable)
    if active.size == 0:
        return np.zeros_like(row)

    signs = np.sign(row)
    if config.method == "equal_weight":
        base = np.zeros_like(row)
        base[active] = signs[active]
        return _normalize_gross(base)

    vol = _rolling_volatility(prices=prices, idx=idx, lookback=max(2, int(config.vol_lookback_bars)))
    inv_vol = np.zeros_like(row)
    valid = (vol > 0.0) & np.isfinite(vol)
    inv_vol[valid] = 1.0 / vol[valid]

    if config.method in {"inverse_vol", "capped_optimization"}:
        base = np.zeros_like(row)
        base[active] = signs[active] * inv_vol[active]
        if np.sum(np.abs(base)) <= 0:
            base[active] = signs[active]
        return _normalize_gross(base)

    if config.method == "vol_target":
        base = np.zeros_like(row)
        base[active] = signs[active]
        base = _normalize_gross(base)
        if np.sum(np.abs(base)) <= 0:
            return base
        portfolio_vol = float(np.sqrt(np.sum((base * np.nan_to_num(vol, nan=0.0)) ** 2)))
        if portfolio_vol <= 0:
            return base
        scaled = base * (float(config.target_volatility) / portfolio_vol)
        return scaled

    raise ValueError(f"Unknown portfolio method: {config.method}")


def _apply_constraints(
    *,
    weights: np.ndarray,
    symbol_order: list[str],
    sector_map: dict[str, str],
    config: PortfolioConstructionConfig,
) -> np.ndarray:
    out = np.array(weights, dtype=float, copy=True)

    max_symbol = max(0.0, float(config.max_symbol_weight))
    if max_symbol > 0:
        out = np.clip(out, -max_symbol, max_symbol)

    if float(config.max_sector_weight) > 0 and sector_map:
        sectors: dict[str, list[int]] = {}
        for idx, symbol in enumerate(symbol_order):
            sector = sector_map.get(symbol)
            if sector:
                sectors.setdefault(sector, []).append(idx)
        for idxs in sectors.values():
            gross_sector = float(np.sum(np.abs(out[idxs])))
            cap = float(config.max_sector_weight)
            if gross_sector > cap > 0:
                scale = cap / gross_sector
                out[idxs] *= scale

    gross = float(np.sum(np.abs(out)))
    max_gross = max(0.0, float(config.max_gross_exposure))
    if gross > 0 and max_gross > 0 and gross > max_gross:
        out *= max_gross / gross

    net = float(np.sum(out))
    min_net = float(config.min_net_exposure)
    max_net = float(config.max_net_exposure)
    if max_net < min_net:
        min_net, max_net = max_net, min_net

    target_net = min(max(net, min_net), max_net)
    if abs(target_net - net) > 1e-12:
        long_mask = out > 0
        short_mask = out < 0
        if target_net < net and np.any(long_mask):
            # reduce long side
            reduction = net - target_net
            long_sum = float(np.sum(out[long_mask]))
            if long_sum > 0:
                out[long_mask] *= max(0.0, 1.0 - reduction / long_sum)
        elif target_net > net and np.any(short_mask):
            # reduce short side magnitude
            increase = target_net - net
            short_mag = float(np.sum(np.abs(out[short_mask])))
            if short_mag > 0:
                out[short_mask] *= max(0.0, 1.0 - increase / short_mag)

    return out


def _rolling_volatility(*, prices: np.ndarray, idx: int, lookback: int) -> np.ndarray:
    assets = prices.shape[1]
    if idx <= 0:
        return np.zeros(assets, dtype=float)
    start = max(1, idx - lookback + 1)
    window = prices[start - 1 : idx + 1]
    prev = window[:-1]
    curr = window[1:]
    rets = np.full_like(curr, np.nan, dtype=float)
    valid = np.isfinite(prev) & np.isfinite(curr) & (prev > 0)
    rets[valid] = curr[valid] / prev[valid] - 1.0
    vol = np.zeros(assets, dtype=float)
    for col in range(assets):
        series = rets[:, col]
        series = series[np.isfinite(series)]
        if series.size >= 2:
            sigma = float(np.std(series))
            if np.isfinite(sigma) and sigma > 0:
                vol[col] = sigma
    return vol


def _normalize_gross(values: np.ndarray) -> np.ndarray:
    gross = float(np.sum(np.abs(values)))
    if gross <= 0:
        return np.zeros_like(values)
    return values / gross

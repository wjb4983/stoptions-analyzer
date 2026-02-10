from __future__ import annotations

from dataclasses import dataclass

import numpy as np


@dataclass(frozen=True)
class CrossSectionalMomentumConfig:
    lookback_days: int = 90
    skip_days: int = 5
    top_quantile: float = 0.2
    bottom_quantile: float = 0.2
    long_only: bool = False
    vol_lookback_days: int = 20
    rebalance_interval: int = 1


def _compute_momentum_scores(
    *,
    close_prices: np.ndarray,
    missing_mask: np.ndarray,
    lookback_days: int,
    skip_days: int,
) -> np.ndarray:
    periods, assets = close_prices.shape
    scores = np.full((periods, assets), np.nan, dtype=float)
    for idx in range(periods):
        end_idx = idx - skip_days
        start_idx = end_idx - lookback_days
        if start_idx < 0 or end_idx < 0:
            continue
        start_px = close_prices[start_idx]
        end_px = close_prices[end_idx]
        valid = (
            np.isfinite(start_px)
            & np.isfinite(end_px)
            & (start_px > 0.0)
            & (end_px > 0.0)
            & (~missing_mask[start_idx])
            & (~missing_mask[end_idx])
        )
        row = np.full(assets, np.nan, dtype=float)
        row[valid] = end_px[valid] / start_px[valid] - 1.0
        scores[idx] = row
    return scores


def assign_rank_buckets(
    scores_row: np.ndarray,
    *,
    top_quantile: float,
    bottom_quantile: float,
    long_only: bool,
) -> tuple[np.ndarray, np.ndarray]:
    top = np.zeros_like(scores_row, dtype=bool)
    bottom = np.zeros_like(scores_row, dtype=bool)

    valid_idx = np.flatnonzero(np.isfinite(scores_row))
    n = int(valid_idx.size)
    if n == 0:
        return top, bottom

    ranked = valid_idx[np.argsort(scores_row[valid_idx])]
    top_n = int(np.ceil(n * max(0.0, min(1.0, top_quantile))))
    bot_n = int(np.ceil(n * max(0.0, min(1.0, bottom_quantile))))

    if top_n > 0:
        top[ranked[-top_n:]] = True
    if (not long_only) and bot_n > 0:
        bottom[ranked[:bot_n]] = True
        overlap = top & bottom
        if np.any(overlap):
            bottom[overlap] = False
    return top, bottom


def _compute_inverse_volatility(
    *,
    close_prices: np.ndarray,
    missing_mask: np.ndarray,
    vol_lookback_days: int,
) -> np.ndarray:
    periods, assets = close_prices.shape
    inv_vol = np.zeros((periods, assets), dtype=float)
    for idx in range(periods):
        start = idx - vol_lookback_days + 1
        if start <= 0:
            continue
        window = close_prices[start - 1 : idx + 1]
        window_missing = missing_mask[start - 1 : idx + 1]
        rets = np.full((window.shape[0] - 1, assets), np.nan, dtype=float)
        prev = window[:-1]
        curr = window[1:]
        valid = (
            np.isfinite(prev)
            & np.isfinite(curr)
            & (prev > 0.0)
            & (~window_missing[:-1])
            & (~window_missing[1:])
        )
        rets[valid] = curr[valid] / prev[valid] - 1.0
        vol = np.zeros(assets, dtype=float)
        valid_vol = np.zeros(assets, dtype=bool)
        for col in range(assets):
            col_vals = rets[:, col]
            finite = col_vals[np.isfinite(col_vals)]
            if finite.size < 2:
                continue
            sigma = float(np.std(finite))
            if sigma > 0.0 and np.isfinite(sigma):
                vol[col] = sigma
                valid_vol[col] = True
        row = np.zeros(assets, dtype=float)
        row[valid_vol] = 1.0 / vol[valid_vol]
        inv_vol[idx] = row
    return inv_vol


def compute_risk_normalized_weights(raw_weights: np.ndarray) -> np.ndarray:
    values = np.asarray(raw_weights, dtype=float)
    gross = float(np.sum(np.abs(values)))
    if gross <= 0.0:
        return np.zeros_like(values)
    return values / gross


def build_cross_sectional_momentum_targets(
    *,
    close_prices: np.ndarray,
    missing_mask: np.ndarray,
    config: CrossSectionalMomentumConfig,
) -> np.ndarray:
    scores = _compute_momentum_scores(
        close_prices=close_prices,
        missing_mask=missing_mask,
        lookback_days=config.lookback_days,
        skip_days=config.skip_days,
    )
    inv_vol = _compute_inverse_volatility(
        close_prices=close_prices,
        missing_mask=missing_mask,
        vol_lookback_days=config.vol_lookback_days,
    )

    periods, assets = close_prices.shape
    targets = np.zeros((periods, assets), dtype=float)
    last = np.zeros(assets, dtype=float)

    rebalance_interval = max(1, int(config.rebalance_interval))
    for idx in range(periods):
        if idx % rebalance_interval != 0:
            targets[idx] = last
            continue

        top, bottom = assign_rank_buckets(
            scores[idx],
            top_quantile=config.top_quantile,
            bottom_quantile=config.bottom_quantile,
            long_only=config.long_only,
        )

        raw = np.zeros(assets, dtype=float)
        raw[top] = inv_vol[idx, top]
        if not config.long_only:
            raw[bottom] = -inv_vol[idx, bottom]

        # fallback equal weights if vol unavailable for selected names
        if np.sum(np.abs(raw)) <= 0.0:
            if np.any(top):
                raw[top] = 1.0
            if (not config.long_only) and np.any(bottom):
                raw[bottom] = -1.0

        normalized = compute_risk_normalized_weights(raw)
        targets[idx] = normalized
        last = normalized

    return targets

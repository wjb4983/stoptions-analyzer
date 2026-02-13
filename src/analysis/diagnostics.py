from __future__ import annotations

from typing import Any, Iterable, Mapping

import numpy as np

MANDATORY_DIAGNOSTIC_SECTIONS: tuple[str, ...] = (
    "information_coefficient",
    "rank_stability",
    "exposure",
)


def compute_signal_diagnostics(
    *,
    scores: Mapping[str, float],
    weights: Mapping[str, float],
    prices_by_ticker: Mapping[str, list[float] | list[dict[str, float]] | tuple[float, ...]],
    horizons: Iterable[int] = (1, 5, 10, 20),
) -> dict[str, Any]:
    """Build a compact diagnostics report for signal governance gates."""

    score_map = {str(k): float(v) for k, v in scores.items()}
    weight_map = {str(k): float(v) for k, v in weights.items()}
    clean_horizons = tuple(sorted({int(h) for h in horizons if int(h) > 0}))

    ic_by_horizon: dict[str, float] = {}
    rank_ic_by_horizon: dict[str, float] = {}
    rank_autocorr_by_horizon: dict[str, float] = {}
    turnover_by_horizon: dict[str, float] = {}

    for horizon in clean_horizons:
        paired_scores: list[float] = []
        realized: list[float] = []
        lagged_proxy: list[float] = []
        for ticker, score in score_map.items():
            closes = _extract_closes(prices_by_ticker.get(ticker, []))
            if len(closes) <= horizon + 1:
                continue
            current = closes[-1]
            prior = closes[-(horizon + 1)]
            lag_current = closes[-2]
            lag_prior = closes[-(horizon + 2)]
            if prior <= 0 or lag_prior <= 0:
                continue
            paired_scores.append(float(score))
            realized.append(float((current / prior) - 1.0))
            lagged_proxy.append(float((lag_current / lag_prior) - 1.0))

        ic_by_horizon[str(horizon)] = _corr(paired_scores, realized)
        rank_ic_by_horizon[str(horizon)] = _corr(_ranks(paired_scores), _ranks(realized))
        rank_autocorr_by_horizon[str(horizon)] = _corr(_ranks(paired_scores), _ranks(lagged_proxy))
        turnover_by_horizon[str(horizon)] = 1.0 - _top_bucket_overlap(paired_scores, lagged_proxy)

    abs_weights = np.asarray([abs(weight_map.get(ticker, 0.0)) for ticker in score_map], dtype=float)
    gross = float(abs_weights.sum())
    squared = float(np.sum(abs_weights**2))
    hhi = float(squared / (gross * gross)) if gross > 0 else 0.0
    effective_n = float(1.0 / hhi) if hhi > 0 else 0.0
    max_abs_weight = float(np.max(abs_weights)) if abs_weights.size else 0.0

    liquidity_penalties: list[float] = []
    aligned_weights: list[float] = []
    aligned_scores: list[float] = []
    for ticker, score in score_map.items():
        closes = _extract_closes(prices_by_ticker.get(ticker, []))
        vols = _extract_volumes(prices_by_ticker.get(ticker, []))
        if not closes or not vols:
            continue
        usable = min(len(closes), len(vols), 20)
        if usable <= 0:
            continue
        dv = np.asarray(closes[-usable:], dtype=float) * np.asarray(vols[-usable:], dtype=float)
        liq = float(np.mean(np.log1p(np.clip(dv, 0.0, None))))
        if liq <= 0:
            continue
        w = abs(weight_map.get(ticker, 0.0))
        liquidity_penalties.append(float(w / liq))
        aligned_weights.append(float(w))
        aligned_scores.append(float(abs(score)))

    weighted_illiquidity = float(np.sum(liquidity_penalties)) if liquidity_penalties else 0.0
    crowding_corr = _corr(aligned_weights, aligned_scores)

    horizons_arr = np.asarray(clean_horizons, dtype=float)
    rank_ic_arr = np.asarray([rank_ic_by_horizon.get(str(h), 0.0) for h in clean_horizons], dtype=float)
    decay_slope = 0.0
    if horizons_arr.size >= 2 and float(np.std(horizons_arr)) > 0:
        decay_slope = float(np.polyfit(horizons_arr, rank_ic_arr, 1)[0])

    diagnostics = {
        "information_coefficient": {
            "ic": ic_by_horizon,
            "rank_ic": rank_ic_by_horizon,
            "ic_decay_slope": decay_slope,
        },
        "rank_stability": {
            "rank_autocorrelation": rank_autocorr_by_horizon,
            "feature_turnover": turnover_by_horizon,
            "avg_rank_autocorrelation": float(np.mean(list(rank_autocorr_by_horizon.values()))) if rank_autocorr_by_horizon else 0.0,
        },
        "exposure": {
            "gross_exposure": gross,
            "hhi_concentration": hhi,
            "effective_breadth": effective_n,
            "max_abs_weight": max_abs_weight,
            "crowding_proxy_weighted_illiquidity": weighted_illiquidity,
            "crowding_proxy_weight_score_corr": crowding_corr,
        },
    }
    diagnostics["diagnostics_ready"] = validate_signal_diagnostics(diagnostics)
    return diagnostics


def validate_signal_diagnostics(diagnostics: Mapping[str, Any] | None) -> bool:
    if not isinstance(diagnostics, Mapping):
        return False
    return all(isinstance(diagnostics.get(section), Mapping) and diagnostics.get(section) for section in MANDATORY_DIAGNOSTIC_SECTIONS)


def _extract_closes(series: list[float] | list[dict[str, float]] | tuple[float, ...] | Any) -> list[float]:
    values = list(series) if isinstance(series, (list, tuple)) else []
    closes: list[float] = []
    for row in values:
        if isinstance(row, dict):
            close = row.get("close")
            if close is not None:
                closes.append(float(close))
        else:
            closes.append(float(row))
    return closes


def _extract_volumes(series: list[float] | list[dict[str, float]] | tuple[float, ...] | Any) -> list[float]:
    values = list(series) if isinstance(series, (list, tuple)) else []
    volumes: list[float] = []
    for row in values:
        if isinstance(row, dict):
            volume = row.get("volume")
            if volume is not None:
                volumes.append(float(volume))
    return volumes


def _corr(x: Iterable[float], y: Iterable[float]) -> float:
    a = np.asarray(list(x), dtype=float)
    b = np.asarray(list(y), dtype=float)
    if a.size < 2 or b.size < 2 or a.size != b.size:
        return 0.0
    if float(np.std(a)) <= 0 or float(np.std(b)) <= 0:
        return 0.0
    return float(np.corrcoef(a, b)[0, 1])


def _ranks(values: Iterable[float]) -> np.ndarray:
    arr = np.asarray(list(values), dtype=float)
    if arr.size == 0:
        return arr
    order = np.argsort(arr)
    ranks = np.empty_like(order, dtype=float)
    ranks[order] = np.arange(1, arr.size + 1, dtype=float)
    return ranks


def _top_bucket_overlap(a_values: Iterable[float], b_values: Iterable[float], quantile: float = 0.2) -> float:
    a = np.asarray(list(a_values), dtype=float)
    b = np.asarray(list(b_values), dtype=float)
    if a.size == 0 or b.size == 0 or a.size != b.size:
        return 0.0
    top_n = max(1, int(np.ceil(a.size * float(quantile))))
    a_idx = set(np.argsort(a)[-top_n:].tolist())
    b_idx = set(np.argsort(b)[-top_n:].tolist())
    return float(len(a_idx & b_idx) / top_n)

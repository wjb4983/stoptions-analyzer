from __future__ import annotations

from dataclasses import dataclass
import math
from typing import Iterable

import numpy as np

from .base import CrossSectionalResult


@dataclass(frozen=True)
class MomentumSettings:
    lookback_days: int = 90
    skip_days: int = 5
    top_quantile: float = 0.2
    bottom_quantile: float = 0.2
    use_volatility_scaling: bool = False
    use_residual: bool = False
    use_multi_horizon: bool = False


def compute_cross_sectional_momentum(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    fundamentals_by_ticker: dict[str, dict] | None,
    settings: MomentumSettings,
) -> CrossSectionalResult:
    min_points = settings.lookback_days + settings.skip_days + 1
    returns: list[float] = []
    tickers: list[str] = []
    skipped: dict[str, str] = {}
    metrics: dict[str, dict[str, float]] = {}

    for ticker, prices in prices_by_ticker.items():
        series = list(prices)
        closes, volumes = _extract_close_volume(series)
        if len(closes) < min_points:
            skipped[ticker] = "insufficient_history"
            continue
        end_index = len(closes) - settings.skip_days - 1
        start_index = end_index - settings.lookback_days
        if start_index < 0:
            skipped[ticker] = "insufficient_history"
            continue
        start_price = closes[start_index]
        end_price = closes[end_index]
        if not _is_valid_price(start_price) or not _is_valid_price(end_price):
            skipped[ticker] = "invalid_price"
            continue
        momentum_return = (end_price / start_price) - 1.0
        ticker_metrics = {"base": float(momentum_return)}
        if settings.use_volatility_scaling:
            vol = _estimate_volatility(closes[start_index : end_index + 1])
            if vol and vol > 0:
                ticker_metrics["vol_scaled"] = float(momentum_return / vol)
        if settings.use_multi_horizon:
            blended = _multi_horizon_return(closes, end_index)
            if blended is not None:
                ticker_metrics["multi_horizon"] = float(blended)
        metrics[ticker] = ticker_metrics
        tickers.append(ticker)
        returns.append(float(momentum_return))

    if not returns:
        return CrossSectionalResult(
            scores={},
            ranking=[],
            longs=[],
            shorts=[],
            weights={},
            metrics={},
            metadata={
                "lookback_days": settings.lookback_days,
                "skip_days": settings.skip_days,
                "top_quantile": settings.top_quantile,
                "bottom_quantile": settings.bottom_quantile,
                "universe": 0,
            },
            skipped=skipped,
        )

    returns_array = np.asarray(returns, dtype=float)
    if settings.use_residual:
        residual = returns_array - float(np.mean(returns_array))
        for ticker, value in zip(tickers, residual):
            metrics.setdefault(ticker, {})["residual"] = float(value)

    combined_scores = []
    for ticker, base_score in zip(tickers, returns_array):
        selected_scores = []
        if settings.use_volatility_scaling:
            selected_scores.append(metrics.get(ticker, {}).get("vol_scaled"))
        if settings.use_residual:
            selected_scores.append(metrics.get(ticker, {}).get("residual"))
        if settings.use_multi_horizon:
            selected_scores.append(metrics.get(ticker, {}).get("multi_horizon"))
        selected_scores = [value for value in selected_scores if value is not None]
        if selected_scores:
            combined_scores.append(float(np.mean(selected_scores)))
            metrics.setdefault(ticker, {})["combined"] = combined_scores[-1]
        else:
            combined_scores.append(float(base_score))
    score_array = np.asarray(combined_scores, dtype=float)
    order = np.argsort(score_array)
    total = score_array.shape[0]
    top_n, bottom_n = _quantile_bucket_sizes(
        total, settings.top_quantile, settings.bottom_quantile
    )

    bottom_idx = order[:bottom_n]
    top_idx = order[-top_n:] if top_n > 0 else np.array([], dtype=int)

    ranking = [(tickers[idx], float(score_array[idx])) for idx in order[::-1]]
    longs = [tickers[idx] for idx in top_idx[::-1]]
    shorts = [tickers[idx] for idx in bottom_idx]
    if longs and shorts:
        long_set = set(longs)
        shorts = [ticker for ticker in shorts if ticker not in long_set]

    weights: dict[str, float] = {}
    if longs:
        long_weight = 1.0 / len(longs)
        weights.update({ticker: long_weight for ticker in longs})
    if shorts:
        short_weight = -1.0 / len(shorts)
        weights.update({ticker: short_weight for ticker in shorts})

    scores = {ticker: float(score) for ticker, score in zip(tickers, score_array)}

    return CrossSectionalResult(
        scores=scores,
        ranking=ranking,
        longs=longs,
        shorts=shorts,
        weights=weights,
        metadata={
            "lookback_days": settings.lookback_days,
            "skip_days": settings.skip_days,
            "top_quantile": settings.top_quantile,
            "bottom_quantile": settings.bottom_quantile,
            "universe": total,
            "use_volatility_scaling": settings.use_volatility_scaling,
            "use_residual": settings.use_residual,
            "use_multi_horizon": settings.use_multi_horizon,
        },
        metrics=metrics,
        skipped=skipped,
    )


def _quantile_bucket_sizes(total: int, top_quantile: float, bottom_quantile: float) -> tuple[int, int]:
    if total <= 0:
        return 0, 0
    top = max(1, int(math.ceil(total * top_quantile))) if top_quantile > 0 else 0
    bottom = max(1, int(math.ceil(total * bottom_quantile))) if bottom_quantile > 0 else 0
    if top + bottom > total:
        bottom = max(0, total - top)
    return top, bottom


def _is_valid_price(value: object) -> bool:
    return isinstance(value, (int, float)) and value > 0


def _extract_close_volume(series: list[object]) -> tuple[list[float], list[float]]:
    closes: list[float] = []
    volumes: list[float] = []
    for item in series:
        if isinstance(item, dict):
            close_value = item.get("close")
            volume_value = item.get("volume")
        else:
            close_value = item
            volume_value = None
        if isinstance(close_value, (int, float)):
            closes.append(float(close_value))
            if isinstance(volume_value, (int, float)):
                volumes.append(float(volume_value))
    return closes, volumes


def _estimate_volatility(closes: list[float]) -> float | None:
    if len(closes) < 2:
        return None
    returns = np.diff(np.log(np.asarray(closes, dtype=float)))
    if returns.size == 0:
        return None
    return float(np.std(returns, ddof=1))


def _multi_horizon_return(closes: list[float], end_index: int) -> float | None:
    windows = [20, 60, 120]
    values: list[float] = []
    for window in windows:
        start_index = end_index - window
        if start_index < 0:
            continue
        start_price = closes[start_index]
        end_price = closes[end_index]
        if not _is_valid_price(start_price) or not _is_valid_price(end_price):
            continue
        values.append((end_price / start_price) - 1.0)
    if not values:
        return None
    return float(np.mean(values))

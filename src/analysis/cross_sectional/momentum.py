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


def compute_cross_sectional_momentum(
    prices_by_ticker: dict[str, list[float] | tuple[float, ...]],
    settings: MomentumSettings,
) -> CrossSectionalResult:
    min_points = settings.lookback_days + settings.skip_days + 1
    returns: list[float] = []
    tickers: list[str] = []
    skipped: dict[str, str] = {}

    for ticker, prices in prices_by_ticker.items():
        series = list(prices)
        if len(series) < min_points:
            skipped[ticker] = "insufficient_history"
            continue
        end_index = len(series) - settings.skip_days - 1
        start_index = end_index - settings.lookback_days
        if start_index < 0:
            skipped[ticker] = "insufficient_history"
            continue
        start_price = series[start_index]
        end_price = series[end_index]
        if not _is_valid_price(start_price) or not _is_valid_price(end_price):
            skipped[ticker] = "invalid_price"
            continue
        momentum_return = (end_price / start_price) - 1.0
        tickers.append(ticker)
        returns.append(float(momentum_return))

    if not returns:
        return CrossSectionalResult(
            scores={},
            ranking=[],
            longs=[],
            shorts=[],
            weights={},
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
    order = np.argsort(returns_array)
    total = returns_array.shape[0]
    top_n, bottom_n = _quantile_bucket_sizes(
        total, settings.top_quantile, settings.bottom_quantile
    )

    bottom_idx = order[:bottom_n]
    top_idx = order[-top_n:] if top_n > 0 else np.array([], dtype=int)

    ranking = [(tickers[idx], float(returns_array[idx])) for idx in order[::-1]]
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

    scores = {ticker: float(score) for ticker, score in zip(tickers, returns)}

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
        },
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

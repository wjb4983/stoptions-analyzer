from __future__ import annotations

from dataclasses import dataclass
from typing import Callable

import numpy as np

from .base import CrossSectionalResult
from .momentum import _extract_close_volume


@dataclass(frozen=True)
class CrossSectionalSettings:
    top_quantile: float = 0.2
    bottom_quantile: float = 0.2


def compute_cross_sectional_low_volatility(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    settings: CrossSectionalSettings,
) -> CrossSectionalResult:
    scores: dict[str, float] = {}
    skipped: dict[str, str] = {}
    for ticker, series in prices_by_ticker.items():
        closes, _volumes = _extract_close_volume(list(series))
        if len(closes) < 2:
            skipped[ticker] = "insufficient_history"
            continue
        returns = np.diff(np.log(np.asarray(closes, dtype=float)))
        if returns.size == 0:
            skipped[ticker] = "insufficient_history"
            continue
        vol = float(np.std(returns, ddof=1))
        if vol <= 0:
            skipped[ticker] = "invalid_price"
            continue
        scores[ticker] = -vol
    return _rank_scores(scores, skipped, settings, metadata={"strategy": "low_volatility"})


def compute_cross_sectional_liquidity(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    settings: CrossSectionalSettings,
) -> CrossSectionalResult:
    scores: dict[str, float] = {}
    skipped: dict[str, str] = {}
    for ticker, series in prices_by_ticker.items():
        closes, volumes = _extract_close_volume(list(series))
        if not closes or not volumes:
            skipped[ticker] = "missing_volume"
            continue
        usable = min(len(closes), len(volumes))
        dollar_volume = np.asarray(closes[:usable]) * np.asarray(volumes[:usable])
        if dollar_volume.size == 0:
            skipped[ticker] = "missing_volume"
            continue
        scores[ticker] = float(np.mean(dollar_volume))
    return _rank_scores(scores, skipped, settings, metadata={"strategy": "liquidity"})


def compute_cross_sectional_value(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    settings: CrossSectionalSettings,
) -> CrossSectionalResult:
    return _missing_data_result(prices_by_ticker, settings, "missing_fundamentals", "value")


def compute_cross_sectional_size(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    settings: CrossSectionalSettings,
) -> CrossSectionalResult:
    return _missing_data_result(prices_by_ticker, settings, "missing_market_cap", "size")


def compute_cross_sectional_quality(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    settings: CrossSectionalSettings,
) -> CrossSectionalResult:
    return _missing_data_result(prices_by_ticker, settings, "missing_fundamentals", "quality")


def compute_cross_sectional_investment(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    settings: CrossSectionalSettings,
) -> CrossSectionalResult:
    return _missing_data_result(prices_by_ticker, settings, "missing_fundamentals", "investment")


def compute_cross_sectional_earnings_momentum(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    settings: CrossSectionalSettings,
) -> CrossSectionalResult:
    return _missing_data_result(prices_by_ticker, settings, "missing_earnings_data", "earnings_momentum")


def compute_cross_sectional_carry(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    settings: CrossSectionalSettings,
) -> CrossSectionalResult:
    return _missing_data_result(prices_by_ticker, settings, "missing_fundamentals", "carry")


def _missing_data_result(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    settings: CrossSectionalSettings,
    reason: str,
    strategy: str,
) -> CrossSectionalResult:
    skipped = {ticker: reason for ticker in prices_by_ticker}
    return _rank_scores({}, skipped, settings, metadata={"strategy": strategy})


def _rank_scores(
    scores: dict[str, float],
    skipped: dict[str, str],
    settings: CrossSectionalSettings,
    metadata: dict[str, object] | None = None,
) -> CrossSectionalResult:
    if not scores:
        return CrossSectionalResult(
            scores={},
            ranking=[],
            longs=[],
            shorts=[],
            weights={},
            metadata={**(metadata or {}), "top_quantile": settings.top_quantile, "bottom_quantile": settings.bottom_quantile},
            skipped=skipped,
        )
    tickers = list(scores.keys())
    values = np.asarray([scores[ticker] for ticker in tickers], dtype=float)
    order = np.argsort(values)
    total = values.shape[0]
    top_n = max(1, int(np.ceil(total * settings.top_quantile))) if settings.top_quantile > 0 else 0
    bottom_n = max(1, int(np.ceil(total * settings.bottom_quantile))) if settings.bottom_quantile > 0 else 0
    if top_n + bottom_n > total:
        bottom_n = max(0, total - top_n)
    bottom_idx = order[:bottom_n]
    top_idx = order[-top_n:] if top_n > 0 else np.array([], dtype=int)
    ranking = [(tickers[idx], float(values[idx])) for idx in order[::-1]]
    longs = [tickers[idx] for idx in top_idx[::-1]]
    shorts = [tickers[idx] for idx in bottom_idx]
    weights: dict[str, float] = {}
    if longs:
        weights.update({ticker: 1.0 / len(longs) for ticker in longs})
    if shorts:
        weights.update({ticker: -1.0 / len(shorts) for ticker in shorts})
    return CrossSectionalResult(
        scores=scores,
        ranking=ranking,
        longs=longs,
        shorts=shorts,
        weights=weights,
        metadata={**(metadata or {}), "top_quantile": settings.top_quantile, "bottom_quantile": settings.bottom_quantile},
        skipped=skipped,
    )

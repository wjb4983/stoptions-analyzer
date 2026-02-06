from __future__ import annotations

from dataclasses import dataclass
from typing import Callable

import numpy as np

from ..common import extract_close_volume
from .base import CrossSectionalResult


@dataclass(frozen=True)
class CrossSectionalSettings:
    top_quantile: float = 0.2
    bottom_quantile: float = 0.2


@dataclass(frozen=True)
class StrategySpec:
    name: str
    compute: Callable[
        [dict[str, list[float] | list[dict] | tuple[float, ...]], dict[str, dict] | None, CrossSectionalSettings],
        CrossSectionalResult,
    ]
    required_data: list[str]
    description: str


def compute_cross_sectional_low_volatility(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    fundamentals_by_ticker: dict[str, dict] | None,
    settings: CrossSectionalSettings,
) -> CrossSectionalResult:
    scores: dict[str, float] = {}
    skipped: dict[str, str] = {}
    for ticker, series in prices_by_ticker.items():
        closes, _volumes = extract_close_volume(list(series))
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
    fundamentals_by_ticker: dict[str, dict] | None,
    settings: CrossSectionalSettings,
) -> CrossSectionalResult:
    scores: dict[str, float] = {}
    skipped: dict[str, str] = {}
    for ticker, series in prices_by_ticker.items():
        closes, volumes = extract_close_volume(list(series))
        if not closes or not volumes:
            skipped[ticker] = "missing_volume"
            continue
        usable = min(len(closes), len(volumes))
        dollar_volume = np.asarray(closes[:usable]) * np.asarray(volumes[:usable])
        if dollar_volume.size == 0:
            skipped[ticker] = "missing_volume"
            continue
        scores[ticker] = float(np.mean(np.log1p(dollar_volume)))
    return _rank_scores(scores, skipped, settings, metadata={"strategy": "liquidity"})


def compute_cross_sectional_value(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    fundamentals_by_ticker: dict[str, dict] | None,
    settings: CrossSectionalSettings,
) -> CrossSectionalResult:
    return _missing_data_result(prices_by_ticker, settings, "missing_fundamentals", "value")


def compute_cross_sectional_size(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    fundamentals_by_ticker: dict[str, dict] | None,
    settings: CrossSectionalSettings,
) -> CrossSectionalResult:
    scores: dict[str, float] = {}
    skipped: dict[str, str] = {}
    for ticker, series in prices_by_ticker.items():
        fundamentals = (fundamentals_by_ticker or {}).get(ticker, {})
        market_cap = _coerce_number(fundamentals.get("market_cap"))
        if market_cap is None:
            shares_outstanding = _coerce_number(fundamentals.get("shares_outstanding"))
            last_close = _latest_close(series)
            if shares_outstanding is None or last_close is None:
                skipped[ticker] = "missing_market_cap"
                continue
            market_cap = shares_outstanding * last_close
        if market_cap <= 0:
            skipped[ticker] = "missing_market_cap"
            continue
        scores[ticker] = -float(np.log(market_cap))
    return _rank_scores(scores, skipped, settings, metadata={"strategy": "size"})


def compute_cross_sectional_quality(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    fundamentals_by_ticker: dict[str, dict] | None,
    settings: CrossSectionalSettings,
) -> CrossSectionalResult:
    return _missing_data_result(prices_by_ticker, settings, "missing_fundamentals", "quality")


def compute_cross_sectional_investment(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    fundamentals_by_ticker: dict[str, dict] | None,
    settings: CrossSectionalSettings,
) -> CrossSectionalResult:
    return _missing_data_result(prices_by_ticker, settings, "missing_fundamentals", "investment")


def compute_cross_sectional_earnings_momentum(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    fundamentals_by_ticker: dict[str, dict] | None,
    settings: CrossSectionalSettings,
) -> CrossSectionalResult:
    return _missing_data_result(prices_by_ticker, settings, "missing_earnings_data", "earnings_momentum")


def compute_cross_sectional_carry(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    fundamentals_by_ticker: dict[str, dict] | None,
    settings: CrossSectionalSettings,
) -> CrossSectionalResult:
    return _missing_data_result(prices_by_ticker, settings, "missing_fundamentals", "carry")


STRATEGY_REGISTRY = {
    "Value": StrategySpec(
        name="Value",
        compute=compute_cross_sectional_value,
        required_data=["fundamentals"],
        description="Ranks on valuation metrics like book-to-market or earnings yield.",
    ),
    "Size": StrategySpec(
        name="Size",
        compute=compute_cross_sectional_size,
        required_data=["market_cap"],
        description="Ranks on market capitalization (small vs large).",
    ),
    "Quality": StrategySpec(
        name="Quality",
        compute=compute_cross_sectional_quality,
        required_data=["fundamentals"],
        description="Ranks on profitability or quality metrics (e.g., ROE).",
    ),
    "Investment": StrategySpec(
        name="Investment",
        compute=compute_cross_sectional_investment,
        required_data=["fundamentals"],
        description="Ranks on asset growth or investment rates.",
    ),
    "Low Volatility": StrategySpec(
        name="Low Volatility",
        compute=compute_cross_sectional_low_volatility,
        required_data=["prices"],
        description="Ranks by trailing volatility; long low-vol names.",
    ),
    "Liquidity": StrategySpec(
        name="Liquidity",
        compute=compute_cross_sectional_liquidity,
        required_data=["prices", "volume"],
        description="Ranks by average dollar volume / liquidity.",
    ),
    "Earnings Momentum": StrategySpec(
        name="Earnings Momentum",
        compute=compute_cross_sectional_earnings_momentum,
        required_data=["earnings", "analyst_revisions"],
        description="Ranks on earnings surprises or analyst revisions.",
    ),
    "Carry / Yield": StrategySpec(
        name="Carry / Yield",
        compute=compute_cross_sectional_carry,
        required_data=["fundamentals"],
        description="Ranks on dividend yield or other carry signals.",
    ),
}


def _missing_data_result(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    settings: CrossSectionalSettings,
    reason: str,
    strategy: str,
) -> CrossSectionalResult:
    skipped = {ticker: reason for ticker in prices_by_ticker}
    return _rank_scores({}, skipped, settings, metadata={"strategy": strategy})


def _latest_close(series: list[float] | list[dict] | tuple[float, ...]) -> float | None:
    values = list(series)
    closes, _volumes = extract_close_volume(values)
    if not closes:
        return None
    return float(closes[-1])


def _coerce_number(value: object) -> float | None:
    if isinstance(value, (int, float)):
        return float(value)
    return None


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

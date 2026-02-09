from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Iterable

import numpy as np

from ..common import estimate_volatility, extract_close_volume, quantile_bucket_sizes
from .base import TimeSeriesResult
from utils.parsing import _coerce_number


@dataclass(frozen=True)
class TimeSeriesSettings:
    lookback_days: int = 90
    skip_days: int = 0
    top_quantile: float = 0.2
    bottom_quantile: float = 0.2
    use_volatility_scaling: bool = False
    use_residual: bool = False
    use_multi_horizon: bool = False
    multi_horizon_windows: tuple[int, ...] = (20, 60, 120)
    use_zscore: bool = False
    winsorize_sigma: float | None = None


@dataclass(frozen=True)
class TimeSeriesStrategySpec:
    name: str
    compute: Callable[
        [dict[str, list[float] | list[dict] | tuple[float, ...]], dict[str, dict] | None, TimeSeriesSettings],
        TimeSeriesResult,
    ]
    required_data: list[str]
    description: str


def compute_time_series_low_volatility(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    fundamentals_by_ticker: dict[str, dict] | None,
    settings: TimeSeriesSettings,
) -> TimeSeriesResult:
    scores: dict[str, float] = {}
    skipped: dict[str, str] = {}
    metrics: dict[str, dict[str, float]] = {}
    for ticker, series in prices_by_ticker.items():
        closes, _volumes = extract_close_volume(list(series))
        windowed = _windowed_series(closes, settings.lookback_days)
        if len(windowed) < 2:
            skipped[ticker] = "insufficient_history"
            continue
        vol = estimate_volatility(windowed)
        if vol is None or vol <= 0:
            skipped[ticker] = "invalid_price"
            continue
        score = -float(vol)
        metrics[ticker] = {"base": score}
        scores[ticker] = score
        if settings.use_multi_horizon:
            multi = _multi_window_stat(closes, settings.multi_horizon_windows, _volatility_score)
            if multi is not None:
                metrics[ticker]["multi_horizon"] = float(multi)
    return _rank_scores(scores, skipped, settings, metrics, prices_by_ticker, metadata={"strategy": "low_volatility"})


def compute_time_series_liquidity(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    fundamentals_by_ticker: dict[str, dict] | None,
    settings: TimeSeriesSettings,
) -> TimeSeriesResult:
    scores: dict[str, float] = {}
    skipped: dict[str, str] = {}
    metrics: dict[str, dict[str, float]] = {}
    for ticker, series in prices_by_ticker.items():
        closes, volumes = extract_close_volume(list(series))
        if not closes or not volumes:
            skipped[ticker] = "missing_volume"
            continue
        usable = min(len(closes), len(volumes))
        trimmed_closes = closes[:usable]
        trimmed_volumes = volumes[:usable]
        if settings.lookback_days > 0:
            trimmed_closes = trimmed_closes[-settings.lookback_days :]
            trimmed_volumes = trimmed_volumes[-settings.lookback_days :]
        dollar_volume = np.asarray(trimmed_closes) * np.asarray(trimmed_volumes)
        if dollar_volume.size == 0:
            skipped[ticker] = "missing_volume"
            continue
        score = float(np.mean(np.log1p(dollar_volume)))
        scores[ticker] = score
        metrics[ticker] = {"base": score}
        if settings.use_multi_horizon:
            multi = _multi_window_stat(list(zip(closes, volumes)), settings.multi_horizon_windows, _liquidity_score)
            if multi is not None:
                metrics[ticker]["multi_horizon"] = float(multi)
    return _rank_scores(scores, skipped, settings, metrics, prices_by_ticker, metadata={"strategy": "liquidity"})


def compute_time_series_value(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    fundamentals_by_ticker: dict[str, dict] | None,
    settings: TimeSeriesSettings,
) -> TimeSeriesResult:
    return _fundamental_ratio_strategy(
        prices_by_ticker,
        fundamentals_by_ticker,
        settings,
        ratio_key="book_value",
        missing_reason="missing_fundamentals",
        strategy="value",
    )


def compute_time_series_size(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    fundamentals_by_ticker: dict[str, dict] | None,
    settings: TimeSeriesSettings,
) -> TimeSeriesResult:
    scores: dict[str, float] = {}
    skipped: dict[str, str] = {}
    metrics: dict[str, dict[str, float]] = {}
    fundamentals_by_ticker = fundamentals_by_ticker or {}
    for ticker, series in prices_by_ticker.items():
        fundamentals = fundamentals_by_ticker.get(ticker, {})
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
        score = -float(np.log(market_cap))
        scores[ticker] = score
        metrics[ticker] = {"base": score}
    return _rank_scores(scores, skipped, settings, metrics, prices_by_ticker, metadata={"strategy": "size"})


def compute_time_series_quality(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    fundamentals_by_ticker: dict[str, dict] | None,
    settings: TimeSeriesSettings,
) -> TimeSeriesResult:
    scores: dict[str, float] = {}
    skipped: dict[str, str] = {}
    metrics: dict[str, dict[str, float]] = {}
    fundamentals_by_ticker = fundamentals_by_ticker or {}
    for ticker, series in prices_by_ticker.items():
        fundamentals = fundamentals_by_ticker.get(ticker, {})
        earnings = _coerce_number(fundamentals.get("earnings_actual"))
        book_value = _coerce_number(fundamentals.get("book_value"))
        market_cap = _coerce_number(fundamentals.get("market_cap"))
        last_close = _latest_close(series)
        denominator = book_value or market_cap or last_close
        if earnings is None or denominator in (None, 0):
            skipped[ticker] = "missing_quality_data"
            continue
        score = float(earnings / denominator)
        scores[ticker] = score
        metrics[ticker] = {"base": score}
    return _rank_scores(scores, skipped, settings, metrics, prices_by_ticker, metadata={"strategy": "quality"})


def compute_time_series_investment(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    fundamentals_by_ticker: dict[str, dict] | None,
    settings: TimeSeriesSettings,
) -> TimeSeriesResult:
    scores: dict[str, float] = {}
    skipped: dict[str, str] = {}
    metrics: dict[str, dict[str, float]] = {}
    fundamentals_by_ticker = fundamentals_by_ticker or {}
    for ticker in prices_by_ticker:
        fundamentals = fundamentals_by_ticker.get(ticker, {})
        total_assets = _coerce_number(fundamentals.get("total_assets"))
        if total_assets is None:
            skipped[ticker] = "missing_investment_data"
            continue
        score = -float(np.log(total_assets))
        scores[ticker] = score
        metrics[ticker] = {"base": score}
    return _rank_scores(scores, skipped, settings, metrics, prices_by_ticker, metadata={"strategy": "investment"})


def compute_time_series_earnings_momentum(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    fundamentals_by_ticker: dict[str, dict] | None,
    settings: TimeSeriesSettings,
) -> TimeSeriesResult:
    scores: dict[str, float] = {}
    skipped: dict[str, str] = {}
    metrics: dict[str, dict[str, float]] = {}
    fundamentals_by_ticker = fundamentals_by_ticker or {}
    for ticker in prices_by_ticker:
        fundamentals = fundamentals_by_ticker.get(ticker, {})
        surprise = _coerce_number(fundamentals.get("earnings_surprise"))
        if surprise is None:
            actual = _coerce_number(fundamentals.get("earnings_actual"))
            estimate = _coerce_number(fundamentals.get("earnings_estimate"))
            if actual is None or estimate is None:
                skipped[ticker] = "missing_earnings_data"
                continue
            surprise = actual - estimate
        score = float(surprise)
        scores[ticker] = score
        metrics[ticker] = {"base": score}
    return _rank_scores(scores, skipped, settings, metrics, prices_by_ticker, metadata={"strategy": "earnings_momentum"})


def compute_time_series_carry(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    fundamentals_by_ticker: dict[str, dict] | None,
    settings: TimeSeriesSettings,
) -> TimeSeriesResult:
    scores: dict[str, float] = {}
    skipped: dict[str, str] = {}
    metrics: dict[str, dict[str, float]] = {}
    fundamentals_by_ticker = fundamentals_by_ticker or {}
    for ticker, series in prices_by_ticker.items():
        fundamentals = fundamentals_by_ticker.get(ticker, {})
        dividends = _coerce_number(fundamentals.get("dividends_ttm"))
        if dividends is None:
            skipped[ticker] = "missing_dividends"
            continue
        last_close = _latest_close(series)
        market_cap = _coerce_number(fundamentals.get("market_cap"))
        denominator = market_cap or last_close
        if denominator in (None, 0):
            skipped[ticker] = "missing_market_cap"
            continue
        score = float(dividends / denominator)
        scores[ticker] = score
        metrics[ticker] = {"base": score}
    return _rank_scores(scores, skipped, settings, metrics, prices_by_ticker, metadata={"strategy": "carry"})


TIME_SERIES_STRATEGY_REGISTRY = {
    "Value": TimeSeriesStrategySpec(
        name="Value",
        compute=compute_time_series_value,
        required_data=["fundamentals"],
        description="Ranks on valuation metrics like book-to-market or earnings yield.",
    ),
    "Size": TimeSeriesStrategySpec(
        name="Size",
        compute=compute_time_series_size,
        required_data=["market_cap"],
        description="Ranks on market capitalization (small vs large).",
    ),
    "Quality": TimeSeriesStrategySpec(
        name="Quality",
        compute=compute_time_series_quality,
        required_data=["fundamentals"],
        description="Ranks on profitability or quality metrics (e.g., ROE).",
    ),
    "Investment": TimeSeriesStrategySpec(
        name="Investment",
        compute=compute_time_series_investment,
        required_data=["fundamentals"],
        description="Ranks on asset growth or investment rates.",
    ),
    "Low Volatility": TimeSeriesStrategySpec(
        name="Low Volatility",
        compute=compute_time_series_low_volatility,
        required_data=["prices"],
        description="Ranks by trailing volatility; long low-vol names.",
    ),
    "Liquidity": TimeSeriesStrategySpec(
        name="Liquidity",
        compute=compute_time_series_liquidity,
        required_data=["prices", "volume"],
        description="Ranks by average dollar volume / liquidity.",
    ),
    "Earnings Momentum": TimeSeriesStrategySpec(
        name="Earnings Momentum",
        compute=compute_time_series_earnings_momentum,
        required_data=["earnings", "analyst_revisions"],
        description="Ranks on earnings surprises or analyst revisions.",
    ),
    "Carry / Yield": TimeSeriesStrategySpec(
        name="Carry / Yield",
        compute=compute_time_series_carry,
        required_data=["fundamentals"],
        description="Ranks on dividend yield or other carry signals.",
    ),
}


def _fundamental_ratio_strategy(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    fundamentals_by_ticker: dict[str, dict] | None,
    settings: TimeSeriesSettings,
    *,
    ratio_key: str,
    missing_reason: str,
    strategy: str,
) -> TimeSeriesResult:
    scores: dict[str, float] = {}
    skipped: dict[str, str] = {}
    metrics: dict[str, dict[str, float]] = {}
    fundamentals_by_ticker = fundamentals_by_ticker or {}
    for ticker, series in prices_by_ticker.items():
        fundamentals = fundamentals_by_ticker.get(ticker, {})
        numerator = _coerce_number(fundamentals.get(ratio_key))
        market_cap = _coerce_number(fundamentals.get("market_cap"))
        if market_cap is None:
            shares_outstanding = _coerce_number(fundamentals.get("shares_outstanding"))
            last_close = _latest_close(series)
            if shares_outstanding is None or last_close is None:
                skipped[ticker] = missing_reason
                continue
            market_cap = shares_outstanding * last_close
        if numerator is None or market_cap <= 0:
            skipped[ticker] = missing_reason
            continue
        score = float(numerator / market_cap)
        scores[ticker] = score
        metrics[ticker] = {"base": score}
    return _rank_scores(scores, skipped, settings, metrics, prices_by_ticker, metadata={"strategy": strategy})


def _latest_close(series: list[float] | list[dict] | tuple[float, ...]) -> float | None:
    values = list(series)
    closes, _volumes = extract_close_volume(values)
    if not closes:
        return None
    return float(closes[-1])


def _windowed_series(values: list[float], lookback_days: int) -> list[float]:
    if lookback_days <= 0:
        return list(values)
    return list(values[-lookback_days:])


def _multi_window_stat(
    series: Iterable,
    windows: Iterable[int],
    stat_func: Callable[[Iterable], float | None],
) -> float | None:
    values: list[float] = []
    series_list = list(series)
    for window in windows:
        if window <= 0:
            continue
        windowed = series_list[-window:]
        stat = stat_func(windowed)
        if stat is not None:
            values.append(float(stat))
    if not values:
        return None
    return float(np.mean(values))


def _volatility_score(closes: Iterable[float]) -> float | None:
    values = list(closes)
    if len(values) < 2:
        return None
    vol = estimate_volatility(values)
    if vol is None or vol <= 0:
        return None
    return -float(vol)


def _liquidity_score(closes_and_volumes: Iterable[tuple[float, float]]) -> float | None:
    values = list(closes_and_volumes)
    if not values:
        return None
    closes, volumes = zip(*values)
    dollar_volume = np.asarray(closes) * np.asarray(volumes)
    if dollar_volume.size == 0:
        return None
    return float(np.mean(np.log1p(dollar_volume)))


def _rank_scores(
    scores: dict[str, float],
    skipped: dict[str, str],
    settings: TimeSeriesSettings,
    metrics: dict[str, dict[str, float]],
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    metadata: dict[str, object] | None = None,
) -> TimeSeriesResult:
    if not scores:
        return TimeSeriesResult(
            scores={},
            ranking=[],
            longs=[],
            shorts=[],
            weights={},
            metadata={
                **(metadata or {}),
                "top_quantile": settings.top_quantile,
                "bottom_quantile": settings.bottom_quantile,
            },
            skipped=skipped,
        )
    tickers = list(scores.keys())
    values = np.asarray([scores[ticker] for ticker in tickers], dtype=float)
    values = _apply_adjustments(
        tickers,
        values,
        metrics,
        settings,
        prices_by_ticker,
    )
    order = np.argsort(values)
    total = values.shape[0]
    top_n, bottom_n = quantile_bucket_sizes(
        total, settings.top_quantile, settings.bottom_quantile
    )
    bottom_idx = order[:bottom_n]
    top_idx = order[-top_n:] if top_n > 0 else np.array([], dtype=int)
    ranking = [(tickers[idx], float(values[idx])) for idx in order[::-1]]
    longs = [tickers[idx] for idx in top_idx[::-1]]
    shorts = [tickers[idx] for idx in bottom_idx]
    if longs and shorts:
        long_set = set(longs)
        shorts = [ticker for ticker in shorts if ticker not in long_set]
    weights: dict[str, float] = {}
    if longs:
        weights.update({ticker: 1.0 / len(longs) for ticker in longs})
    if shorts:
        weights.update({ticker: -1.0 / len(shorts) for ticker in shorts})
    return TimeSeriesResult(
        scores={ticker: float(score) for ticker, score in zip(tickers, values)},
        ranking=ranking,
        longs=longs,
        shorts=shorts,
        weights=weights,
        metadata={
            **(metadata or {}),
            "top_quantile": settings.top_quantile,
            "bottom_quantile": settings.bottom_quantile,
            "use_volatility_scaling": settings.use_volatility_scaling,
            "use_residual": settings.use_residual,
            "use_multi_horizon": settings.use_multi_horizon,
            "use_zscore": settings.use_zscore,
            "winsorize_sigma": settings.winsorize_sigma,
        },
        skipped=skipped,
        metrics=metrics,
    )


def _apply_adjustments(
    tickers: list[str],
    scores: np.ndarray,
    metrics: dict[str, dict[str, float]],
    settings: TimeSeriesSettings,
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
) -> np.ndarray:
    adjusted = scores.astype(float)
    if settings.use_residual:
        mean = float(np.mean(adjusted))
        residual = adjusted - mean
        adjusted = residual
        for ticker, value in zip(tickers, residual):
            metrics.setdefault(ticker, {})["residual"] = float(value)
    if settings.use_multi_horizon:
        blended_scores = []
        for ticker, base_score in zip(tickers, adjusted):
            ticker_metrics = metrics.setdefault(ticker, {})
            multi = ticker_metrics.get("multi_horizon")
            if multi is None:
                ticker_metrics["multi_horizon"] = float(base_score)
                blended_scores.append(float(base_score))
            else:
                blended_scores.append(float(np.mean([base_score, multi])))
        adjusted = np.asarray(blended_scores, dtype=float)
    if settings.use_volatility_scaling:
        scaled_scores = []
        for ticker, base_score in zip(tickers, adjusted):
            closes, _volumes = extract_close_volume(list(prices_by_ticker.get(ticker, [])))
            windowed = _windowed_series(closes, settings.lookback_days)
            vol = estimate_volatility(windowed)
            if vol is None or vol <= 0:
                scaled_scores.append(float(base_score))
                metrics.setdefault(ticker, {})["vol_scaled"] = float(base_score)
            else:
                scaled = float(base_score / vol)
                scaled_scores.append(scaled)
                metrics.setdefault(ticker, {})["vol_scaled"] = scaled
        adjusted = np.asarray(scaled_scores, dtype=float)
    mean = float(np.mean(adjusted))
    std = float(np.std(adjusted))
    if settings.winsorize_sigma is not None and std > 0:
        lower = mean - settings.winsorize_sigma * std
        upper = mean + settings.winsorize_sigma * std
        adjusted = np.clip(adjusted, lower, upper)
    if settings.use_zscore and std > 0:
        adjusted = (adjusted - mean) / std
    return adjusted

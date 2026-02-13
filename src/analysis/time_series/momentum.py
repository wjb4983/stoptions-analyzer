from __future__ import annotations

from dataclasses import dataclass
import importlib.util
from typing import Literal

import numpy as np

from ..common import (
    estimate_volatility,
    extract_close_volume,
    is_valid_price,
    multi_horizon_return,
    quantile_bucket_sizes,
)
from ..diagnostics import compute_signal_diagnostics
from .base import TimeSeriesResult


_numba_spec = importlib.util.find_spec("numba")
if _numba_spec:
    from numba import njit
else:

    def njit(*args, **kwargs):
        def _decorator(func):
            return func

        return _decorator


ComputationMode = Literal["reference", "optimized"]


@dataclass(frozen=True)
class TimeSeriesMomentumSettings:
    lookback_days: int = 90
    skip_days: int = 5
    vol_window_days: int = 20
    target_volatility: float | None = None
    max_leverage: float = 1.0
    top_quantile: float = 0.2
    bottom_quantile: float = 0.2
    use_volatility_scaling: bool = False
    use_residual: bool = False
    use_multi_horizon: bool = False
    use_zscore: bool = False
    winsorize_sigma: float | None = None
    hyperparameters: "MomentumHyperparameters | None" = None


@dataclass(frozen=True)
class MomentumHyperparameters:
    lookback_days: int = 90
    skip_days: int = 5
    vol_window_days: int = 20
    target_volatility: float | None = None
    max_leverage: float = 1.0


@dataclass(frozen=True)
class TimeSeriesMomentumArrays:
    raw_score: np.ndarray
    position_signal: np.ndarray
    confidence: np.ndarray
    scaled_weight: np.ndarray
    tradable_position: np.ndarray


def _resolve_hyperparameters(settings: TimeSeriesMomentumSettings) -> MomentumHyperparameters:
    if settings.hyperparameters is not None:
        return settings.hyperparameters
    return MomentumHyperparameters(
        lookback_days=settings.lookback_days,
        skip_days=settings.skip_days,
        vol_window_days=settings.vol_window_days,
        target_volatility=settings.target_volatility,
        max_leverage=settings.max_leverage,
    )


def build_time_series_momentum_arrays(
    closes: list[float] | tuple[float, ...],
    settings: TimeSeriesMomentumSettings,
    mode: ComputationMode = "optimized",
) -> TimeSeriesMomentumArrays:
    """Build momentum arrays from close data with next-open execution alignment.

    The signal uses only data available through close ``t`` and a tradable
    position is produced by shifting one bar so that execution begins at
    ``t+1`` open.
    """

    params = _resolve_hyperparameters(settings)
    close_array = np.asarray(closes, dtype=float)
    count = close_array.shape[0]
    if mode == "reference":
        raw_score = _compute_raw_score_reference(close_array, params.lookback_days, params.skip_days)
    else:
        raw_score = _compute_raw_score_optimized(close_array, params.lookback_days, params.skip_days)

    position_signal = np.zeros(count, dtype=float)
    valid = np.isfinite(raw_score)
    position_signal[valid] = np.sign(raw_score[valid])
    confidence = np.zeros(count, dtype=float)
    confidence[valid] = np.abs(raw_score[valid])

    scaled_weight = position_signal.astype(float)
    if params.target_volatility is not None and params.target_volatility > 0:
        daily_returns = np.zeros(count, dtype=float)
        daily_returns[1:] = close_array[1:] / close_array[:-1] - 1.0
        if mode == "reference":
            scaled_weight = _apply_vol_target_reference(
                position_signal,
                daily_returns,
                skip_days=params.skip_days,
                vol_window_days=params.vol_window_days,
                target_volatility=float(params.target_volatility),
                max_leverage=float(params.max_leverage),
            )
        else:
            scaled_weight = _apply_vol_target_optimized(
                position_signal,
                daily_returns,
                skip_days=params.skip_days,
                vol_window_days=params.vol_window_days,
                target_volatility=float(params.target_volatility),
                max_leverage=float(params.max_leverage),
            )

    tradable_position = np.zeros(count, dtype=float)
    tradable_position[1:] = scaled_weight[:-1]

    return TimeSeriesMomentumArrays(
        raw_score=raw_score,
        position_signal=position_signal,
        confidence=confidence,
        scaled_weight=scaled_weight,
        tradable_position=tradable_position,
    )


def _compute_raw_score_reference(closes: np.ndarray, lookback_days: int, skip_days: int) -> np.ndarray:
    count = closes.shape[0]
    raw_score = np.full(count, np.nan, dtype=float)
    for idx in range(count):
        end_idx = idx - skip_days
        start_idx = end_idx - lookback_days
        if start_idx < 0 or end_idx < 0:
            continue
        start_price = closes[start_idx]
        end_price = closes[end_idx]
        if not is_valid_price(float(start_price)) or not is_valid_price(float(end_price)):
            continue
        raw_score[idx] = float(end_price / start_price - 1.0)
    return raw_score


def _compute_raw_score_optimized(closes: np.ndarray, lookback_days: int, skip_days: int) -> np.ndarray:
    count = closes.shape[0]
    raw_score = np.full(count, np.nan, dtype=float)
    offset = lookback_days + skip_days
    if offset >= count:
        return raw_score
    starts = closes[: count - offset]
    ends = closes[lookback_days : count - skip_days]
    valid = (starts > 0.0) & (ends > 0.0)
    out_slice = raw_score[offset:]
    out_slice[valid] = ends[valid] / starts[valid] - 1.0
    return raw_score


def _apply_vol_target_reference(
    position_signal: np.ndarray,
    daily_returns: np.ndarray,
    *,
    skip_days: int,
    vol_window_days: int,
    target_volatility: float,
    max_leverage: float,
) -> np.ndarray:
    count = position_signal.shape[0]
    scaled_weight = position_signal.astype(float)
    annualization = np.sqrt(252.0)
    for idx in range(count):
        end_idx = idx - skip_days
        start_idx = end_idx - vol_window_days + 1
        if start_idx < 1 or end_idx < 1:
            scaled_weight[idx] = 0.0
            continue
        window = daily_returns[start_idx : end_idx + 1]
        vol = float(np.std(window, ddof=1)) * annualization if window.size > 1 else 0.0
        if vol <= 0:
            scaled_weight[idx] = 0.0
            continue
        leverage = min(max_leverage, target_volatility / vol)
        scaled_weight[idx] = position_signal[idx] * max(leverage, 0.0)
    return scaled_weight


def _apply_vol_target_optimized(
    position_signal: np.ndarray,
    daily_returns: np.ndarray,
    *,
    skip_days: int,
    vol_window_days: int,
    target_volatility: float,
    max_leverage: float,
) -> np.ndarray:
    annualization = float(np.sqrt(252.0))
    return _rolling_volatility_position_kernel(
        position_signal.astype(float),
        daily_returns.astype(float),
        int(skip_days),
        int(vol_window_days),
        float(target_volatility),
        float(max_leverage),
        annualization,
    )


@njit(cache=True)
def _rolling_volatility_position_kernel(
    position_signal: np.ndarray,
    daily_returns: np.ndarray,
    skip_days: int,
    vol_window_days: int,
    target_volatility: float,
    max_leverage: float,
    annualization: float,
) -> np.ndarray:
    count = position_signal.shape[0]
    scaled_weight = np.zeros(count, dtype=np.float64)
    for idx in range(count):
        end_idx = idx - skip_days
        start_idx = end_idx - vol_window_days + 1
        if start_idx < 1 or end_idx < 1:
            scaled_weight[idx] = 0.0
            continue
        n = end_idx - start_idx + 1
        if n <= 1:
            scaled_weight[idx] = 0.0
            continue
        total = 0.0
        for j in range(start_idx, end_idx + 1):
            total += daily_returns[j]
        mean = total / n
        var = 0.0
        for j in range(start_idx, end_idx + 1):
            diff = daily_returns[j] - mean
            var += diff * diff
        std = np.sqrt(var / (n - 1))
        vol = std * annualization
        if vol <= 0.0:
            scaled_weight[idx] = 0.0
            continue
        leverage = target_volatility / vol
        if leverage > max_leverage:
            leverage = max_leverage
        if leverage < 0.0:
            leverage = 0.0
        scaled_weight[idx] = position_signal[idx] * leverage
    return scaled_weight


def compute_time_series_momentum(
    prices_by_ticker: dict[str, list[float] | list[dict] | tuple[float, ...]],
    fundamentals_by_ticker: dict[str, dict] | None,
    settings: TimeSeriesMomentumSettings,
) -> TimeSeriesResult:
    params = _resolve_hyperparameters(settings)
    min_points = params.lookback_days + params.skip_days + 1
    returns: list[float] = []
    tickers: list[str] = []
    skipped: dict[str, str] = {}
    metrics: dict[str, dict[str, float]] = {}

    for ticker, prices in prices_by_ticker.items():
        series = list(prices)
        closes, _volumes = extract_close_volume(series)
        if len(closes) < min_points:
            skipped[ticker] = "insufficient_history"
            continue
        arrays = build_time_series_momentum_arrays(closes, settings)
        valid_indices = np.flatnonzero(np.isfinite(arrays.raw_score))
        if valid_indices.size == 0:
            skipped[ticker] = "insufficient_history"
            continue
        latest_idx = int(valid_indices[-1])
        momentum_return = float(arrays.raw_score[latest_idx])
        ticker_metrics = {
            "base": momentum_return,
            "latest_position_signal": float(arrays.position_signal[latest_idx]),
            "latest_confidence": float(arrays.confidence[latest_idx]),
            "latest_scaled_weight": float(arrays.scaled_weight[latest_idx]),
            "latest_tradable_position": float(arrays.tradable_position[latest_idx]),
        }
        end_index = latest_idx - params.skip_days
        if settings.use_volatility_scaling:
            start_index = latest_idx - params.skip_days - params.lookback_days
            vol = estimate_volatility(closes[start_index : end_index + 1])
            if vol and vol > 0:
                ticker_metrics["vol_scaled"] = float(momentum_return / vol)
        if settings.use_multi_horizon:
            blended = multi_horizon_return(closes, end_index, windows=[20, 60, 120])
            if blended is not None:
                ticker_metrics["multi_horizon"] = float(blended)
        metrics[ticker] = ticker_metrics
        tickers.append(ticker)
        returns.append(float(momentum_return))

    if not returns:
        return TimeSeriesResult(
            scores={},
            ranking=[],
            longs=[],
            shorts=[],
            weights={},
            metrics={},
            metadata={
                "lookback_days": settings.lookback_days,
                "skip_days": settings.skip_days,
                "vol_window_days": params.vol_window_days,
                "target_volatility": params.target_volatility,
                "max_leverage": params.max_leverage,
                "top_quantile": settings.top_quantile,
                "bottom_quantile": settings.bottom_quantile,
                "universe": 0,
            },
            skipped=skipped,
        )

    scores = _combine_scores(
        tickers=tickers,
        base_scores=np.asarray(returns, dtype=float),
        metrics=metrics,
        settings=settings,
    )
    order = np.argsort(scores)
    total = scores.shape[0]
    top_n, bottom_n = quantile_bucket_sizes(
        total, settings.top_quantile, settings.bottom_quantile
    )

    bottom_idx = order[:bottom_n]
    top_idx = order[-top_n:] if top_n > 0 else np.array([], dtype=int)

    ranking = [(tickers[idx], float(scores[idx])) for idx in order[::-1]]
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

    score_map = {ticker: float(score) for ticker, score in zip(tickers, scores)}

    return TimeSeriesResult(
        scores=score_map,
        ranking=ranking,
        longs=longs,
        shorts=shorts,
        weights=weights,
        metadata={
            "lookback_days": settings.lookback_days,
            "skip_days": settings.skip_days,
            "vol_window_days": params.vol_window_days,
            "target_volatility": params.target_volatility,
            "max_leverage": params.max_leverage,
            "top_quantile": settings.top_quantile,
            "bottom_quantile": settings.bottom_quantile,
            "universe": total,
            "use_volatility_scaling": settings.use_volatility_scaling,
            "use_residual": settings.use_residual,
            "use_multi_horizon": settings.use_multi_horizon,
            "diagnostics": compute_signal_diagnostics(
                scores=score_map,
                weights=weights,
                prices_by_ticker=prices_by_ticker,
            ),
            "use_zscore": settings.use_zscore,
            "winsorize_sigma": settings.winsorize_sigma,
        },
        metrics=metrics,
        skipped=skipped,
    )


def _combine_scores(
    *,
    tickers: list[str],
    base_scores: np.ndarray,
    metrics: dict[str, dict[str, float]],
    settings: TimeSeriesMomentumSettings,
) -> np.ndarray:
    if settings.use_residual:
        residual = base_scores - float(np.mean(base_scores))
        for ticker, value in zip(tickers, residual):
            metrics.setdefault(ticker, {})["residual"] = float(value)

    combined_scores: list[float] = []
    for ticker, base_score in zip(tickers, base_scores):
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
    score_array = _winsorize_and_zscore(score_array, settings)
    return score_array


def _winsorize_and_zscore(
    scores: np.ndarray,
    settings: TimeSeriesMomentumSettings,
) -> np.ndarray:
    if scores.size == 0:
        return scores
    adjusted = scores.astype(float)
    mean = float(np.mean(adjusted))
    std = float(np.std(adjusted))
    if settings.winsorize_sigma is not None and std > 0:
        lower = mean - settings.winsorize_sigma * std
        upper = mean + settings.winsorize_sigma * std
        adjusted = np.clip(adjusted, lower, upper)
    if settings.use_zscore and std > 0:
        adjusted = (adjusted - mean) / std
    return adjusted

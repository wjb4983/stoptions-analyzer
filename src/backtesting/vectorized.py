"""Vectorized backtesting contracts and reference implementation.

Execution convention: signals are generated on bar close and become executable
on the next bar open. The vectorized implementation enforces this by shifting
positions forward by one bar.
"""

from __future__ import annotations

from dataclasses import dataclass
import importlib.util
from typing import Any, Literal, Protocol

import numpy as np

from .event_driven import VectorizedExecutionAdapter, replay_lifecycle
from .execution import (
    BpsSlippage,
    CarryContext,
    ExecutionContext,
    FeeModel,
    LiquidityContext,
    SlippageModel,
    ZeroFee,
    ZeroSlippage,
)


_pd_spec = importlib.util.find_spec("pandas")
if _pd_spec:
    import pandas as pd
else:
    pd = None

_numba_spec = importlib.util.find_spec("numba")
if _numba_spec:
    from numba import njit
else:

    def njit(*args: Any, **kwargs: Any):
        def _decorator(func: Any) -> Any:
            return func

        return _decorator


ExecutionMode = Literal["reference", "optimized"]


class VectorizedBacktestRunner(Protocol):
    """Contract for a pure vectorized backtest engine.

    Any implementation must preserve causal ordering: a signal observed at time
    ``t`` (close) can only affect position from ``t+1`` (next open) onward.
    """

    def run(
        self,
        prices: Any,
        signals: Any,
        *,
        slippage_model: SlippageModel | None = None,
        fee_model: FeeModel | None = None,
        initial_equity: float = 1.0,
    ) -> "BacktestResult":
        """Execute a backtest and return a result object."""


@dataclass(frozen=True)
class BacktestResult:
    """Vectorized backtest outputs with portfolio-level series.

    Attributes:
        equity_curve: Portfolio equity over time in quote currency.
        positions: Position units active each bar after next-open execution lag.
        returns: Net period returns (decimal form, e.g. ``0.01`` for 1%).
        daily_returns: Alias for net period returns.
        pnl: Period profit/loss in quote currency.
        trades: Trade deltas per period after applying execution lag.
        turnover: Sum of absolute trades per period.
        cost_breakdown: Cost components with per-period and total costs.
        metrics: Aggregate performance metrics keyed by metric name.
    """

    equity_curve: Any
    positions: Any
    returns: Any
    daily_returns: Any
    pnl: Any
    trades: Any
    turnover: Any
    cost_breakdown: dict[str, Any]
    metrics: dict[str, float]
    fills: list[dict[str, float | int | str]]
    execution_events: list[dict[str, Any]]


def backtest_vectorized(
    prices: Any,
    signals: Any,
    *,
    slippage_model: SlippageModel | None = None,
    fee_model: FeeModel | None = None,
    borrow_cost_model: Any | None = None,
    volumes: Any | None = None,
    adv: Any | None = None,
    volatility: Any | None = None,
    spread_bps: Any | None = None,
    order_type: str = "market",
    latency_bars: int = 0,
    latency_ms: int = 0,
    queue_rank_proxy: Any | None = 0.5,
    available_bar_volume: Any | None = None,
    max_participation_per_bar: float | None = None,
    weights: Any | None = None,
    initial_equity: float = 1.0,
    execution_mode: ExecutionMode = "optimized",
    timeframe: str | None = None,
    periods_per_year: float | None = None,
    carry_asset_classes: list[str] | tuple[str, ...] | None = None,
    carry_expiry_by_asset: list[str | None] | tuple[str | None, ...] | None = None,
    carry_multipliers: list[float] | tuple[float, ...] | None = None,
    carry_borrow_availability_tiers: list[str] | tuple[str, ...] | None = None,
    carry_financing_benchmarks: list[str] | tuple[str, ...] | None = None,
    borrow_rate_series: Any | None = None,
    borrow_available_flags: Any | None = None,
    symbols: list[str] | tuple[str, ...] | None = None,
    dates: Any | None = None,
    time_in_force: str = "gtc",
    urgency: str = "normal",
) -> BacktestResult:
    """Run a pure vectorized backtest with next-open execution.

    Assumptions:
        * ``signals[t]`` is generated from information known at close of bar ``t``.
        * Execution occurs at open of bar ``t+1`` via a one-bar position shift.
        * Slippage modifies the execution price; fees reduce return directly.
        * Prices represent a single tradable series in quote currency.

    Args:
        prices: 1D array-like close-price series (float, quote currency per unit).
        signals: 1D array-like desired exposure/position signal for each bar.
        slippage_model: Optional execution slippage model.
        fee_model: Optional transaction fee model.
        initial_equity: Starting portfolio equity in quote currency.

    Returns:
        ``BacktestResult`` containing aligned time series and summary metrics.
    """

    price_values, index = _to_array(prices)
    signal_values, _ = _to_array(signals)

    price_values = _ensure_2d(price_values)
    signal_values = _ensure_2d(signal_values)

    if price_values.shape[0] != signal_values.shape[0]:
        raise ValueError("Prices and signals must have the same length.")
    if price_values.shape[1] != signal_values.shape[1]:
        raise ValueError("Prices and signals must have the same number of assets.")

    n_periods, n_assets = price_values.shape
    portfolio_weights = _normalize_weights(weights, n_assets)

    positions = _shift(signal_values, 1)
    requested_trades = positions - _shift(positions, 1)

    gross_asset_returns = np.zeros_like(price_values, dtype=float)
    gross_asset_returns[1:] = price_values[1:] / price_values[:-1] - 1.0
    gross_returns = np.sum(positions * gross_asset_returns * portfolio_weights, axis=1)

    execution_model = slippage_model or ZeroSlippage()
    fees = fee_model or ZeroFee()
    borrow_costs = borrow_cost_model
    if borrow_costs is None:
        from .execution import CarryModel

        borrow_costs = CarryModel(
            asset_classes=carry_asset_classes or ["equity"] * n_assets,
            expiry_by_asset=carry_expiry_by_asset,
            multipliers=carry_multipliers,
            borrow_availability_tiers=carry_borrow_availability_tiers,
            financing_benchmarks=carry_financing_benchmarks,
            borrow_rate_series=np.asarray(borrow_rate_series, dtype=float) if borrow_rate_series is not None else None,
            borrow_available_flags=np.asarray(borrow_available_flags, dtype=bool) if borrow_available_flags is not None else None,
            periods_per_year=_resolve_periods_per_year(timeframe=timeframe, periods_per_year=periods_per_year),
        )

    liquidity = _build_liquidity_context(
        prices=price_values,
        volumes=volumes,
        adv=adv,
        volatility=volatility,
        spread_bps=spread_bps,
        queue_rank_proxy=queue_rank_proxy,
        available_bar_volume=available_bar_volume,
        max_participation_per_bar=max_participation_per_bar,
    )

    adapter = VectorizedExecutionAdapter(max_participation_per_bar=liquidity["max_participation_per_bar"][0, 0])
    resolved_symbols = [str(symbol) for symbol in (symbols or [f"asset_{idx}" for idx in range(n_assets)])]
    resolved_timestamps = [None if index is None else str(index[idx]) for idx in range(n_periods)]
    trades, residual_orders, fill_events, lifecycle_events = adapter.execute(
        requested_trades=requested_trades,
        prices=price_values,
        available_volume=liquidity["available_bar_volume"],
        queue_rank_proxy=liquidity["queue_rank_proxy"],
        order_type=order_type,
        latency_bars=latency_bars,
        latency_ms=latency_ms,
        symbols=resolved_symbols,
        timestamps=resolved_timestamps,
        time_in_force=time_in_force,
        urgency=urgency,
    )

    slippage_cost = _estimate_slippage(
        price_values,
        trades,
        execution_model,
        portfolio_weights,
        mode=execution_mode,
        liquidity=liquidity,
    )
    fee_cost = _estimate_fees(
        price_values,
        trades,
        fees,
        portfolio_weights,
        mode=execution_mode,
        liquidity=liquidity,
    )
    carry_result = _estimate_borrow_cost(
        positions,
        borrow_costs,
        portfolio_weights,
        dates=np.asarray(dates) if dates is not None else None,
        symbols=tuple(symbols) if symbols is not None else None,
        metadata={
            "asset_classes": list(carry_asset_classes or ["equity"] * n_assets),
            "expiry": list(carry_expiry_by_asset or [None] * n_assets),
            "multipliers": list(carry_multipliers or [1.0] * n_assets),
            "borrow_availability_tiers": list(carry_borrow_availability_tiers or ["normal"] * n_assets),
            "financing_benchmarks": list(carry_financing_benchmarks or ["overnight"] * n_assets),
        },
    )
    borrow_cost = carry_result["portfolio"]

    net_returns = gross_returns - slippage_cost - fee_cost - borrow_cost

    equity_curve = initial_equity * np.cumprod(1.0 + net_returns)
    pnl = np.zeros_like(equity_curve)
    pnl[1:] = equity_curve[:-1] * net_returns[1:]
    turnover = np.sum(np.abs(trades) * portfolio_weights, axis=1)

    metrics = _compute_metrics(
        returns=net_returns,
        equity_curve=equity_curve,
        positions=positions,
        turnover=turnover,
        timeframe=timeframe,
        periods_per_year=periods_per_year,
    )
    cost_breakdown = {
        "slippage": _to_series(slippage_cost, index),
        "fees": _to_series(fee_cost, index),
        "borrow": _to_series(borrow_cost, index),
        "borrow_by_asset": _to_aligned_output(carry_result["weighted_by_asset"], index),
        "carry_attribution_by_asset": _to_aligned_output(carry_result["raw_by_asset"], index),
        "total": _to_series(slippage_cost + fee_cost + borrow_cost, index),
        "totals": {
            "slippage": float(np.sum(slippage_cost)),
            "fees": float(np.sum(fee_cost)),
            "borrow": float(np.sum(borrow_cost)),
            "total": float(np.sum(slippage_cost + fee_cost + borrow_cost)),
        },
    }

    return BacktestResult(
        equity_curve=_to_series(equity_curve, index),
        positions=_to_aligned_output(positions, index),
        returns=_to_series(net_returns, index),
        daily_returns=_to_series(net_returns, index),
        pnl=_to_series(pnl, index),
        trades=_to_aligned_output(trades, index),
        turnover=_to_series(turnover, index),
        cost_breakdown=cost_breakdown,
        metrics=metrics,
        fills=[
            {
                "bar_index": int(evt.bar_index),
                "asset_index": int(evt.asset_index),
                "requested_size": float(evt.requested_size),
                "filled_size": float(evt.filled_size),
                "residual_size": float(evt.residual_size),
                "participation_rate": float(evt.participation_rate),
                "available_volume": float(evt.available_volume),
                "order_type": str(evt.order_type),
                "latency_bars": int(evt.latency_bars),
                "latency_ms": int(evt.latency_ms),
                "queue_rank_proxy": float(evt.queue_rank_proxy),
            }
            for evt in fill_events
        ],
        execution_events=replay_lifecycle([event.__dict__ for event in lifecycle_events]),
    )


def _shift(values: np.ndarray, periods: int) -> np.ndarray:
    shifted = np.zeros_like(values, dtype=float)
    if periods <= 0:
        shifted[:] = values
        return shifted
    shifted[periods:] = values[:-periods]
    return shifted


def _estimate_slippage(
    prices: np.ndarray,
    trades: np.ndarray,
    model: SlippageModel,
    weights: np.ndarray,
    *,
    mode: ExecutionMode,
    liquidity: dict[str, np.ndarray],
) -> np.ndarray:
    if mode == "optimized":
        optimized = _estimate_slippage_optimized(trades, model, weights)
        if optimized is not None:
            return optimized
    return _estimate_slippage_reference(prices, trades, model, weights, liquidity)


def _estimate_slippage_reference(
    prices: np.ndarray,
    trades: np.ndarray,
    model: SlippageModel,
    weights: np.ndarray,
    liquidity: dict[str, np.ndarray],
) -> np.ndarray:
    slippage = np.zeros(prices.shape[0], dtype=float)
    if not trades.size:
        return slippage
    for idx in range(prices.shape[0]):
        trade_sizes = trades[idx]
        for asset_idx, trade_size in enumerate(trade_sizes):
            if trade_size == 0:
                continue
            price = float(prices[idx, asset_idx])
            exec_price = float(
                model.apply(
                    price,
                    float(trade_size),
                    _build_bar_context(liquidity, idx, asset_idx, float(trade_size)),
                )
            )
            slippage[idx] += (exec_price - price) / price * trade_size * weights[asset_idx]
    return slippage


def _estimate_slippage_optimized(
    trades: np.ndarray,
    model: SlippageModel,
    weights: np.ndarray,
) -> np.ndarray | None:
    if isinstance(model, ZeroSlippage):
        return np.zeros(trades.shape[0], dtype=float)
    if isinstance(model, BpsSlippage):
        bps_factor = float(model.bps / 10_000.0)
        return _slippage_bps_kernel(trades, weights, bps_factor)
    return None


def _estimate_fees(
    prices: np.ndarray,
    trades: np.ndarray,
    model: FeeModel,
    weights: np.ndarray,
    *,
    mode: ExecutionMode,
    liquidity: dict[str, np.ndarray],
) -> np.ndarray:
    if mode == "optimized":
        optimized = _estimate_fees_optimized(trades, model, weights)
        if optimized is not None:
            return optimized
    return _estimate_fees_reference(prices, trades, model, weights, liquidity)


def _estimate_fees_reference(
    prices: np.ndarray,
    trades: np.ndarray,
    model: FeeModel,
    weights: np.ndarray,
    liquidity: dict[str, np.ndarray],
) -> np.ndarray:
    fee_returns = np.zeros(prices.shape[0], dtype=float)
    if not trades.size:
        return fee_returns
    for idx in range(prices.shape[0]):
        trade_sizes = trades[idx]
        for asset_idx, trade_size in enumerate(trade_sizes):
            if trade_size == 0:
                continue
            fee = float(
                model.calculate(
                    float(prices[idx, asset_idx]),
                    float(trade_size),
                    _build_bar_context(liquidity, idx, asset_idx, float(trade_size)),
                )
            )
            fee_returns[idx] += fee * weights[asset_idx]
    return fee_returns


def _estimate_fees_optimized(
    trades: np.ndarray,
    model: FeeModel,
    weights: np.ndarray,
) -> np.ndarray | None:
    if isinstance(model, ZeroFee):
        return np.zeros(trades.shape[0], dtype=float)
    if hasattr(model, "commission_per_trade"):
        commission = float(getattr(model, "commission_per_trade"))
        return _fixed_commission_kernel(trades, weights, commission)
    return None




def _build_liquidity_context(
    *,
    prices: np.ndarray,
    volumes: Any | None,
    adv: Any | None,
    volatility: Any | None,
    spread_bps: Any | None,
    queue_rank_proxy: Any | None,
    available_bar_volume: Any | None,
    max_participation_per_bar: float | None,
) -> dict[str, np.ndarray]:
    n_periods, n_assets = prices.shape
    volumes_arr = _coerce_liquidity_array(volumes, prices.shape, default=1.0)
    adv_arr = _coerce_liquidity_array(adv, prices.shape, default=np.nan)
    if np.isnan(adv_arr).all():
        adv_arr = _rolling_mean(volumes_arr, window=20)
    else:
        adv_arr = np.where(np.isnan(adv_arr), _rolling_mean(volumes_arr, window=20), adv_arr)
    volatility_arr = _coerce_liquidity_array(volatility, prices.shape, default=np.nan)
    if np.isnan(volatility_arr).all():
        volatility_arr = _rolling_volatility(prices, window=20)
    else:
        volatility_arr = np.where(np.isnan(volatility_arr), _rolling_volatility(prices, window=20), volatility_arr)
    spread_arr = _coerce_liquidity_array(spread_bps, prices.shape, default=2.0)
    queue_arr = _coerce_liquidity_array(queue_rank_proxy, prices.shape, default=0.5)
    available_volume_arr = _coerce_liquidity_array(available_bar_volume, prices.shape, default=np.nan)
    available_volume_arr = np.where(np.isnan(available_volume_arr), volumes_arr, available_volume_arr)
    max_participation = 1.0 if max_participation_per_bar is None else float(max_participation_per_bar)
    max_participation_arr = np.full(prices.shape, max_participation, dtype=float)
    return {
        "volume": volumes_arr,
        "adv": adv_arr,
        "volatility": volatility_arr,
        "spread_bps": spread_arr,
        "queue_rank_proxy": queue_arr,
        "available_bar_volume": available_volume_arr,
        "max_participation_per_bar": max_participation_arr,
    }


def _coerce_liquidity_array(values: Any | None, shape: tuple[int, int], default: float) -> np.ndarray:
    if values is None:
        return np.full(shape, float(default), dtype=float)
    arr = np.asarray(values, dtype=float)
    if arr.ndim == 0:
        return np.full(shape, float(arr), dtype=float)
    if arr.ndim == 1:
        if arr.shape[0] == shape[0]:
            return np.repeat(arr.reshape(-1, 1), shape[1], axis=1)
        if arr.shape[0] == shape[1]:
            return np.repeat(arr.reshape(1, -1), shape[0], axis=0)
        raise ValueError("Liquidity vectors must match either n_periods or n_assets.")
    if arr.shape != shape:
        raise ValueError("Liquidity arrays must match price shape.")
    return arr


def _rolling_mean(values: np.ndarray, window: int) -> np.ndarray:
    out = np.zeros_like(values, dtype=float)
    for col in range(values.shape[1]):
        for idx in range(values.shape[0]):
            start = max(0, idx - window + 1)
            out[idx, col] = float(np.mean(values[start:idx + 1, col]))
    return out


def _rolling_volatility(prices: np.ndarray, window: int) -> np.ndarray:
    rets = np.zeros_like(prices, dtype=float)
    rets[1:] = prices[1:] / np.where(prices[:-1] == 0.0, 1.0, prices[:-1]) - 1.0
    out = np.zeros_like(prices, dtype=float)
    for col in range(prices.shape[1]):
        for idx in range(prices.shape[0]):
            start = max(0, idx - window + 1)
            segment = rets[start:idx + 1, col]
            out[idx, col] = float(np.std(segment, ddof=1)) if segment.size > 1 else 0.0
    return out


def _build_bar_context(
    liquidity: dict[str, np.ndarray],
    idx: int,
    asset_idx: int,
    trade_size: float = 0.0,
) -> LiquidityContext | ExecutionContext:
    available = max(float(liquidity["available_bar_volume"][idx, asset_idx]), 1e-12)
    participation = abs(float(trade_size)) / available if trade_size != 0.0 else 0.0
    return ExecutionContext(
        bar_index=int(idx),
        asset_index=int(asset_idx),
        volume=float(liquidity["volume"][idx, asset_idx]),
        adv=float(liquidity["adv"][idx, asset_idx]),
        volatility=float(liquidity["volatility"][idx, asset_idx]),
        spread_bps=float(liquidity["spread_bps"][idx, asset_idx]),
        order_type="market",
        latency_bars=0,
        latency_ms=0,
        queue_rank_proxy=float(liquidity["queue_rank_proxy"][idx, asset_idx]),
        available_bar_volume=float(liquidity["available_bar_volume"][idx, asset_idx]),
        max_participation_per_bar=float(liquidity["max_participation_per_bar"][idx, asset_idx]),
        realized_participation=float(participation),
    )

def _estimate_borrow_cost(
    positions: np.ndarray,
    model: Any,
    weights: np.ndarray,
    *,
    dates: np.ndarray | None,
    symbols: tuple[str, ...] | None,
    metadata: dict[str, Any] | None,
) -> dict[str, np.ndarray]:
    context = CarryContext(dates=dates, symbols=symbols, metadata=metadata)
    try:
        borrow_by_asset = np.asarray(model.calculate(positions, context), dtype=float)
    except TypeError:
        borrow_by_asset = np.asarray(model.calculate(positions), dtype=float)
    if borrow_by_asset.ndim == 1:
        borrow_by_asset = borrow_by_asset.reshape(-1, 1)
    weighted_by_asset = borrow_by_asset * weights.reshape(1, -1)
    return {
        "raw_by_asset": borrow_by_asset,
        "weighted_by_asset": weighted_by_asset,
        "portfolio": np.sum(weighted_by_asset, axis=1),
    }


def _to_array(values: Any) -> tuple[np.ndarray, Any]:
    if pd is not None and isinstance(values, pd.Series):
        return values.to_numpy(dtype=float), values.index
    return np.asarray(values, dtype=float), None


def _to_series(values: np.ndarray, index: Any) -> Any:
    if pd is not None and index is not None:
        return pd.Series(values, index=index)
    return values


def _to_aligned_output(values: np.ndarray, index: Any) -> Any:
    if values.ndim == 2 and values.shape[1] == 1:
        return _to_series(values[:, 0], index)
    if pd is not None and index is not None:
        return pd.DataFrame(values, index=index)
    return values


def _ensure_2d(values: np.ndarray) -> np.ndarray:
    if values.ndim == 1:
        return values.reshape(-1, 1)
    if values.ndim != 2:
        raise ValueError("Inputs must be 1D or 2D array-like values.")
    return values


def _normalize_weights(weights: Any | None, n_assets: int) -> np.ndarray:
    if weights is None:
        return np.ones(n_assets, dtype=float) / n_assets
    arr = np.asarray(weights, dtype=float)
    if arr.ndim != 1 or arr.shape[0] != n_assets:
        raise ValueError("weights must be a 1D array-like with one value per asset.")
    return arr


def _resolve_periods_per_year(*, timeframe: str | None, periods_per_year: float | None) -> float:
    if periods_per_year is not None and periods_per_year > 0:
        return float(periods_per_year)
    if timeframe is None:
        return 252.0
    mapping = {
        "1m": 252.0 * 390.0,
        "5m": 252.0 * 78.0,
        "15m": 252.0 * 26.0,
        "30m": 252.0 * 13.0,
        "1h": 252.0 * 6.5,
        "1d": 252.0,
    }
    return float(mapping.get(str(timeframe).strip().lower(), 252.0))


def _compute_metrics(
    *,
    returns: np.ndarray,
    equity_curve: np.ndarray,
    positions: np.ndarray,
    turnover: np.ndarray,
    timeframe: str | None,
    periods_per_year: float | None,
) -> dict[str, float]:
    ann_factor = _resolve_periods_per_year(timeframe=timeframe, periods_per_year=periods_per_year)
    total_return = equity_curve[-1] / equity_curve[0] - 1.0 if equity_curve.size else 0.0
    avg_return = float(np.mean(returns)) if returns.size else 0.0
    vol = float(np.std(returns, ddof=1)) if returns.size > 1 else 0.0
    sharpe = avg_return / vol * np.sqrt(ann_factor) if vol else 0.0
    downside = np.minimum(returns, 0.0)
    downside_dev = float(np.sqrt(np.mean(np.square(downside)))) if returns.size else 0.0
    sortino = avg_return / downside_dev * np.sqrt(ann_factor) if downside_dev else 0.0

    running_peak = np.maximum.accumulate(equity_curve) if equity_curve.size else np.array([], dtype=float)
    safe_peak = np.where(running_peak == 0.0, 1.0, running_peak) if equity_curve.size else np.array([], dtype=float)
    drawdown = equity_curve / safe_peak - 1.0 if equity_curve.size else np.array([], dtype=float)
    max_drawdown = float(np.min(drawdown)) if drawdown.size else 0.0
    n_periods = int(returns.size)
    if n_periods > 0 and equity_curve.size and equity_curve[0] != 0 and ann_factor > 0:
        cagr = float((equity_curve[-1] / equity_curve[0]) ** (ann_factor / n_periods) - 1.0)
    else:
        cagr = 0.0
    calmar = cagr / abs(max_drawdown) if max_drawdown < 0 else 0.0

    centered = returns - avg_return if returns.size else np.array([], dtype=float)
    if returns.size > 2:
        m2 = float(np.mean(centered**2))
        m3 = float(np.mean(centered**3))
        skew = m3 / (m2 ** 1.5) if m2 > 0 else 0.0
    else:
        skew = 0.0
    if returns.size > 3:
        m2 = float(np.mean(centered**2))
        m4 = float(np.mean(centered**4))
        kurtosis = m4 / (m2**2) - 3.0 if m2 > 0 else 0.0
    else:
        kurtosis = 0.0

    positive_returns = returns[returns > 0.0]
    negative_returns = returns[returns < 0.0]
    hit_rate = float(np.mean(returns > 0.0)) if returns.size else 0.0
    gross_profit = float(np.sum(positive_returns)) if positive_returns.size else 0.0
    gross_loss = float(-np.sum(negative_returns)) if negative_returns.size else 0.0
    profit_factor = gross_profit / gross_loss if gross_loss > 0 else 0.0

    abs_positions = np.abs(positions)
    if abs_positions.ndim == 2:
        exposure = float(np.mean(np.sum(abs_positions, axis=1) > 0.0))
    else:
        exposure = float(np.mean(abs_positions > 0.0)) if abs_positions.size else 0.0

    turnover_total = float(np.sum(turnover)) if turnover.size else 0.0
    turnover_adjusted_return = float(total_return / (1.0 + turnover_total))

    rolling_window = int(max(5, min(60, max(1, int(np.sqrt(max(n_periods, 1)))))))
    rolling_sharpes: list[float] = []
    rolling_mdds: list[float] = []
    if n_periods >= rolling_window:
        for idx in range(rolling_window, n_periods + 1):
            win_returns = returns[idx - rolling_window:idx]
            win_mean = float(np.mean(win_returns))
            win_vol = float(np.std(win_returns, ddof=1)) if win_returns.size > 1 else 0.0
            rolling_sharpes.append(win_mean / win_vol * np.sqrt(ann_factor) if win_vol else 0.0)
            win_equity = equity_curve[idx - rolling_window:idx]
            win_peak = np.maximum.accumulate(win_equity)
            win_safe_peak = np.where(win_peak == 0.0, 1.0, win_peak)
            win_dd = win_equity / win_safe_peak - 1.0
            rolling_mdds.append(float(np.min(win_dd)) if win_dd.size else 0.0)

    rolling_sharpe_mean = float(np.mean(rolling_sharpes)) if rolling_sharpes else 0.0
    rolling_sharpe_min = float(np.min(rolling_sharpes)) if rolling_sharpes else 0.0
    rolling_sharpe_max = float(np.max(rolling_sharpes)) if rolling_sharpes else 0.0
    rolling_drawdown_mean = float(np.mean(rolling_mdds)) if rolling_mdds else 0.0
    rolling_drawdown_worst = float(np.min(rolling_mdds)) if rolling_mdds else 0.0

    return {
        "periods_per_year": float(ann_factor),
        "total_return": float(total_return),
        "cagr": cagr,
        "avg_return": avg_return,
        "volatility": vol,
        "sharpe": sharpe,
        "sortino": sortino,
        "downside_deviation": downside_dev,
        "max_drawdown": max_drawdown,
        "calmar": calmar,
        "skew": float(skew),
        "kurtosis": float(kurtosis),
        "hit_rate": hit_rate,
        "profit_factor": float(profit_factor),
        "exposure_time": exposure,
        "turnover_adjusted_return": turnover_adjusted_return,
        "rolling_window": float(rolling_window),
        "rolling_sharpe_mean": rolling_sharpe_mean,
        "rolling_sharpe_min": rolling_sharpe_min,
        "rolling_sharpe_max": rolling_sharpe_max,
        "rolling_drawdown_mean": rolling_drawdown_mean,
        "rolling_drawdown_worst": rolling_drawdown_worst,
    }


@njit(cache=True)
def _slippage_bps_kernel(trades: np.ndarray, weights: np.ndarray, bps_factor: float) -> np.ndarray:
    n_periods, n_assets = trades.shape
    slippage = np.zeros(n_periods, dtype=np.float64)
    for idx in range(n_periods):
        total = 0.0
        for asset_idx in range(n_assets):
            trade_size = trades[idx, asset_idx]
            if trade_size != 0.0:
                total += abs(trade_size) * weights[asset_idx]
        slippage[idx] = bps_factor * total
    return slippage


@njit(cache=True)
def _fixed_commission_kernel(
    trades: np.ndarray,
    weights: np.ndarray,
    commission: float,
) -> np.ndarray:
    n_periods, n_assets = trades.shape
    fees = np.zeros(n_periods, dtype=np.float64)
    for idx in range(n_periods):
        total = 0.0
        for asset_idx in range(n_assets):
            if trades[idx, asset_idx] != 0.0:
                total += weights[asset_idx]
        fees[idx] = commission * total
    return fees


def replay_from_event_logs(event_logs: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Return a deterministic normalized lifecycle stream for exact run reconstruction."""

    return replay_lifecycle(event_logs)

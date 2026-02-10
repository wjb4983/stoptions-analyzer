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

from .execution import (
    BpsSlippage,
    FeeModel,
    ShortBorrowCost,
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


def backtest_vectorized(
    prices: Any,
    signals: Any,
    *,
    slippage_model: SlippageModel | None = None,
    fee_model: FeeModel | None = None,
    borrow_cost_model: ShortBorrowCost | None = None,
    weights: Any | None = None,
    initial_equity: float = 1.0,
    execution_mode: ExecutionMode = "optimized",
    timeframe: str | None = None,
    periods_per_year: float | None = None,
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
    trades = positions - _shift(positions, 1)

    gross_asset_returns = np.zeros_like(price_values, dtype=float)
    gross_asset_returns[1:] = price_values[1:] / price_values[:-1] - 1.0
    gross_returns = np.sum(positions * gross_asset_returns * portfolio_weights, axis=1)

    execution_model = slippage_model or ZeroSlippage()
    fees = fee_model or ZeroFee()
    borrow_costs = borrow_cost_model or ShortBorrowCost(0.0)

    slippage_cost = _estimate_slippage(
        price_values,
        trades,
        execution_model,
        portfolio_weights,
        mode=execution_mode,
    )
    fee_cost = _estimate_fees(
        price_values,
        trades,
        fees,
        portfolio_weights,
        mode=execution_mode,
    )
    borrow_cost = _estimate_borrow_cost(positions, borrow_costs, portfolio_weights)

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
) -> np.ndarray:
    if mode == "optimized":
        optimized = _estimate_slippage_optimized(trades, model, weights)
        if optimized is not None:
            return optimized
    return _estimate_slippage_reference(prices, trades, model, weights)


def _estimate_slippage_reference(
    prices: np.ndarray,
    trades: np.ndarray,
    model: SlippageModel,
    weights: np.ndarray,
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
                    {"bar_index": idx, "asset_index": asset_idx},
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
) -> np.ndarray:
    if mode == "optimized":
        optimized = _estimate_fees_optimized(trades, model, weights)
        if optimized is not None:
            return optimized
    return _estimate_fees_reference(prices, trades, model, weights)


def _estimate_fees_reference(
    prices: np.ndarray,
    trades: np.ndarray,
    model: FeeModel,
    weights: np.ndarray,
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
                    {"bar_index": idx, "asset_index": asset_idx},
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


def _estimate_borrow_cost(
    positions: np.ndarray,
    model: ShortBorrowCost,
    weights: np.ndarray,
) -> np.ndarray:
    borrow_by_asset = model.calculate(positions)
    return np.sum(borrow_by_asset * weights, axis=1)


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

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

    metrics = _compute_metrics(net_returns, equity_curve)
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


def _compute_metrics(returns: np.ndarray, equity_curve: np.ndarray) -> dict[str, float]:
    total_return = equity_curve[-1] / equity_curve[0] - 1.0 if equity_curve.size else 0.0
    avg_return = float(np.mean(returns)) if returns.size else 0.0
    vol = float(np.std(returns, ddof=1)) if returns.size > 1 else 0.0
    sharpe = avg_return / vol * np.sqrt(252.0) if vol else 0.0
    return {
        "total_return": float(total_return),
        "avg_return": avg_return,
        "volatility": vol,
        "sharpe": sharpe,
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

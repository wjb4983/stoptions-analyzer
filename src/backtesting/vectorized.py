"""Vectorized backtesting contracts and reference implementation.

Execution convention: signals are generated on bar close and become executable
on the next bar open. The vectorized implementation enforces this by shifting
positions forward by one bar.
"""

from __future__ import annotations

from dataclasses import dataclass
import importlib.util
from typing import Any, Protocol

import numpy as np

from .execution import FeeModel, SlippageModel, ZeroFee, ZeroSlippage


_pd_spec = importlib.util.find_spec("pandas")
if _pd_spec:
    import pandas as pd
else:
    pd = None


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


class BpsSlippage:
    """Fixed basis-point slippage model.

    Args:
        bps: Basis points of notional applied when a trade occurs.
    """

    def __init__(self, bps: float) -> None:
        if bps < 0:
            raise ValueError("Slippage bps must be non-negative.")
        self.bps = bps

    def apply(self, price: float, size: float, liquidity_context: Any | None = None) -> float:
        """Return execution price shifted by ``bps`` in trade direction."""

        if size == 0:
            return price
        impact = (self.bps / 10_000.0) * np.sign(size)
        return price * (1.0 + impact)


@dataclass(frozen=True)
class BacktestResult:
    """Vectorized backtest outputs with portfolio-level series.

    Attributes:
        equity_curve: Portfolio equity over time in quote currency.
        positions: Position units active each bar after next-open execution lag.
        returns: Net period returns (decimal form, e.g. ``0.01`` for 1%).
        pnl: Period profit/loss in quote currency.
        metrics: Aggregate performance metrics keyed by metric name.
    """

    equity_curve: Any
    positions: Any
    returns: Any
    pnl: Any
    metrics: dict[str, float]


def backtest_vectorized(
    prices: Any,
    signals: Any,
    *,
    slippage_model: SlippageModel | None = None,
    fee_model: FeeModel | None = None,
    initial_equity: float = 1.0,
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

    if price_values.shape[0] != signal_values.shape[0]:
        raise ValueError("Prices and signals must have the same length.")
    if price_values.ndim != 1:
        raise ValueError("Prices must be a 1D array-like series.")

    positions = _shift(signal_values, 1)
    trades = positions - _shift(positions, 1)

    returns = np.zeros_like(price_values, dtype=float)
    returns[1:] = price_values[1:] / price_values[:-1] - 1.0

    execution_model = slippage_model or ZeroSlippage()
    fees = fee_model or ZeroFee()
    slippage_cost = _estimate_slippage(price_values, trades, execution_model)
    fee_cost = _estimate_fees(price_values, trades, fees)
    net_returns = positions * returns - slippage_cost - fee_cost

    equity_curve = initial_equity * np.cumprod(1.0 + net_returns)
    pnl = np.zeros_like(equity_curve)
    pnl[1:] = equity_curve[:-1] * net_returns[1:]

    metrics = _compute_metrics(net_returns, equity_curve)

    return BacktestResult(
        equity_curve=_to_series(equity_curve, index),
        positions=_to_series(positions, index),
        returns=_to_series(net_returns, index),
        pnl=_to_series(pnl, index),
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
) -> np.ndarray:
    slippage = np.zeros_like(prices, dtype=float)
    if not trades.size:
        return slippage
    for idx, trade_size in enumerate(trades):
        if trade_size == 0:
            continue
        price = float(prices[idx])
        exec_price = float(model.apply(price, float(trade_size), {"bar_index": idx}))
        slippage[idx] = (exec_price - price) / price * trade_size
    return slippage


def _estimate_fees(prices: np.ndarray, trades: np.ndarray, model: FeeModel) -> np.ndarray:
    fee_returns = np.zeros_like(prices, dtype=float)
    if not trades.size:
        return fee_returns
    for idx, trade_size in enumerate(trades):
        if trade_size == 0:
            continue
        fee = float(model.calculate(float(prices[idx]), float(trade_size), {"bar_index": idx}))
        fee_returns[idx] = fee
    return fee_returns


def _to_array(values: Any) -> tuple[np.ndarray, Any]:
    if pd is not None and isinstance(values, pd.Series):
        return values.to_numpy(dtype=float), values.index
    return np.asarray(values, dtype=float), None


def _to_series(values: np.ndarray, index: Any) -> Any:
    if pd is not None and index is not None:
        return pd.Series(values, index=index)
    return values


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

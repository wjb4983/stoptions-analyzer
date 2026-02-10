"""Vectorized backtesting helpers.

Signals are generated on the bar close and are executed on the next bar open.
This is implemented by shifting signals forward one bar when creating positions
to avoid lookahead bias.
"""

from __future__ import annotations

from dataclasses import dataclass
import importlib.util
from typing import Any

import numpy as np

from .execution import SlippageModel, ZeroSlippage


_pd_spec = importlib.util.find_spec("pandas")
if _pd_spec:
    import pandas as pd
else:
    pd = None


class BpsSlippage:
    """Apply a fixed basis-point slippage cost per trade.

    Args:
        bps: Basis points of notional applied when a trade occurs.
    """

    def __init__(self, bps: float) -> None:
        if bps < 0:
            raise ValueError("Slippage bps must be non-negative.")
        self.bps = bps

    def apply(self, price: float, size: float, liquidity_context: Any | None = None) -> float:
        if size == 0:
            return price
        impact = (self.bps / 10_000.0) * np.sign(size)
        return price * (1.0 + impact)


@dataclass(frozen=True)
class BacktestResult:
    """Container for vectorized backtest results."""

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
    initial_equity: float = 1.0,
) -> BacktestResult:
    """Run a vectorized backtest.

    Signals are generated at the bar close and executed at the next bar open.
    Positions are created by explicitly shifting signals forward one bar to
    prevent lookahead bias.
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
    slippage = _estimate_slippage(price_values, trades, execution_model)
    net_returns = positions * returns - slippage

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
    """Shift a 1D array forward by periods, filling with zeros."""

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

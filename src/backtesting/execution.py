"""Execution and transaction cost contracts for backtesting engines.

Execution convention: strategy signals are generated on a bar close and any
resulting order is filled on the next bar open. Slippage and fees are applied
at that next-open execution time.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Protocol

import numpy as np


class SlippageModel(Protocol):
    """Contract for execution price adjustments due to slippage.

    Implementations return the executed price for a requested trade.

    Args:
        price: Reference execution price in quote currency per unit.
        size: Signed quantity in units. Positive means buy, negative means sell.
        liquidity_context: Optional market context (bar data, order book hints,
            volatility regime, etc.) used by the model.

    Returns:
        Executed unit price after slippage in quote currency per unit.
    """

    def apply(self, price: float, size: float, liquidity_context: Any | None = None) -> float:
        """Return the executed price after slippage."""


class FeeModel(Protocol):
    """Contract for fees assessed at order execution.

    Args:
        price: Executed unit price in quote currency per unit.
        size: Signed quantity in units. Absolute value is traded size.
        liquidity_context: Optional market context for tiered fee schedules.

    Returns:
        Total transaction fee in quote currency for the trade.
    """

    def calculate(self, price: float, size: float, liquidity_context: Any | None = None) -> float:
        """Return total fee in quote currency for an execution."""


class ZeroSlippage:
    """Slippage model that leaves execution price unchanged."""

    def apply(self, price: float, size: float, liquidity_context: Any | None = None) -> float:
        """Return ``price`` unchanged."""

        return price


class ZeroFee:
    """Fee model that charges zero transaction costs."""

    def calculate(self, price: float, size: float, liquidity_context: Any | None = None) -> float:
        """Return ``0.0`` for all executions."""

        return 0.0


@dataclass(frozen=True)
class ExecutionModel:
    """Composed execution contract used by backtest runners.

    Attributes:
        slippage: Model used to transform reference next-open price into an
            executed price.
        fees: Model used to compute per-trade transaction fees in quote currency.
    """

    slippage: SlippageModel
    fees: FeeModel


class BpsSlippage:
    """Fixed basis-point slippage model.

    The model worsens the execution price by ``bps`` in the trade direction.
    """

    def __init__(self, bps: float) -> None:
        if bps < 0:
            raise ValueError("Slippage bps must be non-negative.")
        self.bps = bps

    def apply(self, price: float, size: float, liquidity_context: Any | None = None) -> float:
        if size == 0:
            return price
        impact = (self.bps / 10_000.0) * np.sign(size)
        return float(price * (1.0 + impact))


class FixedCommission:
    """Fee model that applies a fixed commission per non-zero trade."""

    def __init__(self, commission_per_trade: float) -> None:
        if commission_per_trade < 0:
            raise ValueError("Commission must be non-negative.")
        self.commission_per_trade = commission_per_trade

    def calculate(self, price: float, size: float, liquidity_context: Any | None = None) -> float:
        if size == 0:
            return 0.0
        return float(self.commission_per_trade)


class ShortBorrowCost:
    """Daily borrow-rate model applied to short notional exposure.

    ``annual_borrow_rate`` is interpreted as decimal annualized carry
    (for example ``0.03`` for 3%/year). The resulting cost is represented as
    return drag per period and is computed as:

    ``max(-position, 0) * annual_borrow_rate / periods_per_year``.
    """

    def __init__(self, annual_borrow_rate: float, periods_per_year: float = 252.0) -> None:
        if annual_borrow_rate < 0:
            raise ValueError("Borrow rate must be non-negative.")
        if periods_per_year <= 0:
            raise ValueError("periods_per_year must be positive.")
        self.annual_borrow_rate = annual_borrow_rate
        self.periods_per_year = periods_per_year

    def calculate(self, position: float | np.ndarray) -> float | np.ndarray:
        short_exposure = np.clip(-np.asarray(position, dtype=float), 0.0, None)
        return short_exposure * (self.annual_borrow_rate / self.periods_per_year)

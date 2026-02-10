"""Execution and transaction cost contracts for backtesting engines.

Execution convention: strategy signals are generated on a bar close and any
resulting order is filled on the next bar open. Slippage and fees are applied
at that next-open execution time.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Protocol


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

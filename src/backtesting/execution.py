"""Execution helpers for backtesting engines."""

from __future__ import annotations

from typing import Any, Protocol


class SlippageModel(Protocol):
    """Interface for execution price adjustments.

    Implementations should return the executed price for a given trade size
    and optional liquidity context.
    """

    def apply(self, price: float, size: float, liquidity_context: Any | None = None) -> float:
        """Return the executed price after slippage."""


class ZeroSlippage:
    """A slippage model that applies zero cost."""

    def apply(self, price: float, size: float, liquidity_context: Any | None = None) -> float:
        return price

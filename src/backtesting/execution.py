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


@dataclass(frozen=True)
class LiquidityContext:
    """Per-bar execution context consumed by advanced cost models."""

    bar_index: int
    asset_index: int
    volume: float
    adv: float
    volatility: float
    spread_bps: float


def _context_value(liquidity_context: Any | None, key: str, default: float) -> float:
    if liquidity_context is None:
        return float(default)
    if isinstance(liquidity_context, LiquidityContext):
        return float(getattr(liquidity_context, key, default))
    if isinstance(liquidity_context, dict):
        return float(liquidity_context.get(key, default))
    return float(default)


class SpreadSlippage:
    """Side-aware spread crossing model using half spread in bps."""

    def __init__(self, spread_bps: float = 2.0) -> None:
        if spread_bps < 0:
            raise ValueError("spread_bps must be non-negative.")
        self.spread_bps = float(spread_bps)

    def apply(self, price: float, size: float, liquidity_context: Any | None = None) -> float:
        if size == 0:
            return float(price)
        spread_bps = _context_value(liquidity_context, "spread_bps", self.spread_bps)
        half_spread = max(spread_bps, 0.0) / 2.0 / 10_000.0
        return float(price * (1.0 + np.sign(size) * half_spread))


class ParticipationImpactSlippage:
    """Volume participation slippage model with nonlinear impact."""

    def __init__(
        self,
        base_bps: float = 0.0,
        impact_coefficient_bps: float = 20.0,
        participation_exponent: float = 1.0,
        max_participation: float = 1.0,
    ) -> None:
        if base_bps < 0 or impact_coefficient_bps < 0:
            raise ValueError("Impact parameters must be non-negative.")
        if participation_exponent <= 0:
            raise ValueError("participation_exponent must be positive.")
        if max_participation <= 0:
            raise ValueError("max_participation must be positive.")
        self.base_bps = float(base_bps)
        self.impact_coefficient_bps = float(impact_coefficient_bps)
        self.participation_exponent = float(participation_exponent)
        self.max_participation = float(max_participation)

    def apply(self, price: float, size: float, liquidity_context: Any | None = None) -> float:
        if size == 0:
            return float(price)
        volume = max(_context_value(liquidity_context, "volume", 0.0), 1e-12)
        adv = max(_context_value(liquidity_context, "adv", volume), 1e-12)
        denom = max(volume, adv)
        participation = min(abs(size) / denom, self.max_participation)
        impact_bps = self.base_bps + self.impact_coefficient_bps * (participation ** self.participation_exponent)
        impact = impact_bps / 10_000.0
        return float(price * (1.0 + np.sign(size) * impact))


class VolatilityScaledSlippage:
    """Slippage model scaled by per-bar volatility regime."""

    def __init__(
        self,
        base_bps: float = 2.0,
        target_volatility: float = 0.01,
        volatility_exponent: float = 1.0,
        min_volatility: float = 1e-6,
    ) -> None:
        if base_bps < 0:
            raise ValueError("base_bps must be non-negative.")
        if target_volatility <= 0 or volatility_exponent <= 0 or min_volatility <= 0:
            raise ValueError("Volatility scaling parameters must be positive.")
        self.base_bps = float(base_bps)
        self.target_volatility = float(target_volatility)
        self.volatility_exponent = float(volatility_exponent)
        self.min_volatility = float(min_volatility)

    def apply(self, price: float, size: float, liquidity_context: Any | None = None) -> float:
        if size == 0:
            return float(price)
        volatility = max(_context_value(liquidity_context, "volatility", self.target_volatility), self.min_volatility)
        scale = (volatility / self.target_volatility) ** self.volatility_exponent
        impact = (self.base_bps * scale) / 10_000.0
        return float(price * (1.0 + np.sign(size) * impact))


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


class AssetClassCarryCost:
    """Borrow/financing carry model with per-asset-class rates."""

    def __init__(
        self,
        *,
        asset_classes: list[str] | tuple[str, ...],
        annual_short_borrow_rates: dict[str, float] | None = None,
        annual_long_financing_rates: dict[str, float] | None = None,
        periods_per_year: float = 252.0,
    ) -> None:
        if periods_per_year <= 0:
            raise ValueError("periods_per_year must be positive.")
        self.asset_classes = tuple(str(asset_class).lower() for asset_class in asset_classes)
        if not self.asset_classes:
            raise ValueError("asset_classes must contain at least one item.")
        self.periods_per_year = float(periods_per_year)
        self.annual_short_borrow_rates = {
            str(key).lower(): float(value)
            for key, value in (annual_short_borrow_rates or {}).items()
        }
        self.annual_long_financing_rates = {
            str(key).lower(): float(value)
            for key, value in (annual_long_financing_rates or {}).items()
        }

        for value in list(self.annual_short_borrow_rates.values()) + list(self.annual_long_financing_rates.values()):
            if value < 0:
                raise ValueError("Carry rates must be non-negative.")

    def calculate(self, position: float | np.ndarray) -> float | np.ndarray:
        positions = np.asarray(position, dtype=float)
        if positions.ndim == 1:
            position_matrix = positions.reshape(-1, 1)
            squeeze = True
        else:
            position_matrix = positions
            squeeze = False

        if position_matrix.shape[1] != len(self.asset_classes):
            raise ValueError("asset_classes length must match number of assets.")

        carry = np.zeros_like(position_matrix, dtype=float)
        for asset_idx, asset_class in enumerate(self.asset_classes):
            short_rate = self.annual_short_borrow_rates.get(asset_class, 0.0) / self.periods_per_year
            long_rate = self.annual_long_financing_rates.get(asset_class, 0.0) / self.periods_per_year
            asset_positions = position_matrix[:, asset_idx]
            carry[:, asset_idx] = np.clip(asset_positions, 0.0, None) * long_rate
            carry[:, asset_idx] += np.clip(-asset_positions, 0.0, None) * short_rate

        if squeeze:
            return carry[:, 0]
        return carry

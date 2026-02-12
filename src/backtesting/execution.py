"""Execution and transaction cost contracts for backtesting engines.

Execution convention: strategy signals are generated on a bar close and any
resulting order is filled on the next bar open. Slippage and fees are applied
at that next-open execution time.
"""

from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path
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


@dataclass(frozen=True)
class ExecutionContext:
    """Per-order execution state used by fill + slippage models."""

    bar_index: int
    asset_index: int
    volume: float
    adv: float
    volatility: float
    spread_bps: float
    order_type: str = "market"
    latency_bars: int = 0
    latency_ms: int = 0
    queue_rank_proxy: float = 0.5
    available_bar_volume: float = 0.0
    max_participation_per_bar: float = 1.0
    realized_participation: float = 0.0
    submit_timestamp: str | None = None
    time_in_force: str = "gtc"
    urgency: str = "normal"
    child_order_id: str | None = None
    event_type: str | None = None
    event_timestamp: str | None = None


def _context_value(liquidity_context: Any | None, key: str, default: float) -> float:
    if liquidity_context is None:
        return float(default)
    if isinstance(liquidity_context, LiquidityContext):
        return float(getattr(liquidity_context, key, default))
    if isinstance(liquidity_context, ExecutionContext):
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
        queue_rank_proxy = np.clip(_context_value(liquidity_context, "queue_rank_proxy", 0.5), 0.0, 1.0)
        queue_penalty = 1.0 + queue_rank_proxy
        half_spread = max(spread_bps, 0.0) / 2.0 / 10_000.0 * queue_penalty
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
        realized_participation = _context_value(liquidity_context, "realized_participation", np.nan)
        if np.isfinite(realized_participation) and realized_participation >= 0.0:
            participation = min(float(realized_participation), self.max_participation)
        else:
            volume = max(_context_value(liquidity_context, "volume", 0.0), 1e-12)
            adv = max(_context_value(liquidity_context, "adv", volume), 1e-12)
            denom = max(volume, adv)
            participation = min(abs(size) / denom, self.max_participation)
        queue_rank_proxy = np.clip(_context_value(liquidity_context, "queue_rank_proxy", 0.5), 0.0, 1.0)
        participation = participation * (1.0 + queue_rank_proxy)
        impact_bps = self.base_bps + self.impact_coefficient_bps * (participation ** self.participation_exponent)
        impact = impact_bps / 10_000.0
        return float(price * (1.0 + np.sign(size) * impact))

    @classmethod
    def from_calibration_buckets(
        cls,
        buckets: list[dict[str, float]],
        *,
        base_bps: float = 0.0,
        participation_exponent: float = 1.0,
        max_participation: float = 1.0,
    ) -> "ParticipationImpactSlippage":
        impact_coefficient = calibrate_impact_coefficient_bps(buckets)
        return cls(
            base_bps=base_bps,
            impact_coefficient_bps=impact_coefficient,
            participation_exponent=participation_exponent,
            max_participation=max_participation,
        )


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
        realized_participation = max(_context_value(liquidity_context, "realized_participation", 0.0), 0.0)
        queue_rank_proxy = np.clip(_context_value(liquidity_context, "queue_rank_proxy", 0.5), 0.0, 1.0)
        scale = (volatility / self.target_volatility) ** self.volatility_exponent
        execution_stress = 1.0 + realized_participation + queue_rank_proxy * 0.25
        impact = (self.base_bps * scale) / 10_000.0
        return float(price * (1.0 + np.sign(size) * impact * execution_stress))


@dataclass(frozen=True)
class FillEvent:
    bar_index: int
    asset_index: int
    requested_size: float
    filled_size: float
    residual_size: float
    participation_rate: float
    available_volume: float
    order_type: str
    latency_bars: int
    latency_ms: int
    queue_rank_proxy: float


class PartialFillModel:
    """Applies per-bar participation caps and carries residual orders forward."""

    def __init__(self, max_participation_per_bar: float = 1.0) -> None:
        if max_participation_per_bar <= 0:
            raise ValueError("max_participation_per_bar must be positive.")
        self.max_participation_per_bar = float(max_participation_per_bar)

    def run(
        self,
        requested_trades: np.ndarray,
        available_volume: np.ndarray,
        *,
        order_type: str = "market",
        latency_bars: int = 0,
        latency_ms: int = 0,
        queue_rank_proxy: float = 0.5,
    ) -> tuple[np.ndarray, np.ndarray, list[FillEvent]]:
        trades = np.asarray(requested_trades, dtype=float)
        volume = np.asarray(available_volume, dtype=float)
        if trades.shape != volume.shape:
            raise ValueError("requested_trades and available_volume must have the same shape.")

        n_periods, n_assets = trades.shape
        executed = np.zeros_like(trades)
        residual = np.zeros_like(trades)
        pending = np.zeros(n_assets, dtype=float)
        fills: list[FillEvent] = []

        latency = max(int(latency_bars), 0)
        queue = float(np.clip(queue_rank_proxy, 0.0, 1.0))

        for idx in range(n_periods):
            source_idx = idx - latency
            if source_idx >= 0:
                pending += trades[source_idx]
            for asset_idx in range(n_assets):
                requested = float(pending[asset_idx])
                bar_volume = max(float(volume[idx, asset_idx]), 0.0)
                cap = bar_volume * self.max_participation_per_bar
                cap = max(cap, 0.0)
                fill_size = 0.0
                if requested != 0.0 and cap > 0.0:
                    fill_size = np.sign(requested) * min(abs(requested), cap)
                    pending[asset_idx] -= fill_size
                executed[idx, asset_idx] = fill_size
                residual[idx, asset_idx] = pending[asset_idx]
                participation = abs(fill_size) / bar_volume if bar_volume > 0 else 0.0
                fills.append(
                    FillEvent(
                        bar_index=idx,
                        asset_index=asset_idx,
                        requested_size=requested,
                        filled_size=float(fill_size),
                        residual_size=float(pending[asset_idx]),
                        participation_rate=float(participation),
                        available_volume=float(bar_volume),
                        order_type=str(order_type),
                        latency_bars=latency,
                        latency_ms=max(int(latency_ms), 0),
                        queue_rank_proxy=queue,
                    )
                )

        return executed, residual, fills


def calibrate_impact_coefficient_bps(buckets: list[dict[str, float]]) -> float:
    """Estimate linear impact coefficient from participation/slippage buckets."""

    numerator = 0.0
    denominator = 0.0
    for bucket in buckets:
        participation = float(bucket.get("participation", bucket.get("participation_rate", 0.0)))
        slippage_bps = float(bucket.get("slippage_bps", bucket.get("impact_bps", 0.0)))
        weight = float(bucket.get("count", bucket.get("trades", 1.0)))
        if participation <= 0.0 or weight <= 0.0:
            continue
        numerator += weight * participation * slippage_bps
        denominator += weight * participation * participation
    if denominator <= 0.0:
        return 0.0
    return float(max(numerator / denominator, 0.0))


def load_impact_calibration_buckets(path: str | Path) -> list[dict[str, float]]:
    raw = json.loads(Path(path).read_text())
    if not isinstance(raw, list):
        raise ValueError("Impact calibration payload must be a JSON array of bucket objects.")
    normalized: list[dict[str, float]] = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        normalized.append({str(key): float(value) for key, value in item.items()})
    return normalized


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


@dataclass(frozen=True)
class CarryContext:
    """Optional context for carry/financing models."""

    dates: np.ndarray | None = None
    symbols: tuple[str, ...] | None = None
    metadata: dict[str, Any] | None = None


class CarryModel:
    """Extended carry model supporting asset-class and time-varying borrow inputs."""

    def __init__(
        self,
        *,
        asset_classes: list[str] | tuple[str, ...],
        expiry_by_asset: list[str | None] | tuple[str | None, ...] | None = None,
        multipliers: list[float] | tuple[float, ...] | None = None,
        borrow_availability_tiers: list[str] | tuple[str, ...] | None = None,
        financing_benchmarks: list[str] | tuple[str, ...] | None = None,
        annual_short_borrow_rates: dict[str, float] | None = None,
        annual_long_financing_rates: dict[str, float] | None = None,
        borrow_rate_series: np.ndarray | None = None,
        borrow_available_flags: np.ndarray | None = None,
        hard_to_borrow_spike_multiplier: float = 2.0,
        annual_futures_roll_rates: dict[str, float] | None = None,
        annual_options_theta_rates: dict[str, float] | None = None,
        periods_per_year: float = 252.0,
    ) -> None:
        if periods_per_year <= 0:
            raise ValueError("periods_per_year must be positive.")
        self.asset_classes = tuple(str(v).lower() for v in asset_classes)
        self.expiry_by_asset = tuple(expiry_by_asset or [None] * len(self.asset_classes))
        self.multipliers = tuple(float(v) for v in (multipliers or [1.0] * len(self.asset_classes)))
        self.borrow_availability_tiers = tuple(
            str(v).lower() for v in (borrow_availability_tiers or ["normal"] * len(self.asset_classes))
        )
        self.financing_benchmarks = tuple(
            str(v).lower() for v in (financing_benchmarks or ["overnight"] * len(self.asset_classes))
        )
        self.periods_per_year = float(periods_per_year)
        self.annual_short_borrow_rates = {str(k).lower(): float(v) for k, v in (annual_short_borrow_rates or {}).items()}
        self.annual_long_financing_rates = {str(k).lower(): float(v) for k, v in (annual_long_financing_rates or {}).items()}
        self.annual_futures_roll_rates = {str(k).lower(): float(v) for k, v in (annual_futures_roll_rates or {}).items()}
        self.annual_options_theta_rates = {str(k).lower(): float(v) for k, v in (annual_options_theta_rates or {}).items()}
        self.borrow_rate_series = None if borrow_rate_series is None else np.asarray(borrow_rate_series, dtype=float)
        self.borrow_available_flags = None if borrow_available_flags is None else np.asarray(borrow_available_flags, dtype=bool)
        self.hard_to_borrow_spike_multiplier = float(max(hard_to_borrow_spike_multiplier, 1.0))

    def calculate(self, position: float | np.ndarray, context: CarryContext | None = None) -> np.ndarray:
        positions = np.asarray(position, dtype=float)
        if positions.ndim == 1:
            matrix = positions.reshape(-1, 1)
        else:
            matrix = positions
        n_periods, n_assets = matrix.shape
        if n_assets != len(self.asset_classes):
            raise ValueError("asset_classes length must match number of assets.")

        carry = np.zeros_like(matrix, dtype=float)
        dynamic_borrow = self._coerce_dynamic(self.borrow_rate_series, (n_periods, n_assets), default=np.nan)
        available_flags = self._coerce_dynamic(self.borrow_available_flags, (n_periods, n_assets), default=True).astype(bool)

        for asset_idx, asset_class in enumerate(self.asset_classes):
            asset_positions = matrix[:, asset_idx]
            multiplier = self.multipliers[asset_idx]

            short_rate = self.annual_short_borrow_rates.get(asset_class, 0.0)
            long_rate = self.annual_long_financing_rates.get(asset_class, 0.0)
            if np.isfinite(dynamic_borrow[:, asset_idx]).any():
                short_rate_series = np.where(
                    np.isfinite(dynamic_borrow[:, asset_idx]),
                    dynamic_borrow[:, asset_idx],
                    short_rate,
                )
            else:
                short_rate_series = np.full(n_periods, short_rate, dtype=float)

            tier = self.borrow_availability_tiers[asset_idx]
            availability_penalty = np.where(available_flags[:, asset_idx], 1.0, self.hard_to_borrow_spike_multiplier)
            if tier in {"hard", "hard_to_borrow", "htb"}:
                availability_penalty = availability_penalty * self.hard_to_borrow_spike_multiplier

            short_component = np.clip(-asset_positions, 0.0, None) * (short_rate_series / self.periods_per_year) * availability_penalty
            long_component = np.clip(asset_positions, 0.0, None) * (long_rate / self.periods_per_year)

            futures_component = np.zeros(n_periods, dtype=float)
            if asset_class in {"futures", "future"}:
                roll_rate = self.annual_futures_roll_rates.get(asset_class, 0.0) / self.periods_per_year
                futures_component = np.abs(asset_positions) * roll_rate

            options_component = np.zeros(n_periods, dtype=float)
            if asset_class in {"option", "options"}:
                theta_rate = self.annual_options_theta_rates.get(asset_class, 0.0) / self.periods_per_year
                options_component = np.abs(asset_positions) * theta_rate

            carry[:, asset_idx] = (short_component + long_component + futures_component + options_component) * multiplier

        return carry

    @staticmethod
    def _coerce_dynamic(values: np.ndarray | None, shape: tuple[int, int], default: float | bool) -> np.ndarray:
        if values is None:
            return np.full(shape, default)
        arr = np.asarray(values)
        if arr.ndim == 0:
            return np.full(shape, arr.item())
        if arr.ndim == 1:
            if arr.shape[0] == shape[0]:
                return np.repeat(arr.reshape(-1, 1), shape[1], axis=1)
            if arr.shape[0] == shape[1]:
                return np.repeat(arr.reshape(1, -1), shape[0], axis=0)
            raise ValueError("Dynamic carry array must match n_periods or n_assets.")
        if arr.shape != shape:
            raise ValueError("Dynamic carry array must match positions shape.")
        return arr

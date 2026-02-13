"""Execution and transaction cost contracts for backtesting engines.

Execution convention: strategy signals are generated on a bar close and any
resulting order is filled on the next bar open. Slippage and fees are applied
at that next-open execution time.
"""

from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path
from typing import Any, Literal, Protocol

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


@dataclass(frozen=True)
class SlippageCalibrationSelection:
    """Resolved slippage parameters selected from dated snapshots."""

    params: dict[str, Any]
    source: str
    effective_date: str | None
    warning_flags: list[str]


def load_slippage_calibration_snapshots(path: str | Path) -> dict[str, Any]:
    """Load snapshot payload emitted by reports calibration jobs."""

    payload = json.loads(Path(path).read_text())
    if not isinstance(payload, dict):
        raise ValueError("Slippage calibration snapshot payload must be a JSON object.")
    snapshots = payload.get("snapshots", [])
    if not isinstance(snapshots, list):
        raise ValueError("slippage calibration snapshots must be a list.")
    payload["snapshots"] = [item for item in snapshots if isinstance(item, dict)]
    if not isinstance(payload.get("default_params", {}), dict):
        payload["default_params"] = {}
    return payload


def select_slippage_calibration_snapshot(
    payload: dict[str, Any],
    *,
    as_of_date: str | None,
    default_params: dict[str, Any] | None = None,
) -> SlippageCalibrationSelection:
    """Resolve the latest stable calibration snapshot not newer than ``as_of_date``."""

    snapshots = payload.get("snapshots", []) if isinstance(payload, dict) else []
    resolved_defaults = {
        **(default_params or {}),
        **(payload.get("default_params", {}) if isinstance(payload.get("default_params", {}), dict) else {}),
    }
    warnings: list[str] = []

    cutoff = np.datetime64(as_of_date or "2262-04-11")
    candidates: list[tuple[np.datetime64, dict[str, Any]]] = []
    for snapshot in snapshots:
        if not isinstance(snapshot, dict):
            continue
        if not bool(snapshot.get("stable", False)):
            continue
        eff = str(snapshot.get("effective_date", "")).strip()
        if not eff:
            continue
        try:
            eff_date = np.datetime64(eff)
        except Exception:
            continue
        if eff_date <= cutoff:
            candidates.append((eff_date, snapshot))

    if candidates:
        candidates.sort(key=lambda item: item[0])
        selected = candidates[-1][1]
        params = selected.get("params", {}) if isinstance(selected.get("params", {}), dict) else {}
        return SlippageCalibrationSelection(
            params={**resolved_defaults, **params},
            source="snapshot",
            effective_date=str(selected.get("effective_date")),
            warning_flags=warnings,
        )

    if snapshots:
        warnings.append("slippage_calibration_snapshot_unavailable_for_date")
    else:
        warnings.append("slippage_calibration_snapshot_missing")
    return SlippageCalibrationSelection(
        params=resolved_defaults,
        source="default_params",
        effective_date=None,
        warning_flags=warnings,
    )


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


class SquareRootImpactSlippage:
    """Square-root participation impact approximation.

    Impact is modeled as ``impact_bps * sqrt(participation_rate)`` and applied in
    the trade direction.
    """

    def __init__(self, impact_bps: float = 15.0, max_participation: float = 1.0) -> None:
        if impact_bps < 0.0:
            raise ValueError("impact_bps must be non-negative.")
        if max_participation <= 0.0:
            raise ValueError("max_participation must be positive.")
        self.impact_bps = float(impact_bps)
        self.max_participation = float(max_participation)

    def apply(self, price: float, size: float, liquidity_context: Any | None = None) -> float:
        if size == 0.0:
            return float(price)
        participation = max(_context_value(liquidity_context, "realized_participation", np.nan), 0.0)
        if not np.isfinite(participation):
            available = max(_context_value(liquidity_context, "available_bar_volume", 0.0), 1e-12)
            participation = min(abs(float(size)) / available, self.max_participation)
        participation = min(float(participation), self.max_participation)
        impact = (self.impact_bps * np.sqrt(participation)) / 10_000.0
        return float(price * (1.0 + np.sign(size) * impact))


class LatencyQueueDriftSlippage:
    """Adverse drift model that penalizes delayed/queued marketable orders."""

    def __init__(
        self,
        drift_bps_per_bar: float = 1.0,
        queue_drift_bps: float = 2.0,
        latency_ms_per_bar: float = 60_000.0,
    ) -> None:
        if drift_bps_per_bar < 0.0 or queue_drift_bps < 0.0:
            raise ValueError("Latency drift parameters must be non-negative.")
        if latency_ms_per_bar <= 0.0:
            raise ValueError("latency_ms_per_bar must be positive.")
        self.drift_bps_per_bar = float(drift_bps_per_bar)
        self.queue_drift_bps = float(queue_drift_bps)
        self.latency_ms_per_bar = float(latency_ms_per_bar)

    def apply(self, price: float, size: float, liquidity_context: Any | None = None) -> float:
        if size == 0.0:
            return float(price)
        latency_bars = max(_context_value(liquidity_context, "latency_bars", 0.0), 0.0)
        latency_ms = max(_context_value(liquidity_context, "latency_ms", 0.0), 0.0)
        queue_rank = float(np.clip(_context_value(liquidity_context, "queue_rank_proxy", 0.5), 0.0, 1.0))
        equivalent_bars = latency_bars + latency_ms / self.latency_ms_per_bar
        drift_bps = self.drift_bps_per_bar * equivalent_bars + self.queue_drift_bps * queue_rank
        return float(price * (1.0 + np.sign(size) * drift_bps / 10_000.0))


class CompositeSlippage:
    """Composable slippage stack combining spread/impact/drift modules."""

    def __init__(self, components: list[SlippageModel]) -> None:
        self.components = list(components)

    def apply(self, price: float, size: float, liquidity_context: Any | None = None) -> float:
        executed = float(price)
        for component in self.components:
            executed = float(component.apply(executed, size, liquidity_context))
        return executed


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


SchedulerStyle = Literal["twap", "vwap", "pov", "arrival_price", "arrival"]


@dataclass(frozen=True)
class ExecutionTrajectory:
    """Child-order trajectory generated from a parent order."""

    scheduler: str
    parent_size: float
    child_sizes: np.ndarray
    schedule_weights: np.ndarray
    horizon_bars: int


def build_child_order_trajectory(
    *,
    parent_size: float,
    horizon_bars: int,
    scheduler: SchedulerStyle = "twap",
    volume_profile: np.ndarray | None = None,
    pov_rate: float = 0.1,
    arrival_urgency: float = 2.0,
) -> ExecutionTrajectory:
    """Convert a parent order into a scheduled child-order trajectory."""

    horizon = max(1, int(horizon_bars))
    total_size = float(parent_size)
    abs_parent = abs(total_size)

    if abs_parent == 0.0:
        zeros = np.zeros(horizon, dtype=float)
        return ExecutionTrajectory(
            scheduler=str(scheduler),
            parent_size=total_size,
            child_sizes=zeros,
            schedule_weights=zeros,
            horizon_bars=horizon,
        )

    if scheduler == "twap":
        weights = np.full(horizon, 1.0 / horizon, dtype=float)
    elif scheduler == "vwap":
        if volume_profile is None:
            weights = np.full(horizon, 1.0 / horizon, dtype=float)
        else:
            volume = np.asarray(volume_profile, dtype=float).reshape(-1)
            if volume.size != horizon:
                raise ValueError("volume_profile length must match horizon_bars.")
            volume = np.clip(volume, 0.0, None)
            denom = float(np.sum(volume))
            weights = np.full(horizon, 1.0 / horizon, dtype=float) if denom <= 0.0 else volume / denom
    elif scheduler in {"arrival_price", "arrival"}:
        urgency = max(float(arrival_urgency), 1e-6)
        x = np.linspace(0.0, 1.0, horizon)
        profile = np.exp(-urgency * x)
        weights = profile / np.sum(profile)
    elif scheduler == "pov":
        if volume_profile is None:
            volume = np.ones(horizon, dtype=float)
        else:
            volume = np.asarray(volume_profile, dtype=float).reshape(-1)
            if volume.size != horizon:
                raise ValueError("volume_profile length must match horizon_bars.")
            volume = np.clip(volume, 0.0, None)
        remaining = abs_parent
        fills = np.zeros(horizon, dtype=float)
        cap_rate = max(float(pov_rate), 0.0)
        for idx in range(horizon):
            take = min(remaining, cap_rate * float(volume[idx]))
            fills[idx] = take
            remaining -= take
        if remaining > 0:
            fills[-1] += remaining
        weights = fills / abs_parent
    else:
        raise ValueError(f"Unsupported scheduler: {scheduler}")

    children = np.sign(total_size) * abs_parent * weights
    return ExecutionTrajectory(
        scheduler=str(scheduler),
        parent_size=total_size,
        child_sizes=children,
        schedule_weights=weights,
        horizon_bars=horizon,
    )


def simulate_alpha_decay_pnl(
    *,
    parent_size: float,
    arrival_price: float,
    scheduler: SchedulerStyle,
    horizon_bars: int,
    alpha_bps: float,
    alpha_half_life_bars: float = 5.0,
    slippage_bps: float = 2.0,
    fee_bps: float = 0.5,
    volume_profile: np.ndarray | None = None,
    pov_rate: float = 0.1,
    arrival_urgency: float = 2.0,
) -> dict[str, float | list[float]]:
    """Simulate realized alpha PnL with alpha-decay and execution-cost drag."""

    trajectory = build_child_order_trajectory(
        parent_size=parent_size,
        horizon_bars=horizon_bars,
        scheduler=scheduler,
        volume_profile=volume_profile,
        pov_rate=pov_rate,
        arrival_urgency=arrival_urgency,
    )
    half_life = max(float(alpha_half_life_bars), 1e-6)
    decay_lambda = np.log(2.0) / half_life
    abs_parent = max(abs(float(parent_size)), 1e-12)

    alpha_pnl = 0.0
    cost_pnl = 0.0
    realized_alpha_path: list[float] = []
    for idx, child in enumerate(trajectory.child_sizes):
        child_abs = abs(float(child))
        fill_fraction = child_abs / abs_parent
        alpha_eff_bps = float(alpha_bps) * float(np.exp(-decay_lambda * idx))
        notional = child_abs * float(arrival_price)
        alpha_component = notional * (alpha_eff_bps / 10_000.0)
        cost_component = notional * ((float(slippage_bps) + float(fee_bps)) / 10_000.0)
        alpha_pnl += alpha_component
        cost_pnl += cost_component
        realized_alpha_path.append(alpha_eff_bps * fill_fraction)

    return {
        "alpha_pnl": float(alpha_pnl),
        "cost_pnl": float(cost_pnl),
        "net_pnl": float(alpha_pnl - cost_pnl),
        "realized_alpha_bps": float(np.sum(realized_alpha_path)),
        "realized_alpha_path_bps": [float(v) for v in realized_alpha_path],
        "horizon_bars": int(horizon_bars),
    }


class PartialFillModel:
    """Applies per-bar participation caps and carries residual orders forward."""

    def __init__(self, max_participation_per_bar: float = 1.0) -> None:
        if max_participation_per_bar <= 0:
            raise ValueError("max_participation_per_bar must be positive.")
        self.max_participation_per_bar = float(max_participation_per_bar)

    def capped_fill_size(self, requested_size: float, available_volume: float, max_participation: float | None = None) -> float:
        """Return capped fill size under per-bar participation constraints."""

        if requested_size == 0.0:
            return 0.0
        cap_ratio = self.max_participation_per_bar if max_participation is None else max(float(max_participation), 0.0)
        cap = max(float(available_volume), 0.0) * cap_ratio
        if cap <= 0.0:
            return 0.0
        return float(np.sign(requested_size) * min(abs(float(requested_size)), cap))

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
                fill_size = 0.0
                if requested != 0.0:
                    fill_size = self.capped_fill_size(requested, bar_volume)
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

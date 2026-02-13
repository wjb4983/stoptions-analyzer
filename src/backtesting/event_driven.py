"""Event-driven backtesting contracts and reference loop.

Execution convention: strategies emit signals/orders after observing a bar close;
those orders are executed on the next bar open with slippage/fees applied at
that open.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Callable, Dict, Iterable, List, Literal, Optional, Protocol

from .execution import ExecutionContext, FeeModel, PartialFillModel, SlippageModel, ZeroFee, ZeroSlippage


@dataclass(frozen=True)
class Order:
    """Intent to trade on the next bar open.

    Attributes:
        symbol: Tradable instrument identifier.
        quantity: Unsigned trade size in units.
        side: ``"buy"`` or ``"sell"``.
        timestamp: Optional signal timestamp (typically bar close time).
    """

    symbol: str
    quantity: float
    side: str
    timestamp: Optional[str] = None


@dataclass(frozen=True)
class Fill:
    """Executed order details.

    Attributes:
        order: Original order intent.
        price: Executed unit price in quote currency.
        fee: Total transaction fee in quote currency for the fill.
        timestamp: Fill timestamp (typically next bar open time).
    """

    order: Order
    price: float
    fee: float = 0.0
    timestamp: Optional[str] = None


@dataclass
class Position:
    """Position state for a single symbol.

    Units are expressed in shares/contracts; ``average_price`` is in quote
    currency per unit.
    """

    symbol: str
    quantity: float = 0.0
    average_price: float = 0.0

    def apply_fill(self, fill: Fill) -> None:
        """Apply a fill with explicit add, reduce, and reversal handling."""

        signed_qty = fill.order.quantity if fill.order.side == "buy" else -fill.order.quantity
        if signed_qty == 0:
            return

        current_qty = self.quantity
        if current_qty == 0:
            self.quantity = signed_qty
            self.average_price = fill.price
            return

        same_side = current_qty * signed_qty > 0
        if same_side:
            # Add-to-same-side: update VWAP entry.
            new_qty = current_qty + signed_qty
            weighted_cost = self.average_price * current_qty + fill.price * signed_qty
            self.quantity = new_qty
            self.average_price = weighted_cost / new_qty
            return

        fill_abs = abs(signed_qty)
        current_abs = abs(current_qty)
        if fill_abs < current_abs:
            # Reduce-without-crossing: realize PnL externally via cash; keep entry basis unchanged.
            self.quantity = current_qty + signed_qty
            return

        if fill_abs == current_abs:
            # Fully close to flat.
            self.quantity = 0.0
            self.average_price = 0.0
            return

        # Cross-through-zero reversal:
        # 1) Closing leg flattens existing position (realized PnL bookkeeping remains separate).
        # 2) Opening residual establishes new opposite-side position at execution price.
        opening_residual = fill_abs - current_abs
        self.quantity = opening_residual if signed_qty > 0 else -opening_residual
        self.average_price = fill.price


@dataclass
class Portfolio:
    """Portfolio cash and per-symbol positions.

    ``cash`` is denominated in quote currency.
    """

    cash: float
    positions: Dict[str, Position] = field(default_factory=dict)

    def apply_fill(self, fill: Fill) -> None:
        """Book a fill into positions and cash, including transaction fees."""

        position = self.positions.setdefault(fill.order.symbol, Position(fill.order.symbol))
        position.apply_fill(fill)
        cash_delta = fill.price * fill.order.quantity
        if fill.order.side == "buy":
            self.cash -= cash_delta
        else:
            self.cash += cash_delta
        self.cash -= fill.fee


@dataclass(frozen=True)
class EventDrivenResult:
    """Result container for event-driven runs.

    Attributes:
        portfolio: Final portfolio state after all bars are processed.
        fills: Chronological list of executions.
    """

    portfolio: Portfolio
    fills: List[Fill]


class Strategy(Protocol):
    """Strategy contract used by the event-driven runner."""

    def on_bar(self, bar: dict[str, Any], portfolio: Portfolio) -> Iterable[Order]:
        """Generate orders from current bar close and current portfolio state."""


class EventDrivenRunner(Protocol):
    """Contract for event-driven runners with next-open execution semantics."""

    def run(self, bars: Iterable[dict[str, Any]], strategy: Strategy) -> EventDrivenResult:
        """Run the event loop and return final portfolio and fills."""


class EventDrivenBacktester:
    """Reference event loop executing signals at next bar open.

    The runner buffers orders emitted at bar ``t`` and executes them at open of
    bar ``t+1``.
    """

    def __init__(
        self,
        initial_cash: float = 0.0,
        slippage_model: SlippageModel | None = None,
        fee_model: FeeModel | None = None,
    ) -> None:
        """Initialize backtester state and execution models.

        Args:
            initial_cash: Starting cash in quote currency.
            slippage_model: Optional slippage model for next-open execution.
            fee_model: Optional fee model charged per fill.
        """

        self.portfolio = Portfolio(cash=initial_cash)
        self.slippage_model = slippage_model or ZeroSlippage()
        self.fee_model = fee_model or ZeroFee()
        self.fills: List[Fill] = []

    def run(self, bars: Iterable[dict[str, Any]], strategy: Strategy) -> EventDrivenResult:
        """Process bars and execute strategy orders at next bar open.

        Assumptions:
            * Strategy reads each completed bar and emits orders after close.
            * Pending orders are filled at the next bar open.
            * Slippage and fees are applied at fill time.

        Args:
            bars: Iterable of OHLCV-like bar dictionaries containing ``open``.
            strategy: Strategy implementing ``on_bar``.

        Returns:
            ``EventDrivenResult`` containing final portfolio and fill ledger.
        """

        pending_orders: List[Order] = []
        for bar in bars:
            if pending_orders:
                fills = self._execute_orders_at_open(pending_orders, bar)
                for fill in fills:
                    self.portfolio.apply_fill(fill)
                self.fills.extend(fills)
                pending_orders = []
            new_orders = list(strategy.on_bar(bar, self.portfolio))
            pending_orders.extend(new_orders)
        return EventDrivenResult(portfolio=self.portfolio, fills=list(self.fills))

    def _execute_orders_at_open(self, orders: Iterable[Order], bar: dict[str, Any]) -> List[Fill]:
        open_price = float(bar["open"])
        timestamp = bar.get("timestamp")
        fills: List[Fill] = []
        for order in orders:
            signed_size = order.quantity if order.side == "buy" else -order.quantity
            exec_price = self.slippage_model.apply(open_price, signed_size, bar)
            fee = self.fee_model.calculate(exec_price, signed_size, bar)
            fills.append(Fill(order=order, price=exec_price, fee=fee, timestamp=timestamp))
        return fills


OrderState = Literal["new", "working", "partial", "filled", "canceled", "expired"]
OrderEventType = Literal["submit", "amend", "cancel", "fill", "expire"]


@dataclass(frozen=True)
class OrderLifecycleEvent:
    """Deterministic order lifecycle event emitted by the execution adapter."""

    event_type: OrderEventType
    order_id: str
    state: OrderState
    side: str
    quantity: float
    filled_quantity: float
    remaining_quantity: float
    bar_index: int
    symbol: str
    timestamp: str | None = None
    parent_order_id: str | None = None
    price: float | None = None
    fee: float | None = None
    submit_timestamp: str | None = None
    time_in_force: str = "gtc"
    urgency: str = "normal"
    child_order_id: str | None = None


@dataclass
class ManagedOrder:
    order_id: str
    symbol: str
    side: str
    quantity: float
    filled_quantity: float = 0.0
    state: OrderState = "new"
    submit_timestamp: str | None = None
    time_in_force: str = "gtc"
    urgency: str = "normal"
    child_order_id: str | None = None

    @property
    def remaining_quantity(self) -> float:
        return max(self.quantity - self.filled_quantity, 0.0)


class OrderLifecycleBook:
    """Tracks explicit order states and emits reproducible lifecycle events."""

    def __init__(self) -> None:
        self.orders: dict[str, ManagedOrder] = {}
        self.events: list[OrderLifecycleEvent] = []

    def _emit(self, event: OrderLifecycleEvent) -> None:
        self.events.append(event)

    def submit(
        self,
        *,
        order_id: str,
        symbol: str,
        side: str,
        quantity: float,
        bar_index: int,
        context: ExecutionContext,
        parent_order_id: str | None = None,
    ) -> None:
        order = ManagedOrder(
            order_id=order_id,
            symbol=symbol,
            side=side,
            quantity=float(quantity),
            state="working",
            submit_timestamp=context.submit_timestamp,
            time_in_force=context.time_in_force,
            urgency=context.urgency,
            child_order_id=context.child_order_id,
        )
        self.orders[order_id] = order
        self._emit(
            OrderLifecycleEvent(
                event_type="submit",
                order_id=order_id,
                state=order.state,
                side=side,
                quantity=order.quantity,
                filled_quantity=order.filled_quantity,
                remaining_quantity=order.remaining_quantity,
                bar_index=int(bar_index),
                symbol=symbol,
                timestamp=context.event_timestamp,
                parent_order_id=parent_order_id,
                submit_timestamp=context.submit_timestamp,
                time_in_force=context.time_in_force,
                urgency=context.urgency,
                child_order_id=context.child_order_id,
            )
        )

    def amend(self, *, order_id: str, new_quantity: float, bar_index: int, context: ExecutionContext) -> None:
        order = self.orders[order_id]
        order.quantity = max(float(new_quantity), order.filled_quantity)
        order.child_order_id = context.child_order_id
        order.state = "working" if order.filled_quantity == 0 else "partial"
        self._emit(
            OrderLifecycleEvent(
                event_type="amend",
                order_id=order_id,
                state=order.state,
                side=order.side,
                quantity=order.quantity,
                filled_quantity=order.filled_quantity,
                remaining_quantity=order.remaining_quantity,
                bar_index=int(bar_index),
                symbol=order.symbol,
                timestamp=context.event_timestamp,
                submit_timestamp=order.submit_timestamp,
                time_in_force=context.time_in_force,
                urgency=context.urgency,
                child_order_id=context.child_order_id,
            )
        )

    def cancel(self, *, order_id: str, bar_index: int, context: ExecutionContext) -> None:
        order = self.orders[order_id]
        order.state = "canceled"
        self._emit(
            OrderLifecycleEvent(
                event_type="cancel",
                order_id=order_id,
                state=order.state,
                side=order.side,
                quantity=order.quantity,
                filled_quantity=order.filled_quantity,
                remaining_quantity=order.remaining_quantity,
                bar_index=int(bar_index),
                symbol=order.symbol,
                timestamp=context.event_timestamp,
                submit_timestamp=order.submit_timestamp,
                time_in_force=context.time_in_force,
                urgency=context.urgency,
                child_order_id=context.child_order_id,
            )
        )

    def expire(self, *, order_id: str, bar_index: int, context: ExecutionContext) -> None:
        order = self.orders[order_id]
        order.state = "expired"
        self._emit(
            OrderLifecycleEvent(
                event_type="expire",
                order_id=order_id,
                state=order.state,
                side=order.side,
                quantity=order.quantity,
                filled_quantity=order.filled_quantity,
                remaining_quantity=order.remaining_quantity,
                bar_index=int(bar_index),
                symbol=order.symbol,
                timestamp=context.event_timestamp,
                submit_timestamp=order.submit_timestamp,
                time_in_force=context.time_in_force,
                urgency=context.urgency,
                child_order_id=context.child_order_id,
            )
        )

    def fill(
        self,
        *,
        order_id: str,
        fill_quantity: float,
        bar_index: int,
        context: ExecutionContext,
        price: float,
        fee: float = 0.0,
    ) -> None:
        order = self.orders[order_id]
        if fill_quantity <= 0:
            return
        order.filled_quantity = min(order.quantity, order.filled_quantity + float(fill_quantity))
        order.state = "filled" if order.remaining_quantity <= 1e-12 else "partial"
        self._emit(
            OrderLifecycleEvent(
                event_type="fill",
                order_id=order_id,
                state=order.state,
                side=order.side,
                quantity=order.quantity,
                filled_quantity=order.filled_quantity,
                remaining_quantity=order.remaining_quantity,
                bar_index=int(bar_index),
                symbol=order.symbol,
                timestamp=context.event_timestamp,
                price=float(price),
                fee=float(fee),
                submit_timestamp=order.submit_timestamp,
                time_in_force=context.time_in_force,
                urgency=context.urgency,
                child_order_id=context.child_order_id,
            )
        )


class VectorizedExecutionAdapter:
    """Bridges vectorized requested trades into explicit order lifecycle events."""

    def __init__(self, max_participation_per_bar: float) -> None:
        self.fill_model = PartialFillModel(max_participation_per_bar=max_participation_per_bar)

    def execute(
        self,
        *,
        requested_trades: Any,
        prices: Any,
        available_volume: Any,
        queue_rank_proxy: Any,
        max_participation_per_bar: Any | Callable[[int, int], float],
        order_type: str,
        latency_bars: int,
        latency_ms: int,
        symbols: list[str],
        timestamps: list[str | None],
        time_in_force: str = "gtc",
        urgency: str = "normal",
    ) -> tuple[Any, Any, list[Any], list[OrderLifecycleEvent]]:
        participation_schedule = max_participation_per_bar

        def _resolve_participation(bar_index: int, asset_index: int) -> float:
            if callable(participation_schedule):
                value = participation_schedule(bar_index, asset_index)
            else:
                value = participation_schedule[bar_index, asset_index]
            return max(float(value), 0.0)

        trades, residual, fill_events = self.fill_model.run(
            requested_trades,
            available_volume,
            order_type=order_type,
            latency_bars=latency_bars,
            latency_ms=latency_ms,
            queue_rank_proxy=float(queue_rank_proxy[0, 0]),
            max_participation=_resolve_participation,
        )

        lifecycle = OrderLifecycleBook()
        next_order_id = 0
        active: dict[int, str] = {}

        def _ctx(bar_index: int, asset_index: int, *, event_type: str, child: str | None = None) -> ExecutionContext:
            return ExecutionContext(
                bar_index=bar_index,
                asset_index=asset_index,
                volume=float(available_volume[bar_index, asset_index]),
                adv=float(available_volume[bar_index, asset_index]),
                volatility=0.0,
                spread_bps=0.0,
                order_type=order_type,
                latency_bars=latency_bars,
                latency_ms=latency_ms,
                queue_rank_proxy=float(queue_rank_proxy[bar_index, asset_index]),
                available_bar_volume=float(available_volume[bar_index, asset_index]),
                max_participation_per_bar=_resolve_participation(bar_index, asset_index),
                realized_participation=(
                    abs(float(trades[bar_index, asset_index]))
                    / max(float(available_volume[bar_index, asset_index]), 1e-12)
                ),
                submit_timestamp=timestamps[bar_index],
                time_in_force=time_in_force,
                urgency=urgency,
                child_order_id=child,
                event_type=event_type,
                event_timestamp=timestamps[bar_index],
            )

        for bar_idx in range(trades.shape[0]):
            for asset_idx in range(trades.shape[1]):
                req = float(requested_trades[bar_idx, asset_idx])
                oid = active.get(asset_idx)
                if req != 0.0:
                    side = "buy" if req > 0 else "sell"
                    if oid is None:
                        oid = f"ord-{next_order_id}"
                        next_order_id += 1
                        lifecycle.submit(
                            order_id=oid,
                            symbol=symbols[asset_idx],
                            side=side,
                            quantity=abs(req),
                            bar_index=bar_idx,
                            context=_ctx(bar_idx, asset_idx, event_type="submit"),
                        )
                        active[asset_idx] = oid
                    else:
                        existing = lifecycle.orders[oid]
                        if (existing.side == "buy" and req < 0) or (existing.side == "sell" and req > 0):
                            lifecycle.cancel(order_id=oid, bar_index=bar_idx, context=_ctx(bar_idx, asset_idx, event_type="cancel"))
                            replacement = f"ord-{next_order_id}"
                            next_order_id += 1
                            lifecycle.submit(
                                order_id=replacement,
                                symbol=symbols[asset_idx],
                                side=side,
                                quantity=abs(req),
                                bar_index=bar_idx,
                                context=_ctx(bar_idx, asset_idx, event_type="submit", child=replacement),
                                parent_order_id=oid,
                            )
                            oid = replacement
                            active[asset_idx] = oid
                        else:
                            lifecycle.amend(
                                order_id=oid,
                                new_quantity=lifecycle.orders[oid].quantity + abs(req),
                                bar_index=bar_idx,
                                context=_ctx(bar_idx, asset_idx, event_type="amend", child=oid),
                            )

                filled = float(trades[bar_idx, asset_idx])
                if filled != 0.0:
                    oid = active.get(asset_idx)
                    if oid is None:
                        side = "buy" if filled > 0 else "sell"
                        oid = f"ord-{next_order_id}"
                        next_order_id += 1
                        lifecycle.submit(
                            order_id=oid,
                            symbol=symbols[asset_idx],
                            side=side,
                            quantity=abs(filled),
                            bar_index=bar_idx,
                            context=_ctx(bar_idx, asset_idx, event_type="submit"),
                        )
                        active[asset_idx] = oid
                    lifecycle.fill(
                        order_id=oid,
                        fill_quantity=abs(filled),
                        bar_index=bar_idx,
                        context=_ctx(bar_idx, asset_idx, event_type="fill", child=oid),
                        price=float(prices[bar_idx, asset_idx]),
                    )
                    if lifecycle.orders[oid].state == "filled":
                        active.pop(asset_idx, None)

        return trades, residual, fill_events, lifecycle.events


def replay_lifecycle(events: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Deterministically replay lifecycle event logs and return normalized records."""

    normalized = [OrderLifecycleEvent(**event) for event in events]
    normalized.sort(key=lambda evt: (evt.bar_index, evt.order_id, evt.event_type))
    return [evt.__dict__ for evt in normalized]

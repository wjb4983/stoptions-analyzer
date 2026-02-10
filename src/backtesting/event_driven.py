"""Skeletal event-driven backtesting primitives."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, Iterable, List, Optional, Protocol

from .execution import SlippageModel, ZeroSlippage


@dataclass(frozen=True)
class Order:
    """Represents an intent to trade a quantity at the next bar open."""

    symbol: str
    quantity: float
    side: str  # "buy" or "sell"
    timestamp: Optional[str] = None


@dataclass(frozen=True)
class Fill:
    """Executed order details."""

    order: Order
    price: float
    timestamp: Optional[str] = None


@dataclass
class Position:
    """Position state for a single symbol."""

    symbol: str
    quantity: float = 0.0
    average_price: float = 0.0

    def apply_fill(self, fill: Fill) -> None:
        signed_qty = fill.order.quantity if fill.order.side == "buy" else -fill.order.quantity
        new_qty = self.quantity + signed_qty
        if new_qty == 0:
            self.quantity = 0.0
            self.average_price = 0.0
            return
        if self.quantity == 0:
            self.average_price = fill.price
        else:
            weighted_cost = self.average_price * self.quantity + fill.price * signed_qty
            self.average_price = weighted_cost / new_qty
        self.quantity = new_qty


@dataclass
class Portfolio:
    """Minimal portfolio with cash and positions."""

    cash: float
    positions: Dict[str, Position] = field(default_factory=dict)

    def apply_fill(self, fill: Fill) -> None:
        position = self.positions.setdefault(fill.order.symbol, Position(fill.order.symbol))
        position.apply_fill(fill)
        cash_delta = fill.price * fill.order.quantity
        if fill.order.side == "buy":
            self.cash -= cash_delta
        else:
            self.cash += cash_delta


class Strategy(Protocol):
    """Strategy interface used by the event-driven loop."""

    def on_bar(self, bar: dict, portfolio: Portfolio) -> Iterable[Order]:
        ...


class EventDrivenBacktester:
    """Simple event loop that processes bars and executes orders at next open."""

    def __init__(self, initial_cash: float = 0.0, slippage_model: SlippageModel | None = None) -> None:
        self.portfolio = Portfolio(cash=initial_cash)
        self.slippage_model = slippage_model or ZeroSlippage()
        self.fills: List[Fill] = []

    def run(self, bars: Iterable[dict], strategy: Strategy) -> Portfolio:
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
        return self.portfolio

    def _execute_orders_at_open(self, orders: Iterable[Order], bar: dict) -> List[Fill]:
        open_price = float(bar["open"])
        timestamp = bar.get("timestamp")
        fills: List[Fill] = []
        for order in orders:
            signed_size = order.quantity if order.side == "buy" else -order.quantity
            exec_price = self.slippage_model.apply(open_price, signed_size, bar)
            fills.append(Fill(order=order, price=exec_price, timestamp=timestamp))
        return fills

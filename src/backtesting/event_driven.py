"""Event-driven backtesting contracts and reference loop.

Execution convention: strategies emit signals/orders after observing a bar close;
those orders are executed on the next bar open with slippage/fees applied at
that open.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, Iterable, List, Optional, Protocol

from .execution import FeeModel, SlippageModel, ZeroFee, ZeroSlippage


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
        """Apply a fill to update quantity and volume-weighted average price."""

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

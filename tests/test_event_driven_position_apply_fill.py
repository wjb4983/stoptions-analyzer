from __future__ import annotations

from src.backtesting.event_driven import Fill, Order, Position


def _fill(*, side: str, quantity: float, price: float) -> Fill:
    return Fill(order=Order(symbol="ABC", quantity=quantity, side=side), price=price)


def test_position_apply_fill_long_to_flat_resets_average_price() -> None:
    position = Position(symbol="ABC", quantity=10.0, average_price=100.0)

    position.apply_fill(_fill(side="sell", quantity=10.0, price=110.0))

    assert position.quantity == 0.0
    assert position.average_price == 0.0


def test_position_apply_fill_long_to_short_in_single_sell_sets_new_entry_at_fill_price() -> None:
    position = Position(symbol="ABC", quantity=10.0, average_price=100.0)

    position.apply_fill(_fill(side="sell", quantity=15.0, price=110.0))

    assert position.quantity == -5.0
    assert position.average_price == 110.0


def test_position_apply_fill_short_to_long_in_single_buy_sets_new_entry_at_fill_price() -> None:
    position = Position(symbol="ABC", quantity=-8.0, average_price=105.0)

    position.apply_fill(_fill(side="buy", quantity=10.0, price=95.0))

    assert position.quantity == 2.0
    assert position.average_price == 95.0


def test_position_apply_fill_partial_reduction_without_crossing_keeps_entry_basis() -> None:
    long_position = Position(symbol="ABC", quantity=10.0, average_price=100.0)
    short_position = Position(symbol="ABC", quantity=-10.0, average_price=120.0)

    long_position.apply_fill(_fill(side="sell", quantity=4.0, price=130.0))
    short_position.apply_fill(_fill(side="buy", quantity=3.0, price=90.0))

    assert long_position.quantity == 6.0
    assert long_position.average_price == 100.0
    assert short_position.quantity == -7.0
    assert short_position.average_price == 120.0

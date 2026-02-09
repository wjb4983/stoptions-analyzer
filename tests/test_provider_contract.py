from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

from src.data_access.actions_schema import validate_actions_frame
from src.data_access.bars_schema import validate_bars_frame
from src.data_access.provider_base import DataProvider


class FakeFrame:
    def __init__(self, rows: list[dict[str, Any]]):
        self.rows = rows
        self.columns = list(rows[0].keys()) if rows else []


class FakeProvider:
    def list_symbols(self) -> list[str]:
        return ["AAPL", "MSFT"]

    def get_bars(
        self,
        symbols: list[str],
        start: datetime | str,
        end: datetime | str,
        timeframe: str = "1m",
    ) -> FakeFrame:
        return FakeFrame(
            [
                {
                    "symbol": symbols[0],
                    "timestamp_utc": datetime(2024, 1, 2, tzinfo=timezone.utc),
                    "open": 100.0,
                    "high": 101.0,
                    "low": 99.5,
                    "close": 100.5,
                    "volume": 1000.0,
                    "trades": 10,
                }
            ]
        )

    def get_corporate_actions(
        self,
        symbols: list[str],
        start: datetime | str,
        end: datetime | str,
    ) -> FakeFrame:
        return FakeFrame(
            [
                {
                    "symbol": symbols[0],
                    "action_type": "dividend",
                    "action_date": datetime(2024, 1, 15, tzinfo=timezone.utc),
                    "value": 0.5,
                }
            ]
        )


def test_provider_contract() -> None:
    provider = FakeProvider()
    assert isinstance(provider, DataProvider)

    symbols = provider.list_symbols()
    assert isinstance(symbols, list)
    assert all(isinstance(symbol, str) for symbol in symbols)

    bars = provider.get_bars(symbols, "2024-01-01", "2024-01-31")
    assert hasattr(bars, "columns")
    validate_bars_frame(bars)

    actions = provider.get_corporate_actions(symbols, "2024-01-01", "2024-01-31")
    assert hasattr(actions, "columns")
    validate_actions_frame(actions)

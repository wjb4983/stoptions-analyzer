"""Base contract for market data providers."""

from __future__ import annotations

from datetime import datetime
from typing import Any, Protocol, Sequence, runtime_checkable


@runtime_checkable
class DataProvider(Protocol):
    """Protocol for data provider integrations.

    Implementations should return DataFrame-like objects that satisfy the
    canonical schemas in ``bars_schema`` and ``actions_schema``.
    """

    def list_symbols(self) -> list[str]:
        """Return a list of supported symbols."""

    def get_bars(
        self,
        symbols: Sequence[str],
        start: datetime | str,
        end: datetime | str,
        timeframe: str = "1m",
    ) -> Any:
        """Fetch historical bars for one or more symbols."""

    def get_corporate_actions(
        self,
        symbols: Sequence[str],
        start: datetime | str,
        end: datetime | str,
    ) -> Any:
        """Fetch corporate actions for one or more symbols."""

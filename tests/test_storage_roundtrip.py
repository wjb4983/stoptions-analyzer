from __future__ import annotations

from datetime import date

import pytest

from src.data_access.storage import LocalBarStore
from tests.fixtures_datasets import synthetic_vendor_bars_dataset


def test_storage_roundtrip_is_idempotent(tmp_path) -> None:
    store = LocalBarStore(cache_root=tmp_path, timeframe="1m")
    bars = synthetic_vendor_bars_dataset()

    store.store_bars("AAPL", bars, start_date=date(2024, 1, 2), end_date=date(2024, 1, 2))
    loaded_once = store.load_bars("AAPL", "2024-01-02T14:30:00+00:00", "2024-01-02T14:32:00+00:00")
    loaded_twice = store.load_bars("AAPL", "2024-01-02T14:30:00+00:00", "2024-01-02T14:32:00+00:00")

    assert loaded_once == loaded_twice
    assert [row["close"] for row in loaded_once] == [100.5, 101.0, 101.5]


def test_storage_uses_safe_symbol_paths_and_serialization_edges(tmp_path) -> None:
    store = LocalBarStore(cache_root=tmp_path, timeframe="1m")
    bars = synthetic_vendor_bars_dataset()

    index = store.store_bars("BRK/B", bars, start_date=date(2024, 1, 2), end_date=date(2024, 1, 2))

    assert index["ticker"] == "BRK/B"
    assert store.list_years("BRK/B") == [2024]
    assert store.load_index("BRK/B") is not None


def test_storage_rejects_invalid_payloads(tmp_path) -> None:
    store = LocalBarStore(cache_root=tmp_path, timeframe="1m")
    bad_bars = [{"t": 1704205800000, "o": 100.0, "h": 101.0, "l": 99.0, "c": 100.0, "v": 1000.0}]

    store.store_bars("AAPL", bad_bars, start_date=date(2024, 1, 2), end_date=date(2024, 1, 2))

    with pytest.raises(ValueError, match="trades must be an integer"):
        store.load_bars("AAPL", "2024-01-02T14:30:00+00:00", "2024-01-02T14:31:00+00:00")

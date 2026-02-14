from __future__ import annotations

from copy import deepcopy

import pytest

from src.data_access.option_loader import load_option_records
from tests.fixtures_datasets import synthetic_option_snapshots_dataset


class _StubClient:
    def __init__(self, payload: list[dict]):
        self.payload = payload

    def fetch_option_snapshots(self, ticker: str) -> list[dict]:
        return deepcopy(self.payload)


def test_option_loader_contract_columns_and_dtypes(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("src.data_access.option_loader.load_cached_market_data", lambda ticker: None)
    monkeypatch.setattr("src.data_access.option_loader.save_cached_market_data", lambda ticker, payload: None)

    records = load_option_records(_StubClient(synthetic_option_snapshots_dataset()), "AAPL")

    assert records
    expected_fields = {
        "ticker",
        "expiration_date",
        "contract_type",
        "strike_price",
        "implied_volatility",
        "volume",
        "open_interest",
        "day_close",
        "bid",
        "ask",
        "last",
        "greeks",
    }
    for record in records:
        assert expected_fields.issubset(record.keys())
        assert isinstance(record["ticker"], str)
        assert isinstance(record["expiration_date"], str)
        assert isinstance(record["contract_type"], str)
        assert isinstance(record["strike_price"], float)
        assert isinstance(record["greeks"], dict)


def test_option_loader_rejects_duplicate_symbol_contracts(monkeypatch: pytest.MonkeyPatch) -> None:
    duplicate_payload = synthetic_option_snapshots_dataset()
    duplicate_payload.append(deepcopy(duplicate_payload[0]))
    monkeypatch.setattr("src.data_access.option_loader.load_cached_market_data", lambda ticker: None)
    monkeypatch.setattr("src.data_access.option_loader.save_cached_market_data", lambda ticker, payload: None)

    with pytest.raises(ValueError, match="Duplicate option contracts"):
        load_option_records(_StubClient(duplicate_payload), "AAPL")


def test_option_loader_returns_stable_ordering(monkeypatch: pytest.MonkeyPatch) -> None:
    payload = list(reversed(synthetic_option_snapshots_dataset()))
    monkeypatch.setattr("src.data_access.option_loader.load_cached_market_data", lambda ticker: None)
    monkeypatch.setattr("src.data_access.option_loader.save_cached_market_data", lambda ticker, payload: None)

    records = load_option_records(_StubClient(payload), "AAPL")

    identities = [
        (r["ticker"], r["expiration_date"], r["contract_type"], r["strike_price"])
        for r in records
    ]
    assert identities == sorted(identities)

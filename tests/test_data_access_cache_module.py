from __future__ import annotations

from pathlib import Path

import pytest

import src.config as config
from src.data_access import cache


def _patch_data_dir(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setattr(config, "DATA_DIR", tmp_path)
    monkeypatch.setattr(cache, "DATA_DIR", tmp_path)


@pytest.mark.parametrize(
    ("ticker", "expected"),
    [
        ("brk.b", "BRK_B"),
        ("btc/usd", "BTC_USD"),
        ("  spy  ", "__SPY__"),
        ("aapl", "AAPL"),
        ("X-y.z / q", "X_Y_Z___Q"),
    ],
)
def test_safe_ticker_name_normalizes_common_punctuation(ticker: str, expected: str) -> None:
    assert cache._safe_ticker_name(ticker) == expected


def test_cache_path_resolves_under_monkeypatched_data_dir(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_data_dir(monkeypatch, tmp_path)

    path = cache._cache_path("brk.b")

    assert path == tmp_path / "BRK_B.json"


def test_load_cached_market_data_returns_none_when_file_missing(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_data_dir(monkeypatch, tmp_path)

    assert cache.load_cached_market_data("aapl") is None


def test_load_cached_market_data_returns_none_for_corrupt_json(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_data_dir(monkeypatch, tmp_path)
    bad_path = tmp_path / "AAPL.json"
    bad_path.parent.mkdir(parents=True, exist_ok=True)
    bad_path.write_text("{not json", encoding="utf-8")

    assert cache.load_cached_market_data("aapl") is None


def test_save_and_load_cached_market_data_round_trip(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_data_dir(monkeypatch, tmp_path)
    payload = {
        "ticker": "AAPL",
        "prices": [{"date": "2024-01-02", "close": 190.2}],
        "meta": {"provider": "test", "complete": True},
    }

    cache.save_cached_market_data("aapl", payload)

    assert cache.load_cached_market_data("aapl") == payload


def test_save_cached_market_data_last_write_wins(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_data_dir(monkeypatch, tmp_path)
    first = {"ticker": "AAPL", "revision": 1, "value": 10}
    second = {"ticker": "AAPL", "revision": 2, "value": 20}

    cache.save_cached_market_data("aapl", first)
    cache.save_cached_market_data("aapl", second)

    assert cache.load_cached_market_data("aapl") == second

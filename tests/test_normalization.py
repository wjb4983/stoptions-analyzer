from __future__ import annotations

from datetime import datetime, timedelta, timezone
import math

import pytest

from src.data_access.normalization import NormalizationConfig, normalize_bars


def test_timezone_conversion_with_vendor_timezone() -> None:
    pytest.importorskip("zoneinfo")

    bars = [
        {
            "symbol": "AAPL",
            "timestamp": datetime(2024, 1, 2, 9, 30),
            "open": 100.0,
            "high": 101.0,
            "low": 99.0,
            "close": 100.5,
            "volume": 1000.0,
            "trades": 10,
        }
    ]
    config = NormalizationConfig(vendor_timezone="America/New_York")
    normalized = normalize_bars(bars, config)

    assert normalized[0]["timestamp_utc"] == datetime(2024, 1, 2, 14, 30, tzinfo=timezone.utc)


def test_dedupe_rules_prefer_first_and_max_volume() -> None:
    timestamp = datetime(2024, 1, 2, 14, 30, tzinfo=timezone.utc)
    bars = [
        {
            "symbol": "AAPL",
            "timestamp_utc": timestamp,
            "open": 100.0,
            "high": 101.0,
            "low": 99.0,
            "close": 100.5,
            "volume": 100.0,
            "trades": 10,
        },
        {
            "symbol": "AAPL",
            "timestamp_utc": timestamp,
            "open": 105.0,
            "high": 106.0,
            "low": 104.0,
            "close": 105.5,
            "volume": 200.0,
            "trades": 12,
        },
    ]

    prefer_first = normalize_bars(bars, NormalizationConfig(conflict_resolution="prefer_first"))
    assert prefer_first[0]["open"] == 100.0

    max_volume = normalize_bars(bars, NormalizationConfig(conflict_resolution="max_volume"))
    assert max_volume[0]["open"] == 105.0


@pytest.mark.parametrize(
    ("policy", "expect_ffill"),
    [
        ("ffill", True),
        ("nan", False),
    ],
)
def test_missing_bar_policy_synthesizes(policy: str, expect_ffill: bool) -> None:
    start = datetime(2024, 1, 2, 14, 30, tzinfo=timezone.utc)
    bars = [
        {
            "symbol": "AAPL",
            "timestamp_utc": start,
            "open": 100.0,
            "high": 101.0,
            "low": 99.0,
            "close": 100.5,
            "volume": 100.0,
            "trades": 10,
            "vwap": 100.2,
        },
        {
            "symbol": "AAPL",
            "timestamp_utc": start + timedelta(minutes=2),
            "open": 101.0,
            "high": 102.0,
            "low": 100.0,
            "close": 101.5,
            "volume": 150.0,
            "trades": 12,
            "vwap": 101.2,
        },
    ]
    config = NormalizationConfig(
        expected_interval=timedelta(minutes=1),
        missing_bar_policy=policy,
    )
    normalized = normalize_bars(bars, config)

    assert len(normalized) == 3
    missing = normalized[1]
    assert missing["timestamp_utc"] == start + timedelta(minutes=1)
    if expect_ffill:
        assert missing["open"] == 100.0
        assert missing["vwap"] == 100.2
    else:
        assert math.isnan(missing["open"])
        assert math.isnan(missing["vwap"])

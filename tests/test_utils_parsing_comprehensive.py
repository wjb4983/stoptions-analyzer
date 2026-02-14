from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import numpy as np
import pytest

from src.utils import parsing


class TestDateAndNumberParsers:
    @pytest.mark.parametrize(
        ("value", "expected"),
        [
            (" 3.14 ", 3.14),
            ("1e2", 100.0),
            ("", None),
            ("  ", None),
            ("abc", None),
        ],
    )
    def test_parse_float(self, value: str, expected: float | None) -> None:
        assert parsing.parse_float(value) == expected

    @pytest.mark.parametrize(
        ("value", "expected"),
        [
            (" 42 ", 42),
            ("-10", -10),
            ("", None),
            ("  ", None),
            ("3.14", None),
        ],
    )
    def test_parse_int(self, value: str, expected: int | None) -> None:
        assert parsing.parse_int(value) == expected

    @pytest.mark.parametrize(
        ("value", "expected"),
        [
            ("2024-01-31", date(2024, 1, 31)),
            (" 2024-02-29 ", date(2024, 2, 29)),
            ("", None),
            ("2024/01/31", None),
            ("2024-02-30", None),
        ],
    )
    def test_parse_date(self, value: str, expected: date | None) -> None:
        assert parsing.parse_date(value) == expected

    @pytest.mark.parametrize(
        ("value", "expected"),
        [
            ("0.25", 0.25),
            ("25", 0.25),
            ("250", 1.0),
            ("-2", 0.0),
            ("", None),
            ("not-a-number", None),
        ],
    )
    def test_normalize_likelihood_threshold(self, value: str, expected: float | None) -> None:
        assert parsing.normalize_likelihood_threshold(value) == expected


class TestOptionHelpers:
    def test_normalize_option_records_handles_fallbacks_and_non_dict_greeks(self) -> None:
        records = [
            {
                "details": {
                    "ticker": "AAPL240119C00150000",
                    "expiration_date": "2024-01-19",
                    "contract_type": "call",
                    "strike_price": 150,
                    "open_interest": 111,
                },
                "day": {"volume": 22, "close": 1.1},
                "last_quote": {"bp": 1.0, "ap": 1.2},
                "last_trade": {"p": 1.15},
                "greeks": "not-a-dict",
                "implied_vol": 0.31,
            },
            {
                "ticker": "MSFT240119P00300000",
                "greeks": {"delta": "0.4", "gamma": 0.02, "iv": 0.28},
                "volume": "15",
                "open_interest": 30,
                "bid": 2.1,
                "ask": 2.5,
                "last": 2.4,
            },
            "skip-non-dict",
        ]

        normalized = parsing.normalize_option_records(records)

        assert len(normalized) == 2
        first, second = normalized
        assert first["ticker"] == "AAPL240119C00150000"
        assert first["implied_volatility"] == 0.31
        assert first["volume"] == 22
        assert first["greeks"]["iv"] == 0.31
        assert first["greeks"]["delta"] is None

        assert second["ticker"] == "MSFT240119P00300000"
        assert second["volume"] == "15"
        assert second["greeks"]["delta"] == "0.4"

    def test_extract_greeks_uses_contract_level_iv_when_nested_is_missing(self) -> None:
        contract = {"greeks": {"delta": 0.3}, "implied_volatility": 0.4}

        extracted = parsing.extract_greeks(contract)

        assert extracted == {
            "delta": 0.3,
            "gamma": None,
            "theta": None,
            "vega": None,
            "rho": None,
            "iv": 0.4,
        }

    def test_combine_greeks_handles_numeric_and_missing_values(self) -> None:
        long_leg = {"greeks": {"delta": 0.6, "gamma": 0.1, "iv": 0.3}}
        short_leg = {"greeks": {"delta": 0.2, "theta": -0.01, "rho": 0.03, "iv": 0.5}}

        combined = parsing.combine_greeks(long_leg, short_leg)

        assert combined["delta"] == pytest.approx(0.4)
        assert combined["gamma"] == pytest.approx(0.1)
        assert combined["theta"] == pytest.approx(0.01)
        assert combined["vega"] is None
        assert combined["rho"] == pytest.approx(-0.03)
        assert combined["iv"] == pytest.approx(0.4)

    @pytest.mark.parametrize(
        ("contract", "expected"),
        [
            ({"bid": 1.0, "ask": 3.0}, 2.0),
            ({"bid": "1", "ask": 3.0, "last": 2.5}, 2.5),
            ({"day_close": 1.25}, 1.25),
            ({"bid": 0.9}, 0.9),
            ({"ask": 1.1}, 1.1),
            ({"bid": "bad", "ask": "bad"}, None),
        ],
    )
    def test_option_mid_price_fallbacks(self, contract: dict, expected: float | None) -> None:
        assert parsing.option_mid_price(contract) == expected

    @pytest.mark.parametrize(
        ("contract", "expected"),
        [
            ({"greeks": {"delta": -0.45}}, 0.45),
            ({"greeks": {"delta": "0.8"}}, 0.8),
            ({"greeks": {"delta": 1.5}}, 1.0),
            ({"greeks": {"delta": -2}}, 1.0),
            ({"greeks": {"delta": "oops"}}, None),
            ({"greeks": "not-dict", "implied_vol": 0.2}, None),
        ],
    )
    def test_option_likelihood(self, contract: dict, expected: float | None) -> None:
        assert parsing.option_likelihood(contract) == expected


class TestTimeBucketingAndArrays:
    def test_chunk_results_by_year_ignores_invalid_timestamps(self) -> None:
        jan_2023 = int(datetime(2023, 1, 5, 12, 0, 0).timestamp() * 1000)
        jun_2024 = int(datetime(2024, 6, 10, 12, 0, 0).timestamp() * 1000)
        results = [
            {"t": jan_2023, "c": 10.0},
            {"t": jun_2024, "c": 11.0},
            {"t": None},
            {"t": "bad"},
        ]

        buckets = parsing.chunk_results_by_year(results)

        assert set(buckets.keys()) == {2023, 2024}
        assert buckets[2023] == [{"t": jan_2023, "c": 10.0}]
        assert buckets[2024] == [{"t": jun_2024, "c": 11.0}]

    def test_build_npz_payload_guarantees_expected_dtypes_and_shapes(self) -> None:
        entries = [{"t": 1000, "o": 1.0, "h": 2.0, "l": 0.5, "c": 1.5, "v": 10}, {"t": 2000, "c": 2.0}]

        payload = parsing.build_npz_payload(entries)

        assert set(payload.keys()) == {"t", "o", "h", "l", "c", "v", "n"}
        for key, values in payload.items():
            assert isinstance(values, np.ndarray)
            assert values.shape == (2,)
            if key == "t":
                assert values.dtype == np.int64
            else:
                assert values.dtype == float
        assert np.isnan(payload["o"][1])
        assert np.isnan(payload["n"][0])


class TestCacheRootAndContractFormatting:
    def test_normalize_cache_root_handles_empty_and_tilde(self) -> None:
        assert parsing.normalize_cache_root("   ") == parsing.BACKTEST_CACHE_DIR
        assert parsing.normalize_cache_root("~/cache-test") == Path("~/cache-test").expanduser()

    @pytest.mark.parametrize(
        ("value", "expected"),
        [
            (150.0, "150"),
            ("150.00", "150"),
            (150.25, "150.25"),
            ("150.20", "150.2"),
            ("not-numeric", "not-numeric"),
            (None, None),
        ],
    )
    def test_format_strike(self, value: float | int | str | None, expected: str | None) -> None:
        assert parsing.format_strike(value) == expected


class _FakeHTTPError:
    def __init__(self, body: bytes | Exception) -> None:
        self._body = body

    def read(self) -> bytes:
        if isinstance(self._body, Exception):
            raise self._body
        return self._body


class TestHttpFormatting:
    @pytest.mark.parametrize(
        ("payload", "expected"),
        [
            (b"service unavailable", "service unavailable"),
            (b'{"message": "m", "error": "e", "msg": "g"}', "m"),
            (b'{"error": "e", "msg": "g"}', "e"),
            (b'{"msg": "g"}', "g"),
            (b"<html><body><h1>Error</h1><p>Down</p></body></html>", "Error Down"),
            (b"", ""),
        ],
    )
    def test_format_http_error_detail(self, payload: bytes, expected: str) -> None:
        exc = _FakeHTTPError(payload)

        assert parsing.format_http_error_detail(exc) == expected

    def test_format_http_error_detail_returns_empty_on_read_or_decode_failure(self) -> None:
        read_failure = _FakeHTTPError(RuntimeError("cannot read"))
        decode_failure = _FakeHTTPError(b"\xff")

        assert parsing.format_http_error_detail(read_failure) == ""
        assert parsing.format_http_error_detail(decode_failure) == ""


@pytest.mark.parametrize(
    ("now_et", "expected"),
    [
        (datetime(2024, 5, 6, 15, 59, 59), date(2024, 5, 5)),
        (datetime(2024, 5, 6, 16, 0, 0), date(2024, 5, 6)),
        (datetime(2024, 5, 6, 18, 30, 0), date(2024, 5, 6)),
    ],
)
def test_effective_market_date_market_close_boundary(
    monkeypatch: pytest.MonkeyPatch,
    now_et: datetime,
    expected: date,
) -> None:
    class _FakeDateTime:
        @staticmethod
        def now(_tz):
            return now_et.replace(tzinfo=_tz)

    monkeypatch.setattr(parsing, "datetime", _FakeDateTime)

    assert parsing.effective_market_date() == expected

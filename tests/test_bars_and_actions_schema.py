from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

import pytest

from src.data_access.actions_schema import normalize_action_date, validate_actions_frame
from src.data_access.bars_schema import coerce_vendor_bar, validate_bars_frame


class FrameLikeNoDtypes:
    """Minimal DataFrame-like object that intentionally has no `dtypes`."""

    def __init__(self, rows: list[dict[str, Any]]) -> None:
        self._rows = rows
        self.columns = list(rows[0].keys()) if rows else []

    def __getitem__(self, key: str) -> list[Any]:
        return [row[key] for row in self._rows]


@pytest.mark.parametrize(
    ("payload", "symbol_override", "expected_symbol", "expected_timestamp"),
    [
        pytest.param(
            {
                "t": "2024-01-02T10:15:00",
                "o": "100.1",
                "h": "101.2",
                "l": "99.9",
                "c": "100.7",
                "v": "12345",
                "n": "42",
                "vw": "100.5",
                "sym": "AAPL",
            },
            None,
            "AAPL",
            datetime(2024, 1, 2, 10, 15, tzinfo=timezone.utc),
            id="alias-keys-with-sym",
        ),
        pytest.param(
            {
                "timestamp_utc": 1704190500,
                "o": 10,
                "h": 11,
                "l": 9,
                "c": 10.5,
                "v": 2000,
                "n": 12,
                "ticker": "MSFT",
            },
            "OVERRIDE",
            "OVERRIDE",
            datetime(2024, 1, 2, 10, 15, tzinfo=timezone.utc),
            id="symbol-override-and-epoch-seconds",
        ),
        pytest.param(
            {
                "t": 1704190500000,
                "o": 10,
                "h": 11,
                "l": 9,
                "c": 10.5,
                "v": 2000,
                "n": 12,
                "ticker": "MSFT",
            },
            None,
            "MSFT",
            datetime(2024, 1, 2, 10, 15, tzinfo=timezone.utc),
            id="epoch-milliseconds",
        ),
    ],
)
def test_coerce_vendor_bar_aliases_and_timestamp_coercion(
    payload: dict[str, Any],
    symbol_override: str | None,
    expected_symbol: str,
    expected_timestamp: datetime,
) -> None:
    result = coerce_vendor_bar(payload, symbol=symbol_override)

    assert result["symbol"] == expected_symbol
    assert result["timestamp_utc"] == expected_timestamp
    assert result["timestamp_utc"].tzinfo == timezone.utc
    assert isinstance(result["open"], float)
    assert isinstance(result["high"], float)
    assert isinstance(result["low"], float)
    assert isinstance(result["close"], float)
    assert isinstance(result["volume"], float)
    assert isinstance(result["trades"], int)


@pytest.mark.parametrize(
    ("payload", "message"),
    [
        pytest.param(
            {"t": "2024-01-02T10:15:00", "o": 1, "h": 2, "l": 0, "c": 1, "v": 10, "n": 1},
            "Missing required bar fields: symbol",
            id="missing-required-symbol",
        ),
        pytest.param(
            {"t": "2024-01-02T10:15:00", "o": "bad", "h": 2, "l": 0, "c": 1, "v": 10, "n": 1, "sym": "AAPL"},
            "open must be numeric, got 'bad'",
            id="invalid-float-coercion",
        ),
        pytest.param(
            {"t": "2024-01-02T10:15:00", "o": 1, "h": 2, "l": 0, "c": 1, "v": 10, "n": "not-int", "sym": "AAPL"},
            "trades must be an integer, got 'not-int'",
            id="invalid-int-coercion",
        ),
    ],
)
def test_coerce_vendor_bar_errors(payload: dict[str, Any], message: str) -> None:
    with pytest.raises(ValueError) as exc_info:
        coerce_vendor_bar(payload)

    assert str(exc_info.value) == message


@pytest.mark.parametrize(
    ("rows", "message"),
    [
        pytest.param(
            [
                {
                    "symbol": "AAPL",
                    "timestamp_utc": datetime(2024, 1, 2, tzinfo=timezone.utc),
                    "open": 1.0,
                    "high": 2.0,
                    "low": 0.5,
                    "close": 1.5,
                    "volume": 100.0,
                }
            ],
            "Missing required bar columns: trades",
            id="missing-required-trades",
        ),
    ],
)
def test_validate_bars_frame_missing_required_columns(rows: list[dict[str, Any]], message: str) -> None:
    pandas = pytest.importorskip("pandas")
    frame = pandas.DataFrame(rows)

    with pytest.raises(ValueError) as exc_info:
        validate_bars_frame(frame)

    assert str(exc_info.value) == message


@pytest.mark.parametrize(
    ("rows", "message"),
    [
        pytest.param(
            [
                {
                    "symbol": "AAPL",
                    "timestamp_utc": datetime(2024, 1, 2, tzinfo=timezone.utc),
                    "open": "bad",
                    "high": 2.0,
                    "low": 0.5,
                    "close": 1.5,
                    "volume": 100.0,
                    "trades": 1,
                }
            ],
            "open column must be numeric.",
            id="invalid-open-type",
        ),
        pytest.param(
            [
                {
                    "symbol": "AAPL",
                    "timestamp_utc": datetime(2024, 1, 2, tzinfo=timezone.utc),
                    "open": 1.0,
                    "high": 2.0,
                    "low": 0.5,
                    "close": 1.5,
                    "volume": 100.0,
                    "trades": 1.2,
                }
            ],
            "trades column must be integer typed.",
            id="invalid-trades-type",
        ),
    ],
)
def test_validate_bars_frame_type_checks(rows: list[dict[str, Any]], message: str) -> None:
    pandas = pytest.importorskip("pandas")
    frame = pandas.DataFrame(rows)

    with pytest.raises(ValueError) as exc_info:
        validate_bars_frame(frame)

    assert str(exc_info.value) == message


def test_validate_bars_frame_without_dtypes_returns_early() -> None:
    frame = FrameLikeNoDtypes(
        [
            {
                "symbol": "AAPL",
                "timestamp_utc": "not-a-datetime",
                "open": "bad",
                "high": "bad",
                "low": "bad",
                "close": "bad",
                "volume": "bad",
                "trades": "bad",
            }
        ]
    )

    validate_bars_frame(frame)


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        pytest.param("2024-01-15", datetime(2024, 1, 15, tzinfo=timezone.utc), id="iso-string-naive"),
        pytest.param(datetime(2024, 1, 15, 9, 30), datetime(2024, 1, 15, 9, 30, tzinfo=timezone.utc), id="naive-datetime"),
        pytest.param(
            datetime(2024, 1, 15, 9, 30, tzinfo=timezone.utc),
            datetime(2024, 1, 15, 9, 30, tzinfo=timezone.utc),
            id="aware-datetime",
        ),
        pytest.param(1705311000, datetime(2024, 1, 15, 9, 30, tzinfo=timezone.utc), id="epoch-seconds"),
        pytest.param(1705311000000, datetime(2024, 1, 15, 9, 30, tzinfo=timezone.utc), id="epoch-millis"),
    ],
)
def test_normalize_action_date_supported_inputs(value: Any, expected: datetime) -> None:
    normalized = normalize_action_date(value)

    assert normalized == expected
    assert normalized.tzinfo == timezone.utc


@pytest.mark.parametrize("value", [None, object(), ["2024-01-15"]])
def test_normalize_action_date_invalid_types(value: Any) -> None:
    with pytest.raises(ValueError, match="Unsupported action_date type"):
        normalize_action_date(value)


@pytest.mark.parametrize(
    ("rows", "message"),
    [
        pytest.param(
            [
                {
                    "symbol": "AAPL",
                    "action_type": "dividend",
                    "value": 0.25,
                }
            ],
            "Missing required action columns: action_date",
            id="missing-required-action-date",
        ),
    ],
)
def test_validate_actions_frame_missing_required_columns(rows: list[dict[str, Any]], message: str) -> None:
    pandas = pytest.importorskip("pandas")
    frame = pandas.DataFrame(rows)

    with pytest.raises(ValueError) as exc_info:
        validate_actions_frame(frame)

    assert str(exc_info.value) == message


@pytest.mark.parametrize(
    ("rows", "message"),
    [
        pytest.param(
            [
                {
                    "symbol": "AAPL",
                    "action_type": "dividend",
                    "action_date": "2024-01-15",
                    "value": 0.25,
                }
            ],
            "action_date column must be datetime64[ns]-like.",
            id="invalid-action-date-dtype",
        ),
        pytest.param(
            [
                {
                    "symbol": "AAPL",
                    "action_type": "dividend",
                    "action_date": datetime(2024, 1, 15, tzinfo=timezone.utc),
                    "value": "bad",
                }
            ],
            "value column must be numeric.",
            id="invalid-value-numeric",
        ),
        pytest.param(
            [
                {
                    "symbol": "AAPL",
                    "action_type": "split",
                    "action_date": datetime(2024, 1, 15, tzinfo=timezone.utc),
                    "value": 2.0,
                    "ratio": "bad",
                }
            ],
            "ratio column must be numeric.",
            id="invalid-ratio-numeric",
        ),
    ],
)
def test_validate_actions_frame_type_checks(rows: list[dict[str, Any]], message: str) -> None:
    pandas = pytest.importorskip("pandas")
    frame = pandas.DataFrame(rows)

    with pytest.raises(ValueError) as exc_info:
        validate_actions_frame(frame)

    assert str(exc_info.value) == message

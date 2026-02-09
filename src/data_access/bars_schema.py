"""Canonical OHLCV bar schema and validation helpers.

Default cadence is minute bars, but the schema itself is cadence-agnostic.

Canonical fields (column names and types):
- symbol (str): Ticker or instrument identifier.
- timestamp_utc (datetime): UTC timestamp for the bar open time.
- open (float)
- high (float)
- low (float)
- close (float)
- volume (float): Trade volume during the bar.
- trades (int): Number of trades in the bar.
- vwap (float, optional): Volume-weighted average price for the bar.
"""

from __future__ import annotations

from datetime import datetime, timezone
from typing import Any, Mapping

try:  # Optional dependency when validating DataFrames.
    import pandas as pd
    from pandas.api.types import is_datetime64_any_dtype, is_integer_dtype, is_numeric_dtype
except ImportError:  # pragma: no cover - pandas isn't required in this repo.
    pd = None
    is_datetime64_any_dtype = None
    is_integer_dtype = None
    is_numeric_dtype = None


# Required canonical fields that every bar payload must contain.
REQUIRED_BAR_FIELDS: tuple[str, ...] = (
    "symbol",
    "timestamp_utc",
    "open",
    "high",
    "low",
    "close",
    "volume",
    "trades",
)

# Optional canonical fields that may be provided by vendors.
OPTIONAL_BAR_FIELDS: tuple[str, ...] = ("vwap",)

# Canonical field -> expected Python type mapping for downstream validation.
CANONICAL_BAR_FIELDS: dict[str, type] = {
    "symbol": str,
    "timestamp_utc": datetime,
    "open": float,
    "high": float,
    "low": float,
    "close": float,
    "volume": float,
    "trades": int,
    "vwap": float,
}

# Vendor-provided short keys or synonyms mapped to canonical field names.
_VENDOR_ALIASES: dict[str, str] = {
    "t": "timestamp_utc",
    "o": "open",
    "h": "high",
    "l": "low",
    "c": "close",
    "v": "volume",
    "n": "trades",
    "vw": "vwap",
    "s": "symbol",
    "sym": "symbol",
    "ticker": "symbol",
}


# Convert a vendor payload to canonical bar fields, coercing types and timestamps.
def coerce_vendor_bar(payload: Mapping[str, Any], symbol: str | None = None) -> dict[str, Any]:
    """Coerce a vendor payload (e.g., Massive API t/o/h/l/c/v/n) into canonical fields."""
    mapped: dict[str, Any] = {}
    for key, value in payload.items():
        canonical = _VENDOR_ALIASES.get(key, key)
        mapped[canonical] = value

    if symbol:
        mapped["symbol"] = symbol

    missing = [field for field in REQUIRED_BAR_FIELDS if field not in mapped]
    if missing:
        raise ValueError(f"Missing required bar fields: {', '.join(missing)}")

    return {
        "symbol": str(mapped["symbol"]),
        "timestamp_utc": _coerce_timestamp(mapped["timestamp_utc"]),
        "open": _coerce_float(mapped["open"], "open"),
        "high": _coerce_float(mapped["high"], "high"),
        "low": _coerce_float(mapped["low"], "low"),
        "close": _coerce_float(mapped["close"], "close"),
        "volume": _coerce_float(mapped["volume"], "volume"),
        "trades": _coerce_int(mapped["trades"], "trades"),
        "vwap": _coerce_optional_float(mapped.get("vwap"), "vwap"),
    }


# Validate a DataFrame-like object against the canonical schema.
def validate_bars_frame(frame: Any) -> None:
    """Validate that a DataFrame-like object matches the canonical bar schema.

    Raises:
        TypeError: If the object does not look like a DataFrame.
        ValueError: If required columns are missing or invalid types are detected.
    """
    if not hasattr(frame, "columns"):
        raise TypeError("Expected a DataFrame-like object with a 'columns' attribute.")

    missing = [field for field in REQUIRED_BAR_FIELDS if field not in frame.columns]
    if missing:
        raise ValueError(f"Missing required bar columns: {', '.join(missing)}")

    if pd is None or not hasattr(frame, "dtypes"):
        return

    timestamp_series = frame["timestamp_utc"]
    if not is_datetime64_any_dtype(timestamp_series):
        raise ValueError("timestamp_utc column must be datetime64[ns]-like.")

    for field in ("open", "high", "low", "close", "volume", "vwap"):
        if field not in frame.columns:
            continue
        series = frame[field]
        if series.dropna().empty:
            continue
        if not is_numeric_dtype(series):
            raise ValueError(f"{field} column must be numeric.")

    trades_series = frame["trades"]
    if not trades_series.dropna().empty and not is_integer_dtype(trades_series):
        raise ValueError("trades column must be integer typed.")


# Convert a timestamp representation to a timezone-aware UTC datetime.
def _coerce_timestamp(value: Any) -> datetime:
    if isinstance(value, datetime):
        timestamp = value
    elif isinstance(value, (int, float)):
        timestamp = _coerce_epoch_timestamp(value)
    elif isinstance(value, str):
        timestamp = datetime.fromisoformat(value)
    else:
        raise ValueError(f"Unsupported timestamp type: {type(value)!r}")

    if timestamp.tzinfo is None:
        return timestamp.replace(tzinfo=timezone.utc)
    return timestamp.astimezone(timezone.utc)


# Interpret epoch timestamps in seconds or milliseconds.
def _coerce_epoch_timestamp(value: float) -> datetime:
    seconds = value / 1000 if value > 10**12 else value
    return datetime.fromtimestamp(seconds, tz=timezone.utc)


# Coerce numeric values to float with a helpful error message.
def _coerce_float(value: Any, field: str) -> float:
    try:
        return float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{field} must be numeric, got {value!r}") from exc


# Coerce optional numeric values to float (None is preserved).
def _coerce_optional_float(value: Any, field: str) -> float | None:
    if value is None:
        return None
    return _coerce_float(value, field)


# Coerce numeric values to int with a helpful error message.
def _coerce_int(value: Any, field: str) -> int:
    try:
        return int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{field} must be an integer, got {value!r}") from exc

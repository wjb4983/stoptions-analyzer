"""Canonical corporate actions schema and validation helpers."""

from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

try:  # Optional dependency when validating DataFrames.
    import pandas as pd
    from pandas.api.types import is_datetime64_any_dtype, is_numeric_dtype
except ImportError:  # pragma: no cover - pandas isn't required in this repo.
    pd = None
    is_datetime64_any_dtype = None
    is_numeric_dtype = None


REQUIRED_ACTION_FIELDS: tuple[str, ...] = (
    "symbol",
    "action_type",
    "action_date",
    "value",
)

OPTIONAL_ACTION_FIELDS: tuple[str, ...] = (
    "currency",
    "ratio",
    "description",
)

CANONICAL_ACTION_FIELDS: dict[str, type] = {
    "symbol": str,
    "action_type": str,
    "action_date": datetime,
    "value": float,
    "currency": str,
    "ratio": float,
    "description": str,
}


def validate_actions_frame(frame: Any) -> None:
    """Validate that a DataFrame-like object matches the canonical actions schema."""
    if not hasattr(frame, "columns"):
        raise TypeError("Expected a DataFrame-like object with a 'columns' attribute.")

    missing = [field for field in REQUIRED_ACTION_FIELDS if field not in frame.columns]
    if missing:
        raise ValueError(f"Missing required action columns: {', '.join(missing)}")

    if pd is None or not hasattr(frame, "dtypes"):
        return

    action_date_series = frame["action_date"]
    if not is_datetime64_any_dtype(action_date_series):
        raise ValueError("action_date column must be datetime64[ns]-like.")

    for field in ("value", "ratio"):
        if field not in frame.columns:
            continue
        series = frame[field]
        if series.dropna().empty:
            continue
        if not is_numeric_dtype(series):
            raise ValueError(f"{field} column must be numeric.")


def normalize_action_date(value: Any) -> datetime:
    """Normalize a corporate action date into timezone-aware UTC."""
    if isinstance(value, datetime):
        timestamp = value
    elif isinstance(value, (int, float)):
        seconds = value / 1000 if value > 10**12 else value
        timestamp = datetime.fromtimestamp(seconds, tz=timezone.utc)
    elif isinstance(value, str):
        timestamp = datetime.fromisoformat(value)
    else:
        raise ValueError(f"Unsupported action_date type: {type(value)!r}")

    if timestamp.tzinfo is None:
        return timestamp.replace(tzinfo=timezone.utc)
    return timestamp.astimezone(timezone.utc)

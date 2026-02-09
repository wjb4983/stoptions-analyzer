"""Normalization helpers for vendor bar payloads."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from typing import Any, Iterable, Iterator, Mapping

try:
    from zoneinfo import ZoneInfo
except ImportError:  # pragma: no cover - Python < 3.9
    ZoneInfo = None


@dataclass(frozen=True)
class NormalizationConfig:
    """Configuration for bar normalization.

    Attributes:
        vendor_timezone: Timezone assumption for naive timestamps (default: UTC).
        missing_bar_policy: Policy for missing bars ("drop", "ffill", or "nan").
        expected_interval: Expected bar interval for gap detection (None disables).
        adjustment_mode: Adjustment policy for prices ("raw", "split-adjusted", "total-return").
        conflict_resolution: Deterministic dedupe strategy for conflicting bars.
    """

    vendor_timezone: str | timezone = "UTC"
    missing_bar_policy: str = "drop"
    expected_interval: timedelta | None = None
    adjustment_mode: str = "raw"
    conflict_resolution: str = "prefer_last"


def normalize_bars(
    bars: Iterable[Mapping[str, Any]],
    config: NormalizationConfig | None = None,
) -> list[dict[str, Any]]:
    """Normalize vendor bars into sorted, deduped, and adjusted canonical payloads.

    The normalization pipeline:
    1) Coerce timestamps into UTC using an explicit vendor timezone assumption.
    2) Sort and dedupe by (symbol, timestamp_utc) with deterministic resolution.
    3) Apply missing-bar policy if an expected interval is provided.
    4) Apply adjustment mode (raw, split-adjusted, total-return).
    """

    config = config or NormalizationConfig()
    coerced = [_coerce_bar_timestamp(bar, config.vendor_timezone) for bar in bars]
    sorted_bars = sorted(coerced, key=lambda item: (item["symbol"], item["timestamp_utc"]))
    deduped = _dedupe_bars(sorted_bars, config.conflict_resolution)
    filled = _apply_missing_bar_policy(deduped, config.missing_bar_policy, config.expected_interval)
    adjusted = _apply_adjustment_mode(filled, config.adjustment_mode)
    return adjusted


def _coerce_bar_timestamp(bar: Mapping[str, Any], vendor_timezone: str | timezone) -> dict[str, Any]:
    """Return a bar dict with a UTC timestamp, honoring vendor timezone assumptions."""

    timestamp_value = bar.get("timestamp_utc", bar.get("timestamp"))
    if timestamp_value is None:
        raise ValueError("Bar payload missing timestamp_utc or timestamp")

    utc_timestamp = _to_utc_timestamp(timestamp_value, vendor_timezone)
    normalized = dict(bar)
    normalized["timestamp_utc"] = utc_timestamp
    normalized.pop("timestamp", None)
    return normalized


def _to_utc_timestamp(value: Any, vendor_timezone: str | timezone) -> datetime:
    """Convert a timestamp value to a UTC-aware datetime.

    Naive timestamps are localized using the provided vendor timezone assumption.
    """

    if isinstance(value, datetime):
        timestamp = value
    elif isinstance(value, (int, float)):
        seconds = value / 1000 if value > 10**12 else value
        timestamp = datetime.fromtimestamp(seconds, tz=timezone.utc)
    elif isinstance(value, str):
        timestamp = datetime.fromisoformat(value)
    else:
        raise ValueError(f"Unsupported timestamp type: {type(value)!r}")

    if timestamp.tzinfo is None:
        vendor_tz = _resolve_timezone(vendor_timezone)
        return timestamp.replace(tzinfo=vendor_tz).astimezone(timezone.utc)

    return timestamp.astimezone(timezone.utc)


def _resolve_timezone(value: str | timezone) -> timezone:
    if isinstance(value, timezone):
        return value
    if isinstance(value, str):
        if value.upper() == "UTC":
            return timezone.utc
        if ZoneInfo is None:
            raise ValueError("ZoneInfo is required for non-UTC vendor timezones.")
        return ZoneInfo(value)
    raise ValueError(f"Unsupported timezone value: {value!r}")


def _dedupe_bars(
    bars: Iterable[Mapping[str, Any]],
    conflict_resolution: str,
) -> list[dict[str, Any]]:
    """Dedupe bars on (symbol, timestamp_utc) with deterministic resolution."""

    resolution = conflict_resolution.lower()
    latest: dict[tuple[str, datetime], dict[str, Any]] = {}
    for bar in bars:
        key = (bar["symbol"], bar["timestamp_utc"])
        current = latest.get(key)
        if current is None:
            latest[key] = dict(bar)
            continue
        if resolution == "prefer_first":
            continue
        if resolution == "max_volume":
            current_volume = float(current.get("volume") or 0)
            candidate_volume = float(bar.get("volume") or 0)
            if candidate_volume > current_volume:
                latest[key] = dict(bar)
            elif candidate_volume == current_volume:
                latest[key] = dict(bar)
            continue
        if resolution != "prefer_last":
            raise ValueError(f"Unknown conflict_resolution: {conflict_resolution}")
        latest[key] = dict(bar)
    return list(latest.values())


def _apply_missing_bar_policy(
    bars: Iterable[Mapping[str, Any]],
    policy: str,
    expected_interval: timedelta | None,
) -> list[dict[str, Any]]:
    """Apply missing-bar policy (drop, ffill, nan) per symbol."""

    if expected_interval is None:
        return [dict(bar) for bar in bars]

    normalized_policy = policy.lower()
    if normalized_policy not in {"drop", "ffill", "nan"}:
        raise ValueError(f"Unknown missing_bar_policy: {policy}")

    output: list[dict[str, Any]] = []
    for symbol, symbol_bars in _group_by_symbol(bars):
        last_bar: dict[str, Any] | None = None
        last_time: datetime | None = None
        for bar in symbol_bars:
            if last_time is not None:
                gap = bar["timestamp_utc"] - last_time
                while gap > expected_interval:
                    if normalized_policy == "drop":
                        break
                    last_time = last_time + expected_interval
                    output.append(_synthesize_missing_bar(symbol, last_bar, last_time, normalized_policy))
                    gap = bar["timestamp_utc"] - last_time
            output.append(dict(bar))
            last_bar = dict(bar)
            last_time = bar["timestamp_utc"]
    return output


def _group_by_symbol(bars: Iterable[Mapping[str, Any]]) -> Iterator[tuple[str, list[dict[str, Any]]]]:
    grouped: dict[str, list[dict[str, Any]]] = {}
    for bar in bars:
        grouped.setdefault(bar["symbol"], []).append(dict(bar))
    for symbol in sorted(grouped):
        symbol_bars = sorted(grouped[symbol], key=lambda item: item["timestamp_utc"])
        yield symbol, symbol_bars


def _synthesize_missing_bar(
    symbol: str,
    previous_bar: Mapping[str, Any] | None,
    timestamp: datetime,
    policy: str,
) -> dict[str, Any]:
    """Create a placeholder bar for missing data points."""

    placeholder = {
        "symbol": symbol,
        "timestamp_utc": timestamp,
        "open": float("nan"),
        "high": float("nan"),
        "low": float("nan"),
        "close": float("nan"),
        "volume": 0.0,
        "trades": 0,
        "vwap": float("nan"),
    }

    if policy == "ffill" and previous_bar:
        for field in ("open", "high", "low", "close", "vwap"):
            if field in previous_bar:
                placeholder[field] = previous_bar[field]
    return placeholder


def _apply_adjustment_mode(bars: Iterable[Mapping[str, Any]], mode: str) -> list[dict[str, Any]]:
    """Apply adjustment modes for splits and dividends."""

    normalized_mode = mode.lower()
    if normalized_mode == "raw":
        return [dict(bar) for bar in bars]
    if normalized_mode not in {"split-adjusted", "total-return"}:
        raise ValueError(f"Unknown adjustment_mode: {mode}")

    grouped: list[dict[str, Any]] = []
    for symbol, symbol_bars in _group_by_symbol(bars):
        adjusted = _apply_adjustments_for_symbol(symbol_bars, normalized_mode)
        grouped.extend(adjusted)
    return sorted(grouped, key=lambda item: (item["symbol"], item["timestamp_utc"]))


def _apply_adjustments_for_symbol(
    bars: list[dict[str, Any]],
    mode: str,
) -> list[dict[str, Any]]:
    """Adjust a single symbol's bars for splits/dividends.

    Split factors are applied to all prior bars (reverse traversal). Dividends
    follow the total-return adjustment factor of (close - dividend) / close.
    """

    split_factor = 1.0
    total_return_factor = 1.0
    adjusted: list[dict[str, Any]] = []

    for bar in reversed(bars):
        bar_copy = dict(bar)
        price_factor = split_factor * total_return_factor
        if price_factor != 1.0:
            for field in ("open", "high", "low", "close", "vwap"):
                if field in bar_copy and bar_copy[field] is not None:
                    bar_copy[field] = bar_copy[field] / price_factor
            if bar_copy.get("volume") is not None:
                bar_copy["volume"] = bar_copy["volume"] * price_factor
        adjusted.append(bar_copy)

        split = _coerce_factor(bar.get("split_factor", 1.0))
        if split and split > 0:
            split_factor *= split

        if mode == "total-return":
            dividend = bar.get("dividend") or bar.get("cash_dividend") or 0.0
            close_price = bar.get("close")
            if close_price:
                total_return_factor *= (close_price - dividend) / close_price

    return list(reversed(adjusted))


def _coerce_factor(value: Any) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 1.0

"""Massive cache-backed data provider."""

from __future__ import annotations

import importlib.util
import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable, Sequence

import numpy as np

from config import BACKTEST_CACHE_DIR
from data_access.actions_schema import OPTIONAL_ACTION_FIELDS, REQUIRED_ACTION_FIELDS
from data_access.bars_schema import coerce_vendor_bar, validate_bars_frame
from data_access.api_client import MassiveApiClient
from data_access.cache import _safe_ticker_name
from data_access.normalization import validate_asof_membership
from data_access.provider_base import DataProvider, FORWARD_KNOWN_FIELD_NAMES

_pd_spec = importlib.util.find_spec("pandas")
if _pd_spec:
    import pandas as pd
else:
    pd = None


class MassiveCacheProvider(DataProvider):
    """Load cached minute bars from the Massive backtest cache."""

    def __init__(self, cache_root: str | Path | None = None, symbols: Sequence[str] | None = None) -> None:
        self.cache_root = Path(cache_root).expanduser() if cache_root else BACKTEST_CACHE_DIR
        self._symbols = list(symbols) if symbols else []

    def list_symbols(self) -> list[str]:
        return list(self._symbols)

    def get_bars(
        self,
        symbols: Sequence[str],
        start: datetime | str,
        end: datetime | str,
        timeframe: str = "1m",
    ):
        if timeframe not in {"1m", "1min", "1minute"}:
            raise ValueError("Massive cache provider only supports 1-minute bars.")

        start_dt = _normalize_datetime(start)
        end_dt = _normalize_datetime(end)
        selected = list(symbols) if symbols else list(self._symbols)
        if not selected:
            return _empty_bars_frame()

        frames = []
        records: list[dict] = []
        for symbol in selected:
            payloads = list(_load_symbol_bars(symbol, start_dt, end_dt, self.cache_root))
            if pd is not None:
                if payloads:
                    frames.append(_records_to_frame(payloads))
            else:
                records.extend(payloads)

        if pd is not None:
            if not frames:
                return _empty_bars_frame()
            combined = pd.concat(frames, ignore_index=True)
            validate_bars_frame(combined)
            return combined

        return records

    def get_corporate_actions(
        self,
        symbols: Sequence[str],
        start: datetime | str,
        end: datetime | str,
    ):
        start_dt = _normalize_datetime(start)
        end_dt = _normalize_datetime(end)
        selected = list(symbols) if symbols else list(self._symbols)
        records: list[dict[str, object]] = []
        records.extend(_load_cached_corporate_actions(selected, start_dt, end_dt, self.cache_root))

        api_key = os.getenv("MASSIVE_API_KEY")
        if api_key:
            client = MassiveApiClient(api_key)
            for symbol in selected:
                try:
                    records.extend(_fetch_api_corporate_actions(client, symbol, start_dt, end_dt))
                except Exception:
                    continue

        if pd is None:
            return records
        if not records:
            return pd.DataFrame(columns=[*REQUIRED_ACTION_FIELDS, *OPTIONAL_ACTION_FIELDS])
        frame = pd.DataFrame.from_records(records)
        frame["action_date"] = pd.to_datetime(frame["action_date"], utc=True)
        frame = frame.sort_values(["symbol", "action_date", "action_type"]).reset_index(drop=True)
        return frame


def _load_cached_corporate_actions(
    symbols: Sequence[str],
    start_dt: datetime,
    end_dt: datetime,
    cache_root: Path,
) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for symbol in symbols:
        safe = _safe_ticker_name(symbol)
        path = cache_root / safe / "1m" / "corporate_actions.json"
        if not path.exists():
            continue
        try:
            payload = json.loads(path.read_text())
        except json.JSONDecodeError:
            continue
        entries = payload if isinstance(payload, list) else payload.get("actions", [])
        for entry in entries:
            normalized = _normalize_action_record(symbol, entry)
            if normalized is None:
                continue
            action_dt = normalized["action_date"]
            if start_dt <= action_dt <= end_dt:
                rows.append(normalized)
    return rows


def _fetch_api_corporate_actions(
    client: MassiveApiClient,
    symbol: str,
    start_dt: datetime,
    end_dt: datetime,
) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for entry in client.fetch_dividends(symbol):
        row = _normalize_dividend_from_api(symbol, entry)
        if row is not None and start_dt <= row["action_date"] <= end_dt:
            rows.append(row)
    for entry in client.fetch_splits(symbol):
        row = _normalize_split_from_api(symbol, entry)
        if row is not None and start_dt <= row["action_date"] <= end_dt:
            rows.append(row)
    return rows


def _normalize_action_record(symbol: str, row: dict) -> dict[str, object] | None:
    action_type = str(row.get("action_type", "")).strip().lower()
    if action_type not in {"split", "dividend"}:
        return None
    raw_date = row.get("action_date") or row.get("ex_dividend_date") or row.get("execution_date")
    if raw_date is None:
        return None
    dt = _normalize_datetime(str(raw_date))
    value = float(row.get("value", 0.0) or 0.0)
    ratio = row.get("ratio")
    ratio_value = float(ratio) if ratio not in (None, "") else None
    return {
        "symbol": symbol,
        "action_type": action_type,
        "action_date": dt,
        "value": value,
        "currency": row.get("currency", "USD"),
        "ratio": ratio_value,
        "description": row.get("description", ""),
    }


def _normalize_dividend_from_api(symbol: str, row: dict) -> dict[str, object] | None:
    raw_date = row.get("ex_dividend_date") or row.get("pay_date")
    if raw_date is None:
        return None
    return {
        "symbol": symbol,
        "action_type": "dividend",
        "action_date": _normalize_datetime(str(raw_date)),
        "value": float(row.get("cash_amount", row.get("value", 0.0)) or 0.0),
        "currency": row.get("currency", "USD"),
        "ratio": None,
        "description": "massive_dividend",
    }


def _normalize_split_from_api(symbol: str, row: dict) -> dict[str, object] | None:
    raw_date = row.get("execution_date") or row.get("date")
    if raw_date is None:
        return None
    split_to = row.get("split_to")
    split_from = row.get("split_from")
    if split_to is not None and split_from not in (None, 0, "0"):
        factor = float(split_to) / float(split_from)
    else:
        factor = float(row.get("ratio", 1.0) or 1.0)
    return {
        "symbol": symbol,
        "action_type": "split",
        "action_date": _normalize_datetime(str(raw_date)),
        "value": float(factor),
        "currency": row.get("currency", "USD"),
        "ratio": float(factor),
        "description": "massive_split",
    }


def _normalize_datetime(value: datetime | str) -> datetime:
    if isinstance(value, datetime):
        timestamp = value
    else:
        timestamp = datetime.fromisoformat(value)

    if timestamp.tzinfo is None:
        return timestamp.replace(tzinfo=timezone.utc)
    return timestamp.astimezone(timezone.utc)


def _load_symbol_bars(
    symbol: str, start_dt: datetime, end_dt: datetime, cache_root: Path
) -> Iterable[dict]:
    yield from _load_npz_bars(symbol, start_dt, end_dt, cache_root)

    legacy = list(_load_legacy_json_bars(symbol, start_dt, end_dt, cache_root))
    if legacy:
        yield from legacy


def _load_npz_bars(
    symbol: str, start_dt: datetime, end_dt: datetime, cache_root: Path
) -> Iterable[dict]:
    safe = _safe_ticker_name(symbol)
    root = cache_root / safe / "1m"
    if not root.exists():
        return

    start_year = start_dt.year
    end_year = end_dt.year
    for year in range(start_year, end_year + 1):
        path = root / f"{safe}_1m_{year}.npz"
        if not path.exists():
            continue
        yield from _load_npz_file(symbol, path, start_dt, end_dt)


def _load_npz_file(
    symbol: str, path: Path, start_dt: datetime, end_dt: datetime
) -> Iterable[dict]:
    with np.load(path, mmap_mode="r") as data:
        _validate_no_forward_known_fields(data.files)
        timestamps = data.get("t")
        if timestamps is None:
            return
        start_ms = int(start_dt.timestamp() * 1000)
        end_ms = int(end_dt.timestamp() * 1000)
        mask = (timestamps >= start_ms) & (timestamps <= end_ms)
        if not mask.any():
            return
        active_from_values = data.get("active_from")
        active_to_values = data.get("active_to")
        active_from = None if active_from_values is None else int(np.asarray(active_from_values).reshape(-1)[0])
        active_to = None if active_to_values is None else int(np.asarray(active_to_values).reshape(-1)[0])

        idx = np.nonzero(mask)[0]
        for pos in idx:
            ts_ms = int(timestamps[pos])
            ts_utc = datetime.fromtimestamp(ts_ms / 1000, tz=timezone.utc)
            from_dt = None if active_from is None else datetime.fromtimestamp(active_from / 1000, tz=timezone.utc)
            to_dt = None if active_to is None else datetime.fromtimestamp(active_to / 1000, tz=timezone.utc)
            if not validate_asof_membership(symbol=symbol, timestamp_utc=ts_utc, active_from=from_dt, active_to=to_dt):
                continue

            payload = {
                "t": int(timestamps[pos]),
                "o": data.get("o")[pos],
                "h": data.get("h")[pos],
                "l": data.get("l")[pos],
                "c": data.get("c")[pos],
                "v": data.get("v")[pos],
                "n": data.get("n")[pos],
            }
            yield coerce_vendor_bar(payload, symbol=symbol)


def _load_legacy_json_bars(
    symbol: str, start_dt: datetime, end_dt: datetime, cache_root: Path
) -> Iterable[dict]:
    start_date = start_dt.date().isoformat()
    end_date = end_dt.date().isoformat()
    safe = _safe_ticker_name(symbol)
    expected_name = f"{safe}_1m_{start_date}_{end_date}.json"
    candidates = [cache_root / expected_name, BACKTEST_CACHE_DIR / expected_name]
    fallback_glob = list(cache_root.glob(f"{safe}_1m_*.json"))
    candidates.extend(fallback_glob)

    seen: set[Path] = set()
    for path in candidates:
        if path in seen or not path.exists():
            continue
        seen.add(path)
        try:
            payload = json.loads(path.read_text())
        except json.JSONDecodeError:
            continue
        results = payload.get("results", [])
        for entry in results:
            timestamp = entry.get("t")
            if timestamp is None:
                continue
            if not _timestamp_in_range(timestamp, start_dt, end_dt):
                continue
            yield coerce_vendor_bar(entry, symbol=symbol)


def _timestamp_in_range(value: float, start_dt: datetime, end_dt: datetime) -> bool:
    ts = datetime.fromtimestamp(value / 1000, tz=timezone.utc)
    return start_dt <= ts <= end_dt


def _records_to_frame(records: Sequence[dict]):
    if pd is None:
        return records
    frame = pd.DataFrame.from_records(records)
    if not frame.empty:
        frame["timestamp_utc"] = pd.to_datetime(frame["timestamp_utc"], utc=True)
        frame["trades"] = pd.to_numeric(frame["trades"], errors="coerce").astype("Int64")
    return frame


def _empty_bars_frame():
    if pd is None:
        return []
    return pd.DataFrame(
        columns=["symbol", "timestamp_utc", "open", "high", "low", "close", "volume", "trades", "vwap"]
    )


def _validate_no_forward_known_fields(columns: Sequence[str]) -> None:
    lowered = {str(name).lower() for name in columns}
    overlaps = sorted(field for field in FORWARD_KNOWN_FIELD_NAMES if field in lowered)
    if overlaps:
        raise ValueError(f"Forward-known fields detected in provider payload: {', '.join(overlaps)}")

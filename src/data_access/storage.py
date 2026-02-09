from __future__ import annotations

import json
from collections import OrderedDict
from datetime import date, datetime, timezone
from pathlib import Path
import numpy as np

from config import BACKTEST_CACHE_DIR
from data_access.bars_schema import coerce_vendor_bar
from data_access.cache import _safe_ticker_name
from utils.parsing import build_npz_payload, chunk_results_by_year


class LocalBarStore:
    """Read/write year-partitioned bar data with optional in-memory caching."""

    def __init__(
        self,
        cache_root: str | Path | None = None,
        timeframe: str = "1m",
        memory_cache_size: int = 0,
    ) -> None:
        self.cache_root = Path(cache_root).expanduser() if cache_root else BACKTEST_CACHE_DIR
        self.timeframe = timeframe
        self.memory_cache_size = max(0, int(memory_cache_size))
        self._npz_cache: "OrderedDict[Path, dict[str, np.ndarray]]" = OrderedDict()

    def load_index(self, symbol: str) -> dict | None:
        path = self._index_path(symbol)
        if not path.exists():
            return None
        try:
            return json.loads(path.read_text())
        except json.JSONDecodeError:
            return None

    def write_index(self, symbol: str, payload: dict) -> None:
        path = self._index_path(symbol)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(payload, indent=2))

    def list_years(self, symbol: str) -> list[int]:
        index = self.load_index(symbol)
        if index:
            years = index.get("years")
            if isinstance(years, list):
                return sorted({year for year in years if isinstance(year, int)})
        ticker_dir = self._ticker_dir(symbol)
        if not ticker_dir.exists():
            return []
        safe = _safe_ticker_name(symbol)
        years: set[int] = set()
        for path in ticker_dir.glob(f"{safe}_{self.timeframe}_*.npz"):
            suffix = path.stem.split("_")[-1]
            if suffix.isdigit():
                years.add(int(suffix))
        return sorted(years)

    def store_bars(
        self,
        symbol: str,
        results: list[dict],
        start_date: date,
        end_date: date,
        full_range: bool = True,
    ) -> dict:
        buckets = chunk_results_by_year(results)
        safe = _safe_ticker_name(symbol)
        ticker_dir = self._ticker_dir(symbol)
        ticker_dir.mkdir(parents=True, exist_ok=True)
        for year, entries in buckets.items():
            payload = build_npz_payload(entries)
            np.savez_compressed(ticker_dir / f"{safe}_{self.timeframe}_{year}.npz", **payload)

        index_payload = {
            "ticker": symbol,
            "start_date": start_date.isoformat(),
            "end_date": end_date.isoformat(),
            "full_range": full_range,
            "fetched_at": datetime.now(timezone.utc).isoformat(),
            "years": sorted(buckets.keys()),
        }
        self.write_index(symbol, index_payload)
        return index_payload

    def load_bars(
        self,
        symbol: str,
        start: datetime | str,
        end: datetime | str,
    ) -> list[dict]:
        start_dt = _normalize_datetime(start)
        end_dt = _normalize_datetime(end)
        start_ms = int(start_dt.timestamp() * 1000)
        end_ms = int(end_dt.timestamp() * 1000)

        records: list[dict] = []
        for year in range(start_dt.year, end_dt.year + 1):
            path = self._npz_path(symbol, year)
            if not path.exists():
                continue
            payload = self._load_npz_payload(path)
            timestamps = payload.get("t")
            if timestamps is None:
                continue
            mask = (timestamps >= start_ms) & (timestamps <= end_ms)
            if not mask.any():
                continue
            idx = np.nonzero(mask)[0]
            for pos in idx:
                record = {
                    "t": int(timestamps[pos]),
                    "o": payload.get("o")[pos],
                    "h": payload.get("h")[pos],
                    "l": payload.get("l")[pos],
                    "c": payload.get("c")[pos],
                    "v": payload.get("v")[pos],
                    "n": payload.get("n")[pos],
                }
                records.append(coerce_vendor_bar(record, symbol=symbol))
        return records

    def _load_npz_payload(self, path: Path) -> dict[str, np.ndarray]:
        if self.memory_cache_size <= 0:
            with np.load(path, mmap_mode="r") as data:
                return {key: np.array(value) for key, value in data.items()}

        cached = self._npz_cache.get(path)
        if cached is not None:
            self._npz_cache.move_to_end(path)
            return cached

        with np.load(path, mmap_mode="r") as data:
            payload = {key: np.array(value) for key, value in data.items()}

        self._npz_cache[path] = payload
        self._npz_cache.move_to_end(path)
        if len(self._npz_cache) > self.memory_cache_size:
            self._npz_cache.popitem(last=False)
        return payload

    def _ticker_dir(self, symbol: str) -> Path:
        safe = _safe_ticker_name(symbol)
        return self.cache_root / safe / self.timeframe

    def _index_path(self, symbol: str) -> Path:
        return self._ticker_dir(symbol) / "index.json"

    def _npz_path(self, symbol: str, year: int) -> Path:
        safe = _safe_ticker_name(symbol)
        return self._ticker_dir(symbol) / f"{safe}_{self.timeframe}_{year}.npz"


def _normalize_datetime(value: datetime | str) -> datetime:
    if isinstance(value, datetime):
        timestamp = value
    else:
        timestamp = datetime.fromisoformat(value)

    if timestamp.tzinfo is None:
        return timestamp.replace(tzinfo=timezone.utc)
    return timestamp.astimezone(timezone.utc)

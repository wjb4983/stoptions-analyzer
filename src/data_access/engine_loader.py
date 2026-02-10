from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Sequence

import numpy as np

from config import BACKTEST_CACHE_DIR
from data_access.cache import _safe_ticker_name


@dataclass(frozen=True)
class EngineArrayMetadata:
    """Metadata accompanying canonical engine arrays."""

    symbol_to_column: dict[str, int]
    date_index: np.ndarray
    missingness_ratio: float
    missingness_by_symbol: dict[str, float]


@dataclass(frozen=True)
class EngineArrayBundle:
    """Canonical 2D arrays expected by backtest engines."""

    date_index: np.ndarray
    open_prices: np.ndarray
    close_prices: np.ndarray
    missing_mask: np.ndarray
    metadata: EngineArrayMetadata


def load_canonical_price_arrays(
    symbols: Sequence[str],
    start: datetime | str,
    end: datetime | str,
    *,
    cache_root: str | Path | None = None,
    timeframe: str = "1m",
    lookback_window: int = 0,
    validate_split_adjustment: bool = True,
) -> EngineArrayBundle:
    """Load aligned open/close arrays and missing mask from local NPZ cache."""

    if lookback_window < 0:
        raise ValueError("lookback_window must be non-negative")

    normalized_symbols = [str(symbol).strip().upper() for symbol in symbols if str(symbol).strip()]
    if not normalized_symbols:
        raise ValueError("At least one symbol is required.")

    start_dt = _normalize_datetime(start)
    end_dt = _normalize_datetime(end)
    if end_dt < start_dt:
        raise ValueError("end must be greater than or equal to start")

    root = Path(cache_root).expanduser() if cache_root else BACKTEST_CACHE_DIR
    symbol_series: dict[str, tuple[np.ndarray, np.ndarray, np.ndarray]] = {}

    for symbol in normalized_symbols:
        ts, open_values, close_values, split_factors = _load_symbol_npz_range(
            symbol=symbol,
            root=root,
            timeframe=timeframe,
            start_dt=start_dt,
            end_dt=end_dt,
        )

        _validate_timestamps(symbol, ts)
        if validate_split_adjustment:
            _validate_split_adjustment(symbol, split_factors)

        symbol_series[symbol] = (
            ts.astype(np.int64, copy=False),
            np.asarray(open_values, dtype=np.float64),
            np.asarray(close_values, dtype=np.float64),
        )

    all_timestamps = [ts for ts, _, _ in symbol_series.values() if ts.size > 0]
    aligned_index = np.unique(np.concatenate(all_timestamps)) if all_timestamps else np.array([], dtype=np.int64)

    open_prices = np.full((aligned_index.size, len(normalized_symbols)), np.nan, dtype=np.float64)
    close_prices = np.full((aligned_index.size, len(normalized_symbols)), np.nan, dtype=np.float64)
    missing_mask = np.ones((aligned_index.size, len(normalized_symbols)), dtype=bool)

    for col, symbol in enumerate(normalized_symbols):
        ts, open_values, close_values = symbol_series[symbol]
        if ts.size == 0:
            continue
        idx = np.searchsorted(aligned_index, ts)
        open_prices[idx, col] = open_values
        close_prices[idx, col] = close_values
        missing_mask[idx, col] = False

    for col, symbol in enumerate(normalized_symbols):
        available_bars = int((~missing_mask[:, col]).sum())
        if available_bars < lookback_window:
            raise ValueError(
                f"Insufficient history for {symbol}: requires {lookback_window} bars, got {available_bars}."
            )

    missingness_by_symbol = {
        symbol: float(np.mean(missing_mask[:, col])) if aligned_index.size else 1.0
        for col, symbol in enumerate(normalized_symbols)
    }
    metadata = EngineArrayMetadata(
        symbol_to_column={symbol: idx for idx, symbol in enumerate(normalized_symbols)},
        date_index=aligned_index,
        missingness_ratio=float(np.mean(missing_mask)) if missing_mask.size else 1.0,
        missingness_by_symbol=missingness_by_symbol,
    )

    return EngineArrayBundle(
        date_index=aligned_index,
        open_prices=open_prices,
        close_prices=close_prices,
        missing_mask=missing_mask,
        metadata=metadata,
    )


def _load_symbol_npz_range(
    *,
    symbol: str,
    root: Path,
    timeframe: str,
    start_dt: datetime,
    end_dt: datetime,
) -> tuple[np.ndarray, np.ndarray, np.ndarray, np.ndarray | None]:
    safe = _safe_ticker_name(symbol)
    ticker_root = root / safe / timeframe

    start_ms = int(start_dt.timestamp() * 1000)
    end_ms = int(end_dt.timestamp() * 1000)

    timestamps_parts: list[np.ndarray] = []
    open_parts: list[np.ndarray] = []
    close_parts: list[np.ndarray] = []
    split_parts: list[np.ndarray] = []

    for year in range(start_dt.year, end_dt.year + 1):
        path = ticker_root / f"{safe}_{timeframe}_{year}.npz"
        if not path.exists():
            continue
        with np.load(path, mmap_mode="r") as payload:
            timestamps = np.asarray(payload.get("t"), dtype=np.int64)
            if timestamps.size == 0:
                continue
            mask = (timestamps >= start_ms) & (timestamps <= end_ms)
            if not mask.any():
                continue

            open_values = np.asarray(payload.get("o"), dtype=np.float64)
            close_values = np.asarray(payload.get("c"), dtype=np.float64)
            timestamps_parts.append(timestamps[mask])
            open_parts.append(open_values[mask])
            close_parts.append(close_values[mask])

            split_values = _extract_split_factors(payload)
            if split_values is not None:
                split_parts.append(np.asarray(split_values, dtype=np.float64)[mask])

    if not timestamps_parts:
        empty = np.array([], dtype=np.float64)
        return (
            np.array([], dtype=np.int64),
            empty,
            empty,
            None,
        )

    ts = np.concatenate(timestamps_parts)
    open_values = np.concatenate(open_parts)
    close_values = np.concatenate(close_parts)
    split_values = np.concatenate(split_parts) if split_parts else None
    return ts, open_values, close_values, split_values


def _extract_split_factors(payload: np.lib.npyio.NpzFile) -> np.ndarray | None:
    for key in ("split_factor", "split", "sf"):
        values = payload.get(key)
        if values is not None:
            return np.asarray(values)
    return None


def _validate_timestamps(symbol: str, timestamps: np.ndarray) -> None:
    if timestamps.size <= 1:
        return
    deltas = np.diff(timestamps)
    if np.any(deltas == 0):
        raise ValueError(f"Duplicate timestamps detected for {symbol}.")
    if np.any(deltas < 0):
        raise ValueError(f"Non-monotonic timestamps detected for {symbol}.")


def _validate_split_adjustment(symbol: str, split_factors: np.ndarray | None) -> None:
    if split_factors is None or split_factors.size == 0:
        return
    if np.any(~np.isfinite(split_factors)):
        raise ValueError(f"Invalid split factors (non-finite) for {symbol}.")
    if np.any(split_factors <= 0):
        raise ValueError(f"Invalid split factors (<=0) for {symbol}.")


def _normalize_datetime(value: datetime | str) -> datetime:
    if isinstance(value, datetime):
        dt = value
    else:
        dt = datetime.fromisoformat(value)
    if dt.tzinfo is None:
        return dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)

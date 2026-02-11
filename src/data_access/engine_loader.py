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
    coverage_by_symbol: dict[str, float]
    tradable_ratio_by_symbol: dict[str, float]
    excluded_symbols: dict[str, str]
    audit_summary_by_symbol: dict[str, dict[str, float | int]]


@dataclass(frozen=True)
class EngineArrayBundle:
    """Canonical 2D arrays expected by backtest engines."""

    date_index: np.ndarray
    open_prices: np.ndarray
    close_prices: np.ndarray
    missing_mask: np.ndarray
    metadata: EngineArrayMetadata


@dataclass(frozen=True)
class SymbolDatasetAudit:
    """Per-symbol quality audit details captured during dataset loading."""

    symbol: str
    bars_total: int
    bars_in_universe: int
    bars_tradable: int
    missing_ratio: float
    coverage_ratio: float


@dataclass(frozen=True)
class DatasetContractsReport:
    """Validation report describing accepted and excluded symbols."""

    per_symbol: dict[str, SymbolDatasetAudit]
    excluded_symbols: dict[str, str]
    audit_summary_by_symbol: dict[str, dict[str, float | int]]


@dataclass(frozen=True)
class _LoadedSymbolDataset:
    timestamps: np.ndarray
    open_values: np.ndarray
    close_values: np.ndarray
    tradable_mask: np.ndarray
    split_factors: np.ndarray | None
    dividends: np.ndarray | None
    in_universe_total: int
    original_total: int


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
    symbol_series: dict[str, _LoadedSymbolDataset] = {}
    excluded_symbols: dict[str, str] = {}

    for symbol in normalized_symbols:
        loaded = _load_symbol_npz_range(
            symbol=symbol,
            root=root,
            timeframe=timeframe,
            start_dt=start_dt,
            end_dt=end_dt,
        )

        ts = loaded.timestamps
        _validate_timestamps(symbol, ts)
        if validate_split_adjustment:
            _validate_split_adjustment(symbol, loaded.split_factors)
        _validate_adjustment_consistency(symbol=symbol, split_factors=loaded.split_factors, dividends=loaded.dividends)
        _validate_missing_data_contract(
            symbol=symbol,
            open_values=loaded.open_values,
            close_values=loaded.close_values,
            tradable_mask=loaded.tradable_mask,
        )

        if ts.size == 0:
            excluded_symbols[symbol] = "No point-in-time universe overlap in requested window"
            continue

        symbol_series[symbol] = loaded

    if not symbol_series:
        rejected = ", ".join(f"{sym}: {reason}" for sym, reason in excluded_symbols.items())
        raise ValueError(f"No symbols passed dataset contracts. {rejected}")

    accepted_symbols = list(symbol_series)

    all_timestamps = [loaded.timestamps for loaded in symbol_series.values() if loaded.timestamps.size > 0]
    aligned_index = np.unique(np.concatenate(all_timestamps)) if all_timestamps else np.array([], dtype=np.int64)

    open_prices = np.full((aligned_index.size, len(accepted_symbols)), np.nan, dtype=np.float64)
    close_prices = np.full((aligned_index.size, len(accepted_symbols)), np.nan, dtype=np.float64)
    missing_mask = np.ones((aligned_index.size, len(accepted_symbols)), dtype=bool)
    per_symbol_audit: dict[str, SymbolDatasetAudit] = {}

    for col, symbol in enumerate(accepted_symbols):
        loaded = symbol_series[symbol]
        ts = loaded.timestamps
        open_values = loaded.open_values
        close_values = loaded.close_values
        if ts.size == 0:
            continue
        idx = np.searchsorted(aligned_index, ts)
        open_prices[idx, col] = open_values
        close_prices[idx, col] = close_values
        symbol_missing = ~loaded.tradable_mask
        missing_mask[idx, col] = symbol_missing

        missing_ratio = float(np.mean(symbol_missing)) if symbol_missing.size else 1.0
        coverage_ratio = 0.0
        if loaded.in_universe_total > 0:
            coverage_ratio = float(np.sum(loaded.tradable_mask)) / float(loaded.in_universe_total)
        per_symbol_audit[symbol] = SymbolDatasetAudit(
            symbol=symbol,
            bars_total=loaded.original_total,
            bars_in_universe=loaded.in_universe_total,
            bars_tradable=int(np.sum(loaded.tradable_mask)),
            missing_ratio=missing_ratio,
            coverage_ratio=coverage_ratio,
        )

    for col, symbol in enumerate(accepted_symbols):
        available_bars = int((~missing_mask[:, col]).sum())
        if available_bars < lookback_window:
            raise ValueError(
                f"Insufficient history for {symbol}: requires {lookback_window} bars, got {available_bars}."
            )

    missingness_by_symbol = {
        symbol: float(np.mean(missing_mask[:, col])) if aligned_index.size else 1.0
        for col, symbol in enumerate(accepted_symbols)
    }
    coverage_by_symbol = {
        symbol: per_symbol_audit[symbol].coverage_ratio
        for symbol in accepted_symbols
    }
    tradable_ratio_by_symbol = {
        symbol: (1.0 - missingness_by_symbol[symbol])
        for symbol in accepted_symbols
    }
    metadata = EngineArrayMetadata(
        symbol_to_column={symbol: idx for idx, symbol in enumerate(accepted_symbols)},
        date_index=aligned_index,
        missingness_ratio=float(np.mean(missing_mask)) if missing_mask.size else 1.0,
        missingness_by_symbol=missingness_by_symbol,
        coverage_by_symbol=coverage_by_symbol,
        tradable_ratio_by_symbol=tradable_ratio_by_symbol,
        excluded_symbols=excluded_symbols,
        audit_summary_by_symbol={
            symbol: {
                "bars_total": audit.bars_total,
                "bars_in_universe": audit.bars_in_universe,
                "bars_tradable": audit.bars_tradable,
                "missing_ratio": audit.missing_ratio,
                "coverage_ratio": audit.coverage_ratio,
            }
            for symbol, audit in per_symbol_audit.items()
        },
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
) -> _LoadedSymbolDataset:
    safe = _safe_ticker_name(symbol)
    ticker_root = root / safe / timeframe

    start_ms = int(start_dt.timestamp() * 1000)
    end_ms = int(end_dt.timestamp() * 1000)

    timestamps_parts: list[np.ndarray] = []
    open_parts: list[np.ndarray] = []
    close_parts: list[np.ndarray] = []
    split_parts: list[np.ndarray] = []
    dividend_parts: list[np.ndarray] = []
    tradable_parts: list[np.ndarray] = []
    universe_total = 0
    original_total = 0

    for year in range(start_dt.year, end_dt.year + 1):
        path = ticker_root / f"{safe}_{timeframe}_{year}.npz"
        if not path.exists():
            continue
        with np.load(path, mmap_mode="r") as payload:
            timestamps = np.asarray(payload.get("t"), dtype=np.int64)
            if timestamps.size == 0:
                continue
            original_total += int(timestamps.size)
            mask = (timestamps >= start_ms) & (timestamps <= end_ms)
            if not mask.any():
                continue

            open_values = np.asarray(payload.get("o"), dtype=np.float64)
            close_values = np.asarray(payload.get("c"), dtype=np.float64)
            timestamps_parts.append(timestamps[mask])
            open_parts.append(open_values[mask])
            close_parts.append(close_values[mask])

            active_from, active_to = _extract_active_window(payload)
            pit_mask = _build_pit_membership_mask(
                timestamps=timestamps,
                active_from=active_from,
                active_to=active_to,
            )
            universe_total += int(np.sum(pit_mask & mask))

            split_values = _extract_split_factors(payload)
            if split_values is not None:
                split_parts.append(np.asarray(split_values, dtype=np.float64)[mask])

            dividends = _extract_dividend_values(payload)
            if dividends is not None:
                dividend_parts.append(np.asarray(dividends, dtype=np.float64)[mask])

            tradable_values = _extract_tradable_mask(payload, fallback_open=open_values, fallback_close=close_values)
            tradable_parts.append(np.asarray(tradable_values, dtype=bool)[mask] & pit_mask[mask])

    if not timestamps_parts:
        empty = np.array([], dtype=np.float64)
        return _LoadedSymbolDataset(
            timestamps=np.array([], dtype=np.int64),
            open_values=empty,
            close_values=empty,
            tradable_mask=np.array([], dtype=bool),
            split_factors=None,
            dividends=None,
            in_universe_total=0,
            original_total=original_total,
        )

    ts = np.concatenate(timestamps_parts)
    open_values = np.concatenate(open_parts)
    close_values = np.concatenate(close_parts)
    split_values = np.concatenate(split_parts) if split_parts else None
    dividend_values = np.concatenate(dividend_parts) if dividend_parts else None
    tradable_values = np.concatenate(tradable_parts) if tradable_parts else np.ones(ts.size, dtype=bool)

    delisting_return = _extract_terminal_return(path=ticker_root)
    if delisting_return is not None:
        _validate_delisting_contract(symbol=symbol, timestamps=ts, terminal_return=delisting_return)
        close_values[-1] = close_values[-1] * (1.0 + delisting_return)

    return _LoadedSymbolDataset(
        timestamps=ts,
        open_values=open_values,
        close_values=close_values,
        tradable_mask=tradable_values,
        split_factors=split_values,
        dividends=dividend_values,
        in_universe_total=universe_total,
        original_total=original_total,
    )


def _extract_active_window(payload: np.lib.npyio.NpzFile) -> tuple[int | None, int | None]:
    active_from = payload.get("active_from")
    active_to = payload.get("active_to")
    from_value = int(np.asarray(active_from).reshape(-1)[0]) if active_from is not None else None
    to_value = int(np.asarray(active_to).reshape(-1)[0]) if active_to is not None else None
    if from_value is not None and to_value is not None and to_value < from_value:
        raise ValueError("Invalid PIT window: active_to < active_from")
    return from_value, to_value


def _build_pit_membership_mask(*, timestamps: np.ndarray, active_from: int | None, active_to: int | None) -> np.ndarray:
    mask = np.ones(timestamps.size, dtype=bool)
    if active_from is not None:
        mask &= timestamps >= int(active_from)
    if active_to is not None:
        mask &= timestamps <= int(active_to)
    return mask


def _extract_tradable_mask(
    payload: np.lib.npyio.NpzFile,
    *,
    fallback_open: np.ndarray,
    fallback_close: np.ndarray,
) -> np.ndarray:
    for key in ("tradable", "is_tradable"):
        values = payload.get(key)
        if values is not None:
            return np.asarray(values, dtype=bool)
    return np.isfinite(fallback_open) & np.isfinite(fallback_close)


def _extract_terminal_return(*, path: Path) -> float | None:
    terminal_path = path / "terminal.json"
    if not terminal_path.exists():
        return None
    import json

    payload = json.loads(terminal_path.read_text())
    value = payload.get("delisting_return")
    if value is None:
        return None
    return float(value)




def _extract_dividend_values(payload: np.lib.npyio.NpzFile) -> np.ndarray | None:
    for key in ("dividend", "div", "cash_dividend"):
        values = payload.get(key)
        if values is not None:
            return np.asarray(values)
    return None
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




def _validate_adjustment_consistency(
    *,
    symbol: str,
    split_factors: np.ndarray | None,
    dividends: np.ndarray | None,
) -> None:
    if dividends is not None:
        if np.any(~np.isfinite(dividends)):
            raise ValueError(f"Invalid dividend values (non-finite) for {symbol}.")
        if np.any(dividends < 0):
            raise ValueError(f"Invalid dividend values (<0) for {symbol}.")
    if split_factors is None or dividends is None:
        return
    if split_factors.size != dividends.size:
        raise ValueError(f"Split/dividend length mismatch for {symbol}.")
def _validate_missing_data_contract(
    *,
    symbol: str,
    open_values: np.ndarray,
    close_values: np.ndarray,
    tradable_mask: np.ndarray,
) -> None:
    if tradable_mask.size not in (0, open_values.size):
        raise ValueError(f"Tradable mask length mismatch for {symbol}.")
    if tradable_mask.size == 0:
        return
    invalid_tradable = tradable_mask & (~np.isfinite(open_values) | ~np.isfinite(close_values))
    if np.any(invalid_tradable):
        raise ValueError(f"Tradable bars with missing prices detected for {symbol}.")


def _validate_delisting_contract(*, symbol: str, timestamps: np.ndarray, terminal_return: float) -> None:
    if not np.isfinite(terminal_return):
        raise ValueError(f"Invalid delisting return for {symbol}.")
    if timestamps.size == 0:
        raise ValueError(f"Delisting return supplied for {symbol} without terminal bars.")



@dataclass(frozen=True)
class DatasetValidationSummary:
    """Data-quality summary consumed by backtest runners."""

    coverage_by_symbol: dict[str, float]
    missingness_by_symbol: dict[str, float]
    excluded_symbols: dict[str, str]
    reasons_by_symbol: dict[str, str]


def validate_engine_dataset_contracts(
    bundle: EngineArrayBundle,
    *,
    max_missingness: float = 1.0,
) -> DatasetValidationSummary:
    """Validate dataset-level contracts and produce audit summary."""

    if not (0.0 <= max_missingness <= 1.0):
        raise ValueError("max_missingness must be between 0 and 1")

    reasons_by_symbol: dict[str, str] = {}
    for symbol, ratio in bundle.metadata.missingness_by_symbol.items():
        if ratio > max_missingness:
            reasons_by_symbol[symbol] = f"missingness {ratio:.3f} exceeds allowed {max_missingness:.3f}"

    if reasons_by_symbol:
        formatted = ", ".join(f"{sym}: {reason}" for sym, reason in reasons_by_symbol.items())
        raise ValueError(f"Dataset contracts failed: {formatted}")

    return DatasetValidationSummary(
        coverage_by_symbol=dict(bundle.metadata.coverage_by_symbol),
        missingness_by_symbol=dict(bundle.metadata.missingness_by_symbol),
        excluded_symbols=dict(bundle.metadata.excluded_symbols),
        reasons_by_symbol=reasons_by_symbol,
    )

def _normalize_datetime(value: datetime | str) -> datetime:
    if isinstance(value, datetime):
        dt = value
    else:
        dt = datetime.fromisoformat(value)
    if dt.tzinfo is None:
        return dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)

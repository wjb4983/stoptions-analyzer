from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Sequence

from data_access.provider_base import DataProvider

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
    asset_class_by_symbol: dict[str, str]
    expiry_by_symbol: dict[str, str | None]
    multiplier_by_symbol: dict[str, float]
    borrow_availability_tier_by_symbol: dict[str, str]
    financing_benchmark_by_symbol: dict[str, str]
    pit_membership_violations_by_symbol: dict[str, int]
    adjustment_violations_by_symbol: dict[str, int]
    delisted_symbols: list[str]
    survivorship_bias_flags_by_symbol: dict[str, bool]
    leakage_flags_by_symbol: dict[str, bool]


@dataclass(frozen=True)
class EngineArrayBundle:
    """Canonical 2D arrays expected by backtest engines."""

    date_index: np.ndarray
    open_prices: np.ndarray
    close_prices: np.ndarray
    raw_open_prices: np.ndarray
    raw_close_prices: np.ndarray
    split_factors: np.ndarray
    dividends: np.ndarray
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
    pit_membership_violations: int
    adjustment_violations: int
    delisted: bool


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
    asset_class: str
    expiry: str | None
    multiplier: float
    borrow_availability_tier: str
    financing_benchmark: str
    pit_membership_violations: int
    adjustment_violations: int
    delisted: bool
    forward_known_fields: tuple[str, ...]


def load_canonical_price_arrays(
    symbols: Sequence[str],
    start: datetime | str,
    end: datetime | str,
    *,
    cache_root: str | Path | None = None,
    timeframe: str = "1m",
    lookback_window: int = 0,
    validate_split_adjustment: bool = True,
    provider: DataProvider | None = None,
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
        _validate_adjustment_consistency(
            symbol=symbol,
            split_factors=loaded.split_factors,
            dividends=loaded.dividends,
            open_values=loaded.open_values,
            close_values=loaded.close_values,
            tradable_mask=loaded.tradable_mask,
        )
        _validate_missing_data_contract(
            symbol=symbol,
            open_values=loaded.open_values,
            close_values=loaded.close_values,
            tradable_mask=loaded.tradable_mask,
        )
        _validate_no_forward_known_fields(symbol=symbol, forward_known_fields=loaded.forward_known_fields)

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
    raw_open_prices = np.full((aligned_index.size, len(accepted_symbols)), np.nan, dtype=np.float64)
    raw_close_prices = np.full((aligned_index.size, len(accepted_symbols)), np.nan, dtype=np.float64)
    split_factors = np.ones((aligned_index.size, len(accepted_symbols)), dtype=np.float64)
    dividends = np.zeros((aligned_index.size, len(accepted_symbols)), dtype=np.float64)
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
        raw_open_prices[idx, col] = open_values
        raw_close_prices[idx, col] = close_values
        open_prices[idx, col] = open_values
        close_prices[idx, col] = close_values
        if loaded.split_factors is not None:
            split_factors[idx, col] = loaded.split_factors
        if loaded.dividends is not None:
            dividends[idx, col] = loaded.dividends
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
            pit_membership_violations=loaded.pit_membership_violations,
            adjustment_violations=loaded.adjustment_violations,
            delisted=loaded.delisted,
        )

    _apply_provider_corporate_actions(
        provider=provider,
        symbols=accepted_symbols,
        start_dt=start_dt,
        end_dt=end_dt,
        aligned_index=aligned_index,
        symbol_to_column={symbol: idx for idx, symbol in enumerate(accepted_symbols)},
        split_factors=split_factors,
        dividends=dividends,
    )

    adjusted_open_prices, adjusted_close_prices = _build_adjusted_price_views(
        raw_open_prices=raw_open_prices,
        raw_close_prices=raw_close_prices,
        split_factors=split_factors,
        dividends=dividends,
    )
    open_prices = adjusted_open_prices
    close_prices = adjusted_close_prices

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
                "pit_membership_violations": audit.pit_membership_violations,
                "adjustment_violations": audit.adjustment_violations,
                "delisted": int(audit.delisted),
                "survivorship_bias_flag": int(audit.delisted or audit.pit_membership_violations > 0),
                "leakage_flag": int(bool(symbol_series[symbol].forward_known_fields)),
            }
            for symbol, audit in per_symbol_audit.items()
        },
        asset_class_by_symbol={symbol: symbol_series[symbol].asset_class for symbol in accepted_symbols},
        expiry_by_symbol={symbol: symbol_series[symbol].expiry for symbol in accepted_symbols},
        multiplier_by_symbol={symbol: symbol_series[symbol].multiplier for symbol in accepted_symbols},
        borrow_availability_tier_by_symbol={
            symbol: symbol_series[symbol].borrow_availability_tier for symbol in accepted_symbols
        },
        financing_benchmark_by_symbol={symbol: symbol_series[symbol].financing_benchmark for symbol in accepted_symbols},
        pit_membership_violations_by_symbol={
            symbol: symbol_series[symbol].pit_membership_violations for symbol in accepted_symbols
        },
        adjustment_violations_by_symbol={
            symbol: symbol_series[symbol].adjustment_violations for symbol in accepted_symbols
        },
        delisted_symbols=[symbol for symbol in accepted_symbols if symbol_series[symbol].delisted],
        survivorship_bias_flags_by_symbol={
            symbol: bool(symbol_series[symbol].delisted or symbol_series[symbol].pit_membership_violations > 0)
            for symbol in accepted_symbols
        },
        leakage_flags_by_symbol={
            symbol: bool(symbol_series[symbol].forward_known_fields)
            for symbol in accepted_symbols
        },
    )

    return EngineArrayBundle(
        date_index=aligned_index,
        open_prices=open_prices,
        close_prices=close_prices,
        raw_open_prices=raw_open_prices,
        raw_close_prices=raw_close_prices,
        split_factors=split_factors,
        dividends=dividends,
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
    pit_violations_total = 0
    forward_known_fields: set[str] = set()
    universe_total = 0
    original_total = 0
    symbol_metadata: dict[str, str | float | None] = {
        "asset_class": "equity",
        "expiry": None,
        "multiplier": 1.0,
        "borrow_availability_tier": "normal",
        "financing_benchmark": "overnight",
    }

    for year in range(start_dt.year, end_dt.year + 1):
        path = ticker_root / f"{safe}_{timeframe}_{year}.npz"
        if not path.exists():
            continue
        with np.load(path, mmap_mode="r") as payload:
            symbol_metadata.update(_extract_instrument_metadata(payload))
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
            pit_violations_total += int(np.sum(mask & ~pit_mask))

            forward_known_fields.update(_extract_forward_known_fields(payload))

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
            asset_class=str(symbol_metadata["asset_class"]),
            expiry=str(symbol_metadata["expiry"]) if symbol_metadata["expiry"] is not None else None,
            multiplier=float(symbol_metadata["multiplier"]),
            borrow_availability_tier=str(symbol_metadata["borrow_availability_tier"]),
            financing_benchmark=str(symbol_metadata["financing_benchmark"]),
            pit_membership_violations=0,
            adjustment_violations=0,
            delisted=False,
            forward_known_fields=tuple(sorted(forward_known_fields)),
        )

    ts = np.concatenate(timestamps_parts)
    open_values = np.concatenate(open_parts)
    close_values = np.concatenate(close_parts)
    split_values = np.concatenate(split_parts) if split_parts else None
    dividend_values = np.concatenate(dividend_parts) if dividend_parts else None
    tradable_values = np.concatenate(tradable_parts) if tradable_parts else np.ones(ts.size, dtype=bool)

    adjustment_violations = _count_adjustment_violations(
        split_factors=split_values,
        dividends=dividend_values,
        open_values=open_values,
        close_values=close_values,
        tradable_mask=tradable_values,
    )

    delisting_return = _extract_terminal_return(path=ticker_root)
    is_delisted = delisting_return is not None
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
        asset_class=str(symbol_metadata["asset_class"]),
        expiry=str(symbol_metadata["expiry"]) if symbol_metadata["expiry"] is not None else None,
        multiplier=float(symbol_metadata["multiplier"]),
        borrow_availability_tier=str(symbol_metadata["borrow_availability_tier"]),
        financing_benchmark=str(symbol_metadata["financing_benchmark"]),
        pit_membership_violations=pit_violations_total,
        adjustment_violations=adjustment_violations,
        delisted=is_delisted,
        forward_known_fields=tuple(sorted(forward_known_fields)),
    )


def _extract_instrument_metadata(payload: np.lib.npyio.NpzFile) -> dict[str, str | float | None]:
    def _first_scalar(name: str) -> str | float | None:
        values = payload.get(name)
        if values is None:
            return None
        arr = np.asarray(values)
        if arr.size == 0:
            return None
        value = arr.reshape(-1)[0]
        if isinstance(value, bytes):
            return value.decode("utf-8")
        return value.item() if hasattr(value, "item") else value

    asset_class = _first_scalar("asset_class")
    expiry = _first_scalar("expiry")
    multiplier = _first_scalar("multiplier")
    borrow_tier = _first_scalar("borrow_availability_tier")
    financing_benchmark = _first_scalar("financing_benchmark")

    return {
        "asset_class": str(asset_class).strip().lower() if asset_class is not None else "equity",
        "expiry": str(expiry).strip() if expiry not in (None, "", "none") else None,
        "multiplier": float(multiplier) if multiplier is not None else 1.0,
        "borrow_availability_tier": str(borrow_tier).strip().lower() if borrow_tier is not None else "normal",
        "financing_benchmark": str(financing_benchmark).strip().lower()
        if financing_benchmark is not None
        else "overnight",
    }


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



def _extract_forward_known_fields(payload: np.lib.npyio.NpzFile) -> set[str]:
    fields: set[str] = set()
    for key in payload.files:
        lowered = str(key).lower()
        if any(token in lowered for token in ("future", "lookahead", "next_", "target", "label")):
            fields.add(str(key))
    return fields


def _count_adjustment_violations(
    *,
    split_factors: np.ndarray | None,
    dividends: np.ndarray | None,
    open_values: np.ndarray,
    close_values: np.ndarray,
    tradable_mask: np.ndarray,
) -> int:
    violations = 0
    if split_factors is not None:
        violations += int(np.sum(~np.isfinite(split_factors) | (split_factors <= 0)))
    if dividends is not None:
        violations += int(np.sum(~np.isfinite(dividends) | (dividends < 0)))
        if tradable_mask.size == dividends.size:
            violations += int(np.sum((dividends > close_values) & tradable_mask))
    return int(violations)


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
    open_values: np.ndarray,
    close_values: np.ndarray,
    tradable_mask: np.ndarray,
) -> None:
    if dividends is not None:
        if np.any(~np.isfinite(dividends)):
            raise ValueError(f"Invalid dividend values (non-finite) for {symbol}.")
        if np.any(dividends < 0):
            raise ValueError(f"Invalid dividend values (<0) for {symbol}.")
    if split_factors is not None and split_factors.size not in (0, open_values.size):
        raise ValueError(f"Split factor length mismatch for {symbol}.")
    if dividends is not None and dividends.size not in (0, close_values.size):
        raise ValueError(f"Dividend length mismatch for {symbol}.")
    if dividends is not None and tradable_mask.size == dividends.size and np.any((dividends > close_values) & tradable_mask):
        raise ValueError(f"Dividend exceeds close for tradable bars for {symbol}.")
    if split_factors is None or dividends is None:
        return
    if split_factors.size != dividends.size:
        raise ValueError(f"Split/dividend length mismatch for {symbol}.")
def _validate_no_forward_known_fields(*, symbol: str, forward_known_fields: tuple[str, ...]) -> None:
    if forward_known_fields:
        raise ValueError(f"Forward-known fields are not allowed for {symbol}: {', '.join(forward_known_fields)}")


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
    survivorship_bias_flags_by_symbol: dict[str, bool]
    leakage_flags_by_symbol: dict[str, bool]


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
        survivorship_bias_flags_by_symbol=dict(bundle.metadata.survivorship_bias_flags_by_symbol),
        leakage_flags_by_symbol=dict(bundle.metadata.leakage_flags_by_symbol),
    )



def _apply_provider_corporate_actions(
    *,
    provider: DataProvider | None,
    symbols: list[str],
    start_dt: datetime,
    end_dt: datetime,
    aligned_index: np.ndarray,
    symbol_to_column: dict[str, int],
    split_factors: np.ndarray,
    dividends: np.ndarray,
) -> None:
    if provider is None or aligned_index.size == 0 or not symbols:
        return
    actions = provider.get_corporate_actions(symbols=symbols, start=start_dt, end=end_dt)
    rows = _normalize_action_rows(actions)
    if not rows:
        return
    timestamp_to_row = {int(ts): idx for idx, ts in enumerate(aligned_index.tolist())}
    for row in rows:
        symbol = str(row.get("symbol", "")).upper()
        if symbol not in symbol_to_column:
            continue
        action_ts = _normalize_action_timestamp(row.get("action_date"))
        row_idx = timestamp_to_row.get(action_ts)
        if row_idx is None:
            continue
        col_idx = symbol_to_column[symbol]
        action_type = str(row.get("action_type", "")).strip().lower()
        value = float(row.get("value", 0.0) or 0.0)
        if action_type == "split":
            ratio = row.get("ratio", value)
            split_value = float(ratio or 1.0)
            if split_value > 0:
                split_factors[row_idx, col_idx] = split_value
        elif action_type == "dividend":
            if value >= 0:
                dividends[row_idx, col_idx] = value


def _normalize_action_rows(actions: Any) -> list[dict[str, Any]]:
    if actions is None:
        return []
    if isinstance(actions, list):
        return [dict(row) for row in actions if isinstance(row, dict)]
    if hasattr(actions, "to_dict"):
        try:
            return [dict(row) for row in actions.to_dict(orient="records")]
        except Exception:
            pass
    if hasattr(actions, "rows"):
        return [dict(row) for row in getattr(actions, "rows", []) if isinstance(row, dict)]
    return []


def _normalize_action_timestamp(value: Any) -> int:
    dt = value
    if isinstance(value, str):
        dt = datetime.fromisoformat(value)
    elif isinstance(value, (int, float)):
        ts = float(value)
        if ts > 10**12:
            return int(ts)
        return int(ts * 1000)
    if isinstance(dt, datetime):
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        else:
            dt = dt.astimezone(timezone.utc)
        return int(dt.timestamp() * 1000)
    raise ValueError(f"Unsupported action timestamp type: {type(value)!r}")


def _build_adjusted_price_views(
    *,
    raw_open_prices: np.ndarray,
    raw_close_prices: np.ndarray,
    split_factors: np.ndarray,
    dividends: np.ndarray,
) -> tuple[np.ndarray, np.ndarray]:
    open_adj = np.asarray(raw_open_prices, dtype=float).copy()
    close_adj = np.asarray(raw_close_prices, dtype=float).copy()
    if open_adj.size == 0:
        return open_adj, close_adj
    n_rows, n_cols = open_adj.shape
    for col in range(n_cols):
        factor = 1.0
        for row in range(n_rows - 1, -1, -1):
            if np.isfinite(open_adj[row, col]):
                open_adj[row, col] = open_adj[row, col] * factor
            if np.isfinite(close_adj[row, col]):
                close_adj[row, col] = close_adj[row, col] * factor
            split = float(split_factors[row, col]) if np.isfinite(split_factors[row, col]) else 1.0
            close_raw = raw_close_prices[row, col]
            div = float(dividends[row, col]) if np.isfinite(dividends[row, col]) else 0.0
            total_return_factor = 1.0
            if np.isfinite(close_raw) and close_raw > 0.0 and div > 0.0 and div <= close_raw:
                total_return_factor = (close_raw - div) / close_raw
            if split > 0:
                factor *= split * total_return_factor
    return open_adj, close_adj

def _normalize_datetime(value: datetime | str) -> datetime:
    if isinstance(value, datetime):
        dt = value
    else:
        dt = datetime.fromisoformat(value)
    if dt.tzinfo is None:
        return dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)

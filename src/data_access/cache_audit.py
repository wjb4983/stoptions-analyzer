from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import numpy as np

from data_access.cache import _safe_ticker_name


def _year_from_filename(path: Path, safe_symbol: str, timeframe: str) -> int | None:
    prefix = f"{safe_symbol}_{timeframe}_"
    stem = path.stem
    if not stem.startswith(prefix):
        return None
    year_str = stem[len(prefix) :]
    if len(year_str) != 4 or not year_str.isdigit():
        return None
    return int(year_str)


def _format_timestamp(ts_ms: int | None) -> str | None:
    if ts_ms is None:
        return None
    return datetime.fromtimestamp(ts_ms / 1000, tz=timezone.utc).isoformat()


def audit_symbol_history(
    symbol: str,
    cache_root: Path | str,
    min_years: int,
    timeframe: str,
    *,
    min_bars_per_year: int = 1,
) -> dict[str, Any]:
    safe_symbol = _safe_ticker_name(symbol)
    root = Path(cache_root).expanduser()
    symbol_dir = root / safe_symbol / timeframe

    now_year = datetime.now(timezone.utc).year
    required_years = max(1, int(min_years))
    required_window = list(range(now_year - required_years + 1, now_year + 1))

    per_year_bar_counts: dict[str, int] = {}
    corrupted_years: list[int] = []
    available_years: set[int] = set()
    covered_years: set[int] = set()
    earliest_ts: int | None = None
    latest_ts: int | None = None

    if symbol_dir.exists():
        for path in sorted(symbol_dir.glob(f"{safe_symbol}_{timeframe}_*.npz")):
            year = _year_from_filename(path, safe_symbol, timeframe)
            if year is None:
                continue
            available_years.add(year)
            try:
                with np.load(path, mmap_mode="r") as payload:
                    timestamps = payload.get("t")
                    if timestamps is None:
                        per_year_bar_counts[str(year)] = 0
                        continue
                    ts = np.asarray(timestamps, dtype=np.int64).reshape(-1)
            except Exception:
                corrupted_years.append(year)
                per_year_bar_counts[str(year)] = 0
                continue

            bar_count = int(ts.size)
            per_year_bar_counts[str(year)] = bar_count
            if year in required_window and bar_count >= int(min_bars_per_year):
                covered_years.add(year)
            if bar_count:
                min_ts = int(np.min(ts))
                max_ts = int(np.max(ts))
                earliest_ts = min_ts if earliest_ts is None else min(earliest_ts, min_ts)
                latest_ts = max_ts if latest_ts is None else max(latest_ts, max_ts)

    missing_year_gaps = [year for year in required_window if year not in available_years]
    insufficient_bars_years = [
        year
        for year in required_window
        if year in available_years and year not in corrupted_years and per_year_bar_counts.get(str(year), 0) < int(min_bars_per_year)
    ]
    covered_required_years = sorted(covered_years)
    coverage_ratio = float(len(covered_required_years) / len(required_window)) if required_window else 0.0

    return {
        "symbol": symbol,
        "safe_symbol": safe_symbol,
        "timeframe": timeframe,
        "cache_root": str(root),
        "required_year_window": required_window,
        "available_years": sorted(available_years),
        "covered_years": covered_required_years,
        "coverage_ratio": coverage_ratio,
        "earliest_timestamp": _format_timestamp(earliest_ts),
        "latest_timestamp": _format_timestamp(latest_ts),
        "missing_year_gaps": missing_year_gaps,
        "insufficient_bars_years": insufficient_bars_years,
        "per_year_bar_counts": per_year_bar_counts,
        "corruption_flags": {"has_corruption": bool(corrupted_years), "corrupted_years": sorted(corrupted_years)},
    }


def audit_universe_history(
    symbols: list[str],
    cache_root: Path | str,
    min_years: int,
    timeframe: str,
    *,
    strict: bool = True,
    min_symbol_coverage_ratio: float = 1.0,
    min_bars_per_year: int = 1,
) -> dict[str, Any]:
    normalized_symbols = [str(symbol).strip().upper() for symbol in symbols if str(symbol).strip()]
    symbol_audits: dict[str, dict[str, Any]] = {}

    for symbol in normalized_symbols:
        symbol_audits[symbol] = audit_symbol_history(
            symbol=symbol,
            cache_root=cache_root,
            min_years=min_years,
            timeframe=timeframe,
            min_bars_per_year=min_bars_per_year,
        )

    failing_symbols: list[str] = []
    failing_details: dict[str, dict[str, Any]] = {}
    per_symbol_pass: dict[str, bool] = {}
    required_symbol_threshold = 1.0 if strict else max(0.0, min(1.0, float(min_symbol_coverage_ratio)))

    for symbol, audit in symbol_audits.items():
        has_corruption = bool(audit.get("corruption_flags", {}).get("has_corruption", False))
        coverage_ratio = float(audit.get("coverage_ratio", 0.0))
        symbol_pass = (not has_corruption) and (coverage_ratio >= required_symbol_threshold)
        per_symbol_pass[symbol] = symbol_pass
        if not symbol_pass:
            failing_symbols.append(symbol)
            failing_details[symbol] = {
                "coverage_ratio": coverage_ratio,
                "missing_year_gaps": audit.get("missing_year_gaps", []),
                "insufficient_bars_years": audit.get("insufficient_bars_years", []),
                "corrupted_years": audit.get("corruption_flags", {}).get("corrupted_years", []),
            }

    total = len(normalized_symbols)
    passing = total - len(failing_symbols)
    pass_ratio = float(passing / total) if total else 1.0
    universe_pass_threshold = 1.0 if strict else max(0.0, min(1.0, float(min_symbol_coverage_ratio)))

    return {
        "strict": bool(strict),
        "min_years": int(min_years),
        "timeframe": timeframe,
        "min_symbol_coverage_ratio": float(min_symbol_coverage_ratio),
        "min_bars_per_year": int(min_bars_per_year),
        "total_symbols": total,
        "passing_symbols": passing,
        "failing_symbols": failing_symbols,
        "failing_details": failing_details,
        "per_symbol_pass": per_symbol_pass,
        "symbol_audits": symbol_audits,
        "pass_ratio": pass_ratio,
        "required_pass_ratio": universe_pass_threshold,
        "pass": pass_ratio >= universe_pass_threshold,
    }

from __future__ import annotations

from datetime import datetime, timezone

import numpy as np

from data_access.cache import _safe_ticker_name
from data_access.cache_audit import audit_symbol_history, audit_universe_history


def _write_npz(cache_root, symbol: str, year: int, bars: int = 4):
    safe = _safe_ticker_name(symbol)
    folder = cache_root / safe / "1m"
    folder.mkdir(parents=True, exist_ok=True)
    start = int(datetime(year, 1, 2, tzinfo=timezone.utc).timestamp() * 1000)
    t = np.arange(start, start + bars * 60_000, 60_000, dtype=np.int64)
    c = np.linspace(100.0, 101.0, bars)
    np.savez_compressed(folder / f"{safe}_1m_{year}.npz", t=t, c=c)


def test_audit_symbol_history_happy_path_full_history(tmp_path):
    now_year = datetime.now(timezone.utc).year
    min_years = 3
    for year in range(now_year - min_years + 1, now_year + 1):
        _write_npz(tmp_path, "AAPL", year, bars=5)

    report = audit_symbol_history("AAPL", tmp_path, min_years=min_years, timeframe="1m", min_bars_per_year=1)

    assert report["missing_year_gaps"] == []
    assert report["corruption_flags"]["has_corruption"] is False
    assert report["coverage_ratio"] == 1.0
    assert len(report["covered_years"]) == min_years
    assert report["earliest_timestamp"] is not None and report["latest_timestamp"] is not None


def test_audit_universe_history_sparse_symbol_failure(tmp_path):
    now_year = datetime.now(timezone.utc).year
    _write_npz(tmp_path, "AAPL", now_year, bars=5)

    report = audit_universe_history(
        symbols=["AAPL", "MSFT"],
        cache_root=tmp_path,
        min_years=2,
        timeframe="1m",
        strict=True,
        min_symbol_coverage_ratio=1.0,
        min_bars_per_year=1,
    )

    assert report["pass"] is False
    assert set(report["failing_symbols"]) == {"AAPL", "MSFT"}
    assert "missing_year_gaps" in report["failing_details"]["MSFT"]


def test_audit_symbol_history_handles_corrupted_npz(tmp_path):
    now_year = datetime.now(timezone.utc).year
    safe = _safe_ticker_name("TSLA")
    folder = tmp_path / safe / "1m"
    folder.mkdir(parents=True, exist_ok=True)
    (folder / f"{safe}_1m_{now_year}.npz").write_text("not a real npz", encoding="utf-8")

    report = audit_symbol_history("TSLA", tmp_path, min_years=1, timeframe="1m", min_bars_per_year=1)

    assert report["corruption_flags"]["has_corruption"] is True
    assert now_year in report["corruption_flags"]["corrupted_years"]
    assert report["coverage_ratio"] == 0.0

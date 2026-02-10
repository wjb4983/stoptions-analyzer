from __future__ import annotations

from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import numpy as np

from src.analysis.reporting import build_drawdown_rows, format_backtest_report
from src.backtesting import cache_runner


def test_build_drawdown_rows_returns_sorted_worst_first() -> None:
    base = datetime(2024, 1, 1, tzinfo=timezone.utc)
    timestamps = np.array(
        [int((base + timedelta(days=idx)).timestamp() * 1000) for idx in range(5)],
        dtype=np.int64,
    )
    equity = np.array([1.0, 1.1, 0.9, 0.95, 1.05], dtype=float)

    rows = build_drawdown_rows(timestamps, equity, top_n=2)

    assert len(rows) == 2
    assert rows[0]["drawdown"] <= rows[1]["drawdown"]
    assert rows[0]["drawdown"] < 0


def test_run_time_series_momentum_backtest_persists_exports(tmp_path: Path) -> None:
    cache_root = tmp_path / "cache"
    output_root = tmp_path / "outputs"
    cache_runner.BACKTEST_OUTPUT_DIR = output_root

    symbol = "AAA"
    safe = symbol
    symbol_dir = cache_root / safe / "1m"
    symbol_dir.mkdir(parents=True)

    start = datetime(2024, 1, 1, tzinfo=timezone.utc)
    timestamps = np.array(
        [int((start + timedelta(minutes=idx)).timestamp() * 1000) for idx in range(8)],
        dtype=np.int64,
    )
    close = np.array([100, 101, 102, 103, 104, 105, 106, 107], dtype=float)
    open_ = close - 0.1
    np.savez_compressed(
        symbol_dir / f"{safe}_1m_2024.npz",
        t=timestamps,
        o=open_,
        c=close,
        h=close,
        l=close,
        v=np.ones_like(close),
        n=np.ones_like(close),
    )

    output = cache_runner.run_time_series_momentum_backtest(
        tickers=[symbol],
        start_date=date(2024, 1, 1),
        end_date=date(2024, 1, 1),
        cache_root=cache_root,
        lookback_days=2,
        skip_days=1,
        costs_bps=5.0,
    )

    assert "Summary Metrics" in output
    run_dirs = list(output_root.glob("tsmom_backtest_*"))
    assert len(run_dirs) == 1
    run_dir = run_dirs[0]
    for name in [
        "equity.csv",
        "equity.json",
        "returns.csv",
        "returns.json",
        "trades.csv",
        "trades.json",
        "metrics.csv",
        "metrics.json",
        "report.txt",
    ]:
        assert (run_dir / name).exists(), name


def test_format_backtest_report_contains_required_sections() -> None:
    report = format_backtest_report(
        title="TSMOM",
        params={"lookback_days": 20},
        metrics={"total_return": 0.2},
        drawdown_rows=[
            {
                "timestamp": "2024-01-01T00:00:00",
                "drawdown": -0.1,
                "equity": 0.9,
                "running_peak": 1.0,
            }
        ],
        turnover_stats={"mean": 0.1, "total": 1.0, "max": 0.2},
        cost_totals={"total": 0.01, "slippage": 0.01, "fees": 0.0, "borrow": 0.0},
    )

    assert "Summary Metrics" in report
    assert "Drawdown Table" in report
    assert "Turnover and Cost Attribution" in report

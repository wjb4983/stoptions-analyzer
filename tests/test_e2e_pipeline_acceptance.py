from __future__ import annotations

from datetime import date, datetime, timedelta, timezone
from pathlib import Path
import csv
import json

import numpy as np
import pytest

from src.analysis.time_series.momentum import TimeSeriesMomentumSettings, build_time_series_momentum_arrays
from src.backtesting import cache_runner
from src.data_access.engine_loader import load_canonical_price_arrays


def _write_symbol_dataset(*, cache_root: Path, symbol: str, timestamps: np.ndarray, close: np.ndarray) -> None:
    symbol_dir = cache_root / symbol / "1m"
    symbol_dir.mkdir(parents=True, exist_ok=True)
    np.savez_compressed(
        symbol_dir / f"{symbol}_1m_2024.npz",
        t=timestamps,
        o=close - 0.1,
        c=close,
        h=close + 0.2,
        l=close - 0.2,
        v=np.full(close.shape, 1000.0),
        n=np.ones(close.shape),
    )


def _build_minimal_synthetic_cache(cache_root: Path) -> tuple[list[str], np.ndarray]:
    symbols = ["AAA", "BBB"]
    start = datetime(2024, 1, 1, 9, 30, tzinfo=timezone.utc)
    timestamps = np.array(
        [int((start + timedelta(minutes=idx)).timestamp() * 1000) for idx in range(48)],
        dtype=np.int64,
    )
    base_close = np.linspace(100.0, 104.0, num=timestamps.size, dtype=float)
    _write_symbol_dataset(cache_root=cache_root, symbol="AAA", timestamps=timestamps, close=base_close)
    _write_symbol_dataset(cache_root=cache_root, symbol="BBB", timestamps=timestamps, close=base_close + 0.75)
    return symbols, timestamps


def _latest_dir(root: Path, pattern: str) -> Path:
    return sorted(root.glob(pattern))[-1]


def test_e2e_acceptance_pipeline_from_ingestion_to_manifest_and_leaderboard(tmp_path: Path) -> None:
    cache_root = tmp_path / "cache"
    output_root = tmp_path / "outputs"
    cache_runner.BACKTEST_OUTPUT_DIR = output_root

    symbols, _ = _build_minimal_synthetic_cache(cache_root)

    bundle = load_canonical_price_arrays(
        symbols=symbols,
        start="2024-01-01T09:30:00+00:00",
        end="2024-01-01T10:17:00+00:00",
        cache_root=cache_root,
    )
    assert bundle.close_prices.shape == (48, 2)

    momentum_arrays = build_time_series_momentum_arrays(
        closes=bundle.close_prices[:, 0].tolist(),
        settings=TimeSeriesMomentumSettings(lookback_days=5, skip_days=1, vol_window_days=5),
    )
    assert momentum_arrays.tradable_position.shape[0] == bundle.close_prices.shape[0]
    assert np.count_nonzero(np.isfinite(momentum_arrays.raw_score)) > 0

    cache_runner.run_multi_signal_backtest(
        tickers=symbols,
        start_date=date(2024, 1, 1),
        end_date=date(2024, 1, 1),
        cache_root=cache_root,
        lookback_days=5,
        skip_days=1,
        costs_bps=3.0,
        entry_signals=["ts_momentum"],
        exit_signals=["none"],
    )

    leaderboard_dir_one = _latest_dir(output_root, "tsmom_multi_signal_*")
    leaderboard_csv = leaderboard_dir_one / "leaderboard.csv"
    leaderboard_json = leaderboard_dir_one / "leaderboard.json"
    assert leaderboard_csv.exists()
    assert leaderboard_json.exists()

    leaderboard_rows_one = list(csv.DictReader(leaderboard_csv.read_text().splitlines()))
    assert len(leaderboard_rows_one) == 1
    row_one = leaderboard_rows_one[0]
    for key in ["cagr", "sharpe", "max_drawdown", "turnover_total", "hit_rate"]:
        assert key in row_one

    run_dir_one = Path(row_one["run_dir"])
    metrics_rows = json.loads((run_dir_one / "metrics.json").read_text())
    metrics = {str(row["metric"]): float(row["value"]) for row in metrics_rows}
    for metric_name in ["cagr", "sharpe", "max_drawdown", "turnover_total"]:
        assert metric_name in metrics
    assert "hit_rate" in metrics or "win_rate" in metrics

    manifest = json.loads((run_dir_one / "manifest.json").read_text())
    assert manifest["config_hash"]
    assert manifest["code_commit_hash"]
    assert isinstance(manifest.get("dependency_versions"), dict)
    assert manifest["dependency_versions"]

    cache_runner.run_multi_signal_backtest(
        tickers=symbols,
        start_date=date(2024, 1, 1),
        end_date=date(2024, 1, 1),
        cache_root=cache_root,
        lookback_days=5,
        skip_days=1,
        costs_bps=3.0,
        entry_signals=["ts_momentum"],
        exit_signals=["none"],
    )
    leaderboard_dir_two = _latest_dir(output_root, "tsmom_multi_signal_*")
    rows_two = list(csv.DictReader((leaderboard_dir_two / "leaderboard.csv").read_text().splitlines()))
    assert len(rows_two) == 1
    row_two = rows_two[0]

    deterministic_fields = ["cagr", "sharpe", "max_drawdown", "turnover_total", "hit_rate", "total_return"]
    for field in deterministic_fields:
        assert float(row_two[field]) == pytest.approx(float(row_one[field]), rel=0.0, abs=1e-12)


def test_e2e_rejects_bad_schema_with_actionable_error(tmp_path: Path) -> None:
    cache_root = tmp_path / "cache"
    symbol = "BROKEN"
    symbol_dir = cache_root / symbol / "1m"
    symbol_dir.mkdir(parents=True, exist_ok=True)

    start = datetime(2024, 1, 1, 9, 30, tzinfo=timezone.utc)
    timestamps = np.array(
        [
            int((start + timedelta(minutes=2)).timestamp() * 1000),
            int((start + timedelta(minutes=1)).timestamp() * 1000),
        ],
        dtype=np.int64,
    )
    close = np.array([100.0, 100.5], dtype=float)
    np.savez_compressed(
        symbol_dir / f"{symbol}_1m_2024.npz",
        t=timestamps,
        o=close - 0.1,
        c=close,
    )

    with pytest.raises(ValueError, match="Non-monotonic timestamps detected for BROKEN"):
        load_canonical_price_arrays(
            symbols=[symbol],
            start="2024-01-01T09:30:00+00:00",
            end="2024-01-01T09:32:00+00:00",
            cache_root=cache_root,
        )

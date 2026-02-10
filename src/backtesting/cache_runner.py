from __future__ import annotations

import csv
import json
import random
import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime
from pathlib import Path

import numpy as np

from analysis.reporting import build_drawdown_rows, format_backtest_report
from backtesting.execution import BpsSlippage
from backtesting.vectorized import backtest_vectorized
from config import BACKTEST_CACHE_DIR, BACKTEST_OUTPUT_DIR
from data_access.api_client import MassiveApiClient
from data_access.cache import _safe_ticker_name

from data_access.engine_loader import EngineArrayBundle, load_canonical_price_arrays
from utils.parsing import build_npz_payload, chunk_results_by_year
from backtesting.signals import build_targets, parse_entry_signal_config, parse_exit_signal_config, required_lookback_window


def run_backtest_cache(
    tickers: list[str],
    start_date: date,
    end_date: date,
    cache_root: Path,
    api_key: str,
) -> str:
    api_client = MassiveApiClient(api_key)
    cache_root.mkdir(parents=True, exist_ok=True)
    BACKTEST_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    def _process_ticker(ticker: str) -> str:
        safe_ticker = _safe_ticker_name(ticker)
        ticker_dir = cache_root / safe_ticker / "1m"
        ticker_dir.mkdir(parents=True, exist_ok=True)
        index_path = ticker_dir / "index.json"
        expected_years = list(range(start_date.year, end_date.year + 1))
        try:
            cache_ready = False
            if index_path.exists():
                index_data = json.loads(index_path.read_text())
                years = index_data.get("years", [])
                cache_ready = (
                    index_data.get("full_range") is True
                    and set(expected_years).issubset(set(years))
                )
            if cache_ready:
                sample_text = f"{ticker}: cached data ready"
                sample_year = random.choice(expected_years)
                sample_path = ticker_dir / f"{safe_ticker}_1m_{sample_year}.npz"
                if sample_path.exists():
                    with np.load(sample_path, mmap_mode="r") as data:
                        if data["t"].size > 0:
                            idx = random.randrange(data["t"].size)
                            sample_text = (
                                f"{ticker}: sample close={data['c'][idx]} "
                                f"timestamp={int(data['t'][idx])}"
                            )
                return sample_text
            legacy_path = (
                cache_root
                / f"{safe_ticker}_1m_{start_date.isoformat()}_{end_date.isoformat()}.json"
            )
            if not legacy_path.exists():
                legacy_path = (
                    BACKTEST_CACHE_DIR
                    / f"{safe_ticker}_1m_{start_date.isoformat()}_{end_date.isoformat()}.json"
                )
            if legacy_path.exists():
                results = json.loads(legacy_path.read_text()).get("results", [])
            else:
                results = api_client.fetch_aggregates_range(
                    ticker, start_date, end_date, minutes_per_bar=1
                )
            buckets = chunk_results_by_year(results)
            for year, entries in buckets.items():
                payload = build_npz_payload(entries)
                np.savez_compressed(
                    ticker_dir / f"{safe_ticker}_1m_{year}.npz", **payload
                )
            index_payload = {
                "ticker": ticker,
                "start_date": start_date.isoformat(),
                "end_date": end_date.isoformat(),
                "full_range": True,
                "fetched_at": datetime.now().isoformat(),
                "years": sorted(buckets.keys()),
            }
            index_path.write_text(json.dumps(index_payload, indent=2))
            if results:
                sample = random.choice(results)
                return (
                    f"{ticker}: sample close={sample.get('c')} "
                    f"timestamp={sample.get('t')}"
                )
            return f"{ticker}: no data returned"
        except Exception as exc:
            return f"{ticker}: error fetching data ({exc})"

    lines: list[str] = []
    max_workers = min(8, max(1, len(tickers)))
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_map = {executor.submit(_process_ticker, ticker): ticker for ticker in tickers}
        for future in as_completed(future_map):
            lines.append(future.result())

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = BACKTEST_OUTPUT_DIR / f"backtest_cache_{timestamp}.txt"
    output_path.write_text("\n".join(lines))
    return "\n".join(lines) + f"\n\nSaved summary to: {output_path}"


def run_time_series_momentum_backtest(
    *,
    tickers: list[str],
    start_date: date,
    end_date: date,
    cache_root: Path,
    lookback_days: int,
    skip_days: int,
    costs_bps: float,
    entry_signal: str = "ts_momentum",
    entry_signal_params: dict[str, object] | None = None,
    exit_signal: str = "none",
    exit_signal_params: dict[str, object] | None = None,
) -> str:
    BACKTEST_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    entry_cfg = parse_entry_signal_config(
        entry_signal,
        entry_signal_params,
        default_lookback_days=lookback_days,
        default_skip_days=skip_days,
    )
    exit_cfg = parse_exit_signal_config(
        exit_signal,
        exit_signal_params,
        default_lookback_days=lookback_days,
        default_skip_days=skip_days,
    )
    start_dt = datetime.combine(start_date, datetime.min.time())
    end_dt = datetime.combine(end_date, datetime.max.time())
    lookback_window = required_lookback_window(entry_cfg, exit_cfg)
    arrays = load_backtest_engine_arrays(
        tickers=tickers,
        start=start_dt.isoformat(),
        end=end_dt.isoformat(),
        cache_root=cache_root,
        timeframe="1m",
        lookback_window=lookback_window,
    )

    prices = _fill_missing_prices(arrays.close_prices)
    signals = build_targets(
        close_prices=prices,
        missing_mask=arrays.missing_mask,
        entry_config=entry_cfg,
        exit_config=exit_cfg,
    )

    result = backtest_vectorized(
        prices=prices,
        signals=signals,
        slippage_model=BpsSlippage(costs_bps),
        initial_equity=1.0,
    )

    timestamps = arrays.date_index
    symbol_order = [
        symbol
        for symbol, _idx in sorted(
            arrays.metadata.symbol_to_column.items(), key=lambda item: item[1]
        )
    ]

    equity = _to_numpy_1d(result.equity_curve)
    returns = _to_numpy_1d(result.returns)
    turnover = _to_numpy_1d(result.turnover)
    trades = _to_numpy_2d(result.trades)

    drawdown_rows = build_drawdown_rows(timestamps, equity)
    turnover_stats = {
        "mean": float(np.mean(turnover)) if turnover.size else 0.0,
        "total": float(np.sum(turnover)) if turnover.size else 0.0,
        "max": float(np.max(turnover)) if turnover.size else 0.0,
    }
    cost_totals = {
        key: float(value)
        for key, value in result.cost_breakdown.get("totals", {}).items()
    }

    metrics = dict(result.metrics)
    metrics["turnover_total"] = turnover_stats["total"]
    metrics["cost_total"] = cost_totals.get("total", 0.0)

    run_dir = _persist_backtest_outputs(
        timestamps=timestamps,
        symbol_order=symbol_order,
        equity=equity,
        returns=returns,
        trades=trades,
        metrics=metrics,
    )

    report_text = format_backtest_report(
        title="Time-Series Momentum Backtest",
        params={
            "tickers": ", ".join(tickers),
            "lookback_days": lookback_days,
            "skip_days": skip_days,
            "costs_bps": costs_bps,
            "entry_signal": entry_signal,
            "entry_signal_params": json.dumps(entry_signal_params or {}),
            "exit_signal": exit_signal,
            "exit_signal_params": json.dumps(exit_signal_params or {}),
            "start_date": start_date.isoformat(),
            "end_date": end_date.isoformat(),
        },
        metrics=metrics,
        drawdown_rows=drawdown_rows,
        turnover_stats=turnover_stats,
        cost_totals=cost_totals,
    )
    (run_dir / "report.txt").write_text(report_text)

    return report_text + f"\n\nSaved outputs to: {run_dir}"



def load_backtest_engine_arrays(
    tickers: list[str],
    start: datetime | str,
    end: datetime | str,
    *,
    cache_root: Path | None = None,
    timeframe: str = "1m",
    lookback_window: int = 0,
) -> EngineArrayBundle:
    """Load canonical float64 arrays and metadata for backtest engines."""

    return load_canonical_price_arrays(
        symbols=tickers,
        start=start,
        end=end,
        cache_root=cache_root,
        timeframe=timeframe,
        lookback_window=lookback_window,
        validate_split_adjustment=True,
    )


def _fill_missing_prices(close_prices: np.ndarray) -> np.ndarray:
    values = np.asarray(close_prices, dtype=float).copy()
    for col in range(values.shape[1]):
        column = values[:, col]
        valid_idx = np.flatnonzero(np.isfinite(column))
        if valid_idx.size == 0:
            values[:, col] = 1.0
            continue
        first = int(valid_idx[0])
        column[:first] = column[first]
        for idx in range(first + 1, column.size):
            if not np.isfinite(column[idx]):
                column[idx] = column[idx - 1]
    return values


def _to_numpy_1d(values: object) -> np.ndarray:
    if hasattr(values, "to_numpy"):
        return np.asarray(values.to_numpy(), dtype=float)
    arr = np.asarray(values, dtype=float)
    if arr.ndim != 1:
        return np.asarray(arr.reshape(arr.shape[0]), dtype=float)
    return arr


def _to_numpy_2d(values: object) -> np.ndarray:
    if hasattr(values, "to_numpy"):
        arr = np.asarray(values.to_numpy(), dtype=float)
    else:
        arr = np.asarray(values, dtype=float)
    if arr.ndim == 1:
        return arr.reshape(-1, 1)
    return arr


def _persist_backtest_outputs(
    *,
    timestamps: np.ndarray,
    symbol_order: list[str],
    equity: np.ndarray,
    returns: np.ndarray,
    trades: np.ndarray,
    metrics: dict[str, float],
) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir = BACKTEST_OUTPUT_DIR / f"tsmom_backtest_{timestamp}"
    run_dir.mkdir(parents=True, exist_ok=True)

    time_strings = [datetime.utcfromtimestamp(int(ts) / 1000.0).isoformat() for ts in timestamps]

    _write_series_csv_json(
        run_dir=run_dir,
        stem="equity",
        field_name="equity",
        timestamps=time_strings,
        values=equity,
    )
    _write_series_csv_json(
        run_dir=run_dir,
        stem="returns",
        field_name="returns",
        timestamps=time_strings,
        values=returns,
    )

    _write_trades_csv_json(
        run_dir=run_dir,
        timestamps=time_strings,
        symbol_order=symbol_order,
        trades=trades,
    )

    metrics_rows = [{"metric": key, "value": float(value)} for key, value in metrics.items()]
    metrics_csv = run_dir / "metrics.csv"
    with metrics_csv.open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["metric", "value"])
        writer.writeheader()
        writer.writerows(metrics_rows)
    (run_dir / "metrics.json").write_text(json.dumps(metrics_rows, indent=2))

    return run_dir


def _write_series_csv_json(
    *,
    run_dir: Path,
    stem: str,
    field_name: str,
    timestamps: list[str],
    values: np.ndarray,
) -> None:
    rows = [
        {"timestamp": ts, field_name: float(value)}
        for ts, value in zip(timestamps, values, strict=False)
    ]
    csv_path = run_dir / f"{stem}.csv"
    with csv_path.open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["timestamp", field_name])
        writer.writeheader()
        writer.writerows(rows)
    (run_dir / f"{stem}.json").write_text(json.dumps(rows, indent=2))


def _write_trades_csv_json(
    *,
    run_dir: Path,
    timestamps: list[str],
    symbol_order: list[str],
    trades: np.ndarray,
) -> None:
    rows: list[dict[str, object]] = []
    for row_idx, ts in enumerate(timestamps):
        for col_idx, symbol in enumerate(symbol_order):
            rows.append(
                {
                    "timestamp": ts,
                    "symbol": symbol,
                    "trade": float(trades[row_idx, col_idx]),
                }
            )

    csv_path = run_dir / "trades.csv"
    with csv_path.open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["timestamp", "symbol", "trade"])
        writer.writeheader()
        writer.writerows(rows)
    (run_dir / "trades.json").write_text(json.dumps(rows, indent=2))


def _parse_json_object(raw: str | None) -> dict[str, object]:
    if raw is None or not raw.strip():
        return {}
    parsed = json.loads(raw)
    if not isinstance(parsed, dict):
        raise ValueError("signal params must be a JSON object")
    return parsed


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run the time-series momentum backtest with pluggable entry/exit signals.",
        epilog=(
            "Examples:\n"
            "  PYTHONPATH=src python -m backtesting.cache_runner --tickers AAPL,MSFT "
            "--start-date 2024-01-01 --end-date 2024-03-01 --entry-signal ts_momentum --exit-signal none\n"
            "  PYTHONPATH=src python -m backtesting.cache_runner --tickers AAPL --start-date 2024-01-01 "
            "--end-date 2024-03-01 --entry-signal breakout --entry-signal-params '{\"breakout_window\": 55}' "
            "--exit-signal trailing_stop --exit-signal-params '{\"trailing_stop_pct\": 0.08}'"
        ),
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument("--tickers", required=True, help="Comma-separated ticker list.")
    parser.add_argument("--start-date", required=True, help="Backtest start date (YYYY-MM-DD).")
    parser.add_argument("--end-date", required=True, help="Backtest end date (YYYY-MM-DD).")
    parser.add_argument("--cache-root", default=str(BACKTEST_CACHE_DIR), help="Cache root directory.")
    parser.add_argument("--lookback-days", type=int, default=90)
    parser.add_argument("--skip-days", type=int, default=5)
    parser.add_argument("--costs-bps", type=float, default=5.0)
    parser.add_argument("--entry-signal", default="ts_momentum", choices=["ts_momentum", "ma_trend", "breakout"])
    parser.add_argument("--entry-signal-params", default="{}", help="JSON object with entry signal parameters.")
    parser.add_argument("--exit-signal", default="none", choices=["none", "momentum_flip", "trailing_stop", "max_hold"])
    parser.add_argument("--exit-signal-params", default="{}", help="JSON object with exit signal parameters.")
    return parser


def main() -> None:
    parser = _build_parser()
    args = parser.parse_args()
    output = run_time_series_momentum_backtest(
        tickers=[part.strip() for part in args.tickers.split(",") if part.strip()],
        start_date=date.fromisoformat(args.start_date),
        end_date=date.fromisoformat(args.end_date),
        cache_root=Path(args.cache_root),
        lookback_days=args.lookback_days,
        skip_days=args.skip_days,
        costs_bps=args.costs_bps,
        entry_signal=args.entry_signal,
        entry_signal_params=_parse_json_object(args.entry_signal_params),
        exit_signal=args.exit_signal,
        exit_signal_params=_parse_json_object(args.exit_signal_params),
    )
    print(output)


if __name__ == "__main__":
    main()

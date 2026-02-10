from __future__ import annotations

import csv
import json
import random
import argparse
import logging
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor, as_completed
from datetime import date, datetime
from itertools import product
from pathlib import Path
from typing import Any, Callable

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


LOGGER = logging.getLogger(__name__)


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

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
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
    signal_rebalance_interval: int = 1,
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
    signals = _throttle_signal_changes(signals, interval=max(1, int(signal_rebalance_interval)))

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

    trade_log_rows = _build_trade_log_rows(
        timestamps=timestamps,
        symbol_order=symbol_order,
        prices=prices,
        trades=trades,
        costs_bps=costs_bps,
    )
    trade_log_summary = _format_trade_log_summary(trade_log_rows)

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
            "signal_rebalance_interval": signal_rebalance_interval,
        },
        metrics=metrics,
        drawdown_rows=drawdown_rows,
        turnover_stats=turnover_stats,
        cost_totals=cost_totals,
    )
    (run_dir / "trade_log.csv").write_text(_trade_log_csv(trade_log_rows))
    (run_dir / "trade_log.json").write_text(json.dumps(trade_log_rows, indent=2))

    final_report = report_text + "\n\n" + trade_log_summary
    (run_dir / "report.txt").write_text(final_report)

    return final_report + f"\n\nSaved outputs to: {run_dir}"


def generate_sweep_combinations(
    *,
    entry_grid: dict[str, list[dict[str, Any]]],
    exit_grid: dict[str, list[dict[str, Any]]],
    core_grid: dict[str, list[Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Build Cartesian combinations for entry/exit/core sweep definitions."""

    valid: list[dict[str, Any]] = []
    invalid: list[dict[str, Any]] = []

    normalized_core = {key: list(values or []) for key, values in core_grid.items()}
    core_keys = sorted(normalized_core)
    core_values = [normalized_core[key] for key in core_keys]
    if any(len(values) == 0 for values in core_values):
        return valid, [{"reason": "core_grid contains empty value list", "core_grid": core_grid}]

    for entry_signal, entry_params_grid in entry_grid.items():
        for exit_signal, exit_params_grid in exit_grid.items():
            for entry_params in entry_params_grid or []:
                for exit_params in exit_params_grid or []:
                    for core_combo in product(*core_values):
                        core_params = dict(zip(core_keys, core_combo, strict=True))
                        combo = {
                            "entry_signal": entry_signal,
                            "entry_signal_params": dict(entry_params),
                            "exit_signal": exit_signal,
                            "exit_signal_params": dict(exit_params),
                            **core_params,
                        }
                        if _is_valid_combo_definition(combo):
                            valid.append(combo)
                        else:
                            invalid.append({"reason": "invalid combo parameters", "combo": combo})
    return valid, invalid


def run_parameter_sweep(
    *,
    tickers: list[str],
    start_date: date,
    end_date: date,
    cache_root: Path,
    entry_grid: dict[str, list[dict[str, Any]]],
    exit_grid: dict[str, list[dict[str, Any]]],
    core_grid: dict[str, list[Any]],
    seed: int = 42,
    max_workers: int | None = None,
    fail_fast: bool = False,
    continue_on_error: bool = True,
    top_n: int = 10,
    evaluator: Callable[[dict[str, Any]], dict[str, Any]] | None = None,
) -> str:
    """Run a parallel sweep over signal/core-parameter combinations."""

    if fail_fast and continue_on_error:
        raise ValueError("fail_fast and continue_on_error cannot both be true")

    combos, invalid_rows = generate_sweep_combinations(
        entry_grid=entry_grid,
        exit_grid=exit_grid,
        core_grid=core_grid,
    )
    if not combos:
        raise ValueError("No valid combinations generated for sweep")

    random.Random(seed).shuffle(combos)
    worker = evaluator or _execute_sweep_combo
    default_workers = min(8, max(1, len(combos)))
    n_workers = max_workers or default_workers
    use_process_pool = evaluator is None
    executor_cls = ProcessPoolExecutor if use_process_pool else ThreadPoolExecutor

    rows: list[dict[str, Any]] = []
    errors: list[dict[str, Any]] = []
    LOGGER.info("Starting sweep for %s combinations with %s workers", len(combos), n_workers)

    with executor_cls(max_workers=n_workers) as executor:
        futures = {
            executor.submit(
                worker,
                {
                    "combo_index": idx,
                    "seed": seed,
                    "tickers": tickers,
                    "start_date": start_date.isoformat(),
                    "end_date": end_date.isoformat(),
                    "cache_root": str(cache_root),
                    **combo,
                },
            ): combo
            for idx, combo in enumerate(combos)
        }
        completed = 0
        for future in as_completed(futures):
            combo = futures[future]
            completed += 1
            try:
                rows.append(future.result())
            except Exception as exc:
                error_row = {"error": str(exc), "combo": combo}
                errors.append(error_row)
                LOGGER.exception("Sweep combo failed (%s/%s)", completed, len(combos))
                if fail_fast:
                    raise
                if not continue_on_error:
                    raise
            LOGGER.info("Sweep progress: %s/%s", completed, len(combos))

    ranked_rows = sorted(rows, key=lambda row: float(row["sharpe"]), reverse=True)
    run_dir = _persist_sweep_outputs(
        ranked_rows=ranked_rows,
        invalid_rows=invalid_rows,
        errors=errors,
        top_n=top_n,
    )
    return (
        f"Sweep complete: {len(ranked_rows)} successful combos, "
        f"{len(invalid_rows)} skipped, {len(errors)} failed. "
        f"Saved outputs to: {run_dir}"
    )


def _execute_sweep_combo(payload: dict[str, Any]) -> dict[str, Any]:
    combo_seed = int(payload["seed"]) + int(payload["combo_index"])
    random.seed(combo_seed)
    np.random.seed(combo_seed)

    tickers = [str(ticker) for ticker in payload["tickers"]]
    start_date = date.fromisoformat(str(payload["start_date"]))
    end_date = date.fromisoformat(str(payload["end_date"]))
    lookback_days = int(payload["lookback_days"])
    skip_days = int(payload["skip_days"])
    costs_bps = float(payload["costs_bps"])

    entry_cfg = parse_entry_signal_config(
        str(payload["entry_signal"]),
        dict(payload.get("entry_signal_params", {})),
        default_lookback_days=lookback_days,
        default_skip_days=skip_days,
    )
    exit_cfg = parse_exit_signal_config(
        str(payload["exit_signal"]),
        dict(payload.get("exit_signal_params", {})),
        default_lookback_days=lookback_days,
        default_skip_days=skip_days,
    )

    arrays = load_backtest_engine_arrays(
        tickers=tickers,
        start=datetime.combine(start_date, datetime.min.time()).isoformat(),
        end=datetime.combine(end_date, datetime.max.time()).isoformat(),
        cache_root=Path(str(payload["cache_root"])),
        timeframe="1m",
        lookback_window=required_lookback_window(entry_cfg, exit_cfg),
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

    turnover = _to_numpy_1d(result.turnover)
    cost_totals = {
        key: float(value)
        for key, value in result.cost_breakdown.get("totals", {}).items()
    }
    metrics = dict(result.metrics)
    return {
        "entry_signal": payload["entry_signal"],
        "entry_signal_params": json.dumps(payload.get("entry_signal_params", {}), sort_keys=True),
        "exit_signal": payload["exit_signal"],
        "exit_signal_params": json.dumps(payload.get("exit_signal_params", {}), sort_keys=True),
        "lookback_days": lookback_days,
        "skip_days": skip_days,
        "costs_bps": costs_bps,
        "total_return": float(metrics.get("total_return", 0.0)),
        "sharpe": float(metrics.get("sharpe", 0.0)),
        "max_drawdown": float(metrics.get("max_drawdown", 0.0)),
        "volatility": float(metrics.get("volatility", 0.0)),
        "turnover_total": float(np.sum(turnover)) if turnover.size else 0.0,
        "cost_total": cost_totals.get("total", 0.0),
    }


def _persist_sweep_outputs(
    *,
    ranked_rows: list[dict[str, Any]],
    invalid_rows: list[dict[str, Any]],
    errors: list[dict[str, Any]],
    top_n: int,
) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    run_dir = BACKTEST_OUTPUT_DIR / f"tsmom_sweep_{timestamp}"
    run_dir.mkdir(parents=True, exist_ok=True)

    leaderboard_csv = run_dir / "leaderboard.csv"
    fieldnames = list(ranked_rows[0].keys()) if ranked_rows else []
    with leaderboard_csv.open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(ranked_rows)
    (run_dir / "leaderboard.json").write_text(json.dumps(ranked_rows, indent=2))

    (run_dir / "per_combo_summary.csv").write_text(leaderboard_csv.read_text())
    (run_dir / "per_combo_summary.json").write_text(json.dumps(ranked_rows, indent=2))
    (run_dir / "skipped_invalid_combos.json").write_text(json.dumps(invalid_rows, indent=2))
    (run_dir / "errors.json").write_text(json.dumps(errors, indent=2))

    top_rows = ranked_rows[: max(0, top_n)]
    report_lines = ["Top sweep combinations", "======================", ""]
    for idx, row in enumerate(top_rows, start=1):
        report_lines.append(
            f"#{idx}: sharpe={row['sharpe']:.6f} total_return={row['total_return']:.6f} "
            f"entry={row['entry_signal']} exit={row['exit_signal']} "
            f"core=(lookback_days={row['lookback_days']}, skip_days={row['skip_days']}, costs_bps={row['costs_bps']})"
        )
    (run_dir / "top_n_report.txt").write_text("\n".join(report_lines))

    return run_dir


def _is_valid_combo_definition(combo: dict[str, Any]) -> bool:
    try:
        parse_entry_signal_config(
            str(combo["entry_signal"]),
            dict(combo.get("entry_signal_params", {})),
            default_lookback_days=int(combo["lookback_days"]),
            default_skip_days=int(combo["skip_days"]),
        )
        parse_exit_signal_config(
            str(combo["exit_signal"]),
            dict(combo.get("exit_signal_params", {})),
            default_lookback_days=int(combo["lookback_days"]),
            default_skip_days=int(combo["skip_days"]),
        )
        float(combo["costs_bps"])
        return True
    except Exception:
        return False





def run_multi_signal_backtest(
    *,
    tickers: list[str],
    start_date: date,
    end_date: date,
    cache_root: Path,
    lookback_days: int,
    skip_days: int,
    costs_bps: float,
    entry_signals: list[str],
    exit_signals: list[str],
) -> str:
    """Run all selected entry/exit combinations with shared core parameters."""

    if not entry_signals:
        raise ValueError("At least one entry signal must be selected")
    if not exit_signals:
        raise ValueError("At least one exit signal must be selected")

    rows: list[dict[str, Any]] = []
    combo_reports: list[str] = []

    for entry_signal in entry_signals:
        for exit_signal in exit_signals:
            report = run_time_series_momentum_backtest(
                tickers=tickers,
                start_date=start_date,
                end_date=end_date,
                cache_root=cache_root,
                lookback_days=lookback_days,
                skip_days=skip_days,
                costs_bps=costs_bps,
                signal_rebalance_interval=390,
                entry_signal=entry_signal,
                entry_signal_params={
                    "min_abs_return": 0.01,
                    "long_only": True,
                } if entry_signal == "ts_momentum" else {},
                exit_signal=exit_signal,
                exit_signal_params={
                    "min_abs_return": 0.01,
                } if exit_signal == "momentum_flip" else {},
            )
            run_dir = _extract_saved_output_dir(report)
            metrics = _load_metrics_from_run_dir(run_dir)
            rows.append(
                {
                    "entry_signal": entry_signal,
                    "exit_signal": exit_signal,
                    "total_return": float(metrics.get("total_return", 0.0)),
                    "sharpe": float(metrics.get("sharpe", 0.0)),
                    "max_drawdown": float(metrics.get("max_drawdown", 0.0)),
                    "volatility": float(metrics.get("volatility", 0.0)),
                    "turnover_total": float(metrics.get("turnover_total", 0.0)),
                    "cost_total": float(metrics.get("cost_total", 0.0)),
                    "run_dir": str(run_dir),
                }
            )
            combo_reports.append(
                f"entry={entry_signal} exit={exit_signal}\n{report}\n"
            )

    ranked_rows = sorted(rows, key=lambda row: float(row["sharpe"]), reverse=True)
    leaderboard_dir = _persist_multi_signal_outputs(ranked_rows)

    summary_lines = [
        "Multi-signal backtest completed.",
        f"Combinations: {len(rows)}",
        f"Leaderboard outputs: {leaderboard_dir}",
        "",
        "Ranked combinations (by sharpe):",
    ]
    for idx, row in enumerate(ranked_rows, start=1):
        summary_lines.append(
            f"#{idx} entry={row['entry_signal']} exit={row['exit_signal']} "
            f"sharpe={row['sharpe']:.6f} total_return={row['total_return']:.6f} "
            f"max_drawdown={row['max_drawdown']:.6f} volatility={row['volatility']:.6f} "
            f"turnover_total={row['turnover_total']:.6f} cost_total={row['cost_total']:.6f}"
        )

    return "\n".join(summary_lines + ["", "Detailed combo reports:", "", *combo_reports])


def _extract_saved_output_dir(report_text: str) -> Path:
    marker = "Saved outputs to: "
    idx = report_text.rfind(marker)
    if idx < 0:
        raise ValueError("Could not locate saved output directory in report text")
    raw = report_text[idx + len(marker):].strip().splitlines()[0].strip()
    return Path(raw)


def _load_metrics_from_run_dir(run_dir: Path) -> dict[str, float]:
    metrics_path = run_dir / "metrics.json"
    rows = json.loads(metrics_path.read_text())
    metrics: dict[str, float] = {}
    for row in rows:
        metric = str(row.get("metric", ""))
        value = float(row.get("value", 0.0))
        metrics[metric] = value
    return metrics


def _persist_multi_signal_outputs(rows: list[dict[str, Any]]) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    run_dir = BACKTEST_OUTPUT_DIR / f"tsmom_multi_signal_{timestamp}"
    run_dir.mkdir(parents=True, exist_ok=True)

    fieldnames = [
        "entry_signal",
        "exit_signal",
        "total_return",
        "sharpe",
        "max_drawdown",
        "volatility",
        "turnover_total",
        "cost_total",
        "run_dir",
    ]
    with (run_dir / "leaderboard.csv").open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    (run_dir / "leaderboard.json").write_text(json.dumps(rows, indent=2))
    return run_dir
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


def _throttle_signal_changes(signals: np.ndarray, *, interval: int) -> np.ndarray:
    if interval <= 1:
        return np.asarray(signals, dtype=float)
    values = np.asarray(signals, dtype=float)
    throttled = np.zeros_like(values)
    throttled[0] = values[0]
    for idx in range(1, values.shape[0]):
        if idx % interval == 0:
            throttled[idx] = values[idx]
        else:
            throttled[idx] = throttled[idx - 1]
    return throttled


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




def _build_trade_log_rows(
    *,
    timestamps: np.ndarray,
    symbol_order: list[str],
    prices: np.ndarray,
    trades: np.ndarray,
    costs_bps: float,
) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for col_idx, symbol in enumerate(symbol_order):
        side = 0
        entry_price = 0.0
        running_pnl = 0.0
        for row_idx, ts in enumerate(timestamps):
            trade = float(trades[row_idx, col_idx])
            if trade == 0.0:
                continue
            price = float(prices[row_idx, col_idx])
            event = "adjust"
            trade_pnl = 0.0
            trade_cost = abs(trade) * (costs_bps / 10_000.0)
            if side == 0 and abs(trade) > 0:
                side = 1 if trade > 0 else -1
                entry_price = price
                event = "entry"
            elif side != 0 and side * trade < 0:
                trade_pnl = ((price - entry_price) / entry_price) * side if entry_price > 0 else 0.0
                running_pnl += trade_pnl - trade_cost
                event = "exit" if abs(trade) == abs(side) else "flip"
                next_side = side + int(round(trade))
                side = next_side
                if side != 0:
                    entry_price = price
                else:
                    entry_price = 0.0
            rows.append(
                {
                    "timestamp": datetime.utcfromtimestamp(int(ts) / 1000.0).isoformat(),
                    "symbol": symbol,
                    "event": event,
                    "trade": trade,
                    "price": price,
                    "trade_pnl": float(trade_pnl),
                    "trade_cost": float(trade_cost),
                    "running_pnl": float(running_pnl),
                }
            )
    return rows


def _format_trade_log_summary(rows: list[dict[str, object]], max_rows: int = 40) -> str:
    lines = ["Trade Log", "---------"]
    if not rows:
        lines.append("No trade events generated.")
        return "\n".join(lines)
    lines.append("timestamp | symbol | event | trade | price | trade_pnl | trade_cost | running_pnl")
    for row in rows[:max_rows]:
        lines.append(
            f"{row['timestamp']} | {row['symbol']} | {row['event']} | {float(row['trade']):.2f} | "
            f"{float(row['price']):.4f} | {float(row['trade_pnl']):.6f} | {float(row['trade_cost']):.6f} | {float(row['running_pnl']):.6f}"
        )
    if len(rows) > max_rows:
        lines.append(f"... ({len(rows) - max_rows} more trade events)")
    return "\n".join(lines)


def _trade_log_csv(rows: list[dict[str, object]]) -> str:
    header = "timestamp,symbol,event,trade,price,trade_pnl,trade_cost,running_pnl"
    body = [
        f"{row['timestamp']},{row['symbol']},{row['event']},{row['trade']},{row['price']},{row['trade_pnl']},{row['trade_cost']},{row['running_pnl']}"
        for row in rows
    ]
    return "\n".join([header, *body]) + "\n"

def _persist_backtest_outputs(
    *,
    timestamps: np.ndarray,
    symbol_order: list[str],
    equity: np.ndarray,
    returns: np.ndarray,
    trades: np.ndarray,
    metrics: dict[str, float],
) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
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


def _add_common_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--tickers", required=True, help="Comma-separated ticker list.")
    parser.add_argument("--start-date", required=True, help="Backtest start date (YYYY-MM-DD).")
    parser.add_argument("--end-date", required=True, help="Backtest end date (YYYY-MM-DD).")
    parser.add_argument("--cache-root", default=str(BACKTEST_CACHE_DIR), help="Cache root directory.")


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Backtest runner and parameter sweep tools.")
    subparsers = parser.add_subparsers(dest="command")

    run_parser = subparsers.add_parser("run", help="Run one backtest combo.")
    _add_common_args(run_parser)
    run_parser.add_argument("--lookback-days", type=int, default=90)
    run_parser.add_argument("--skip-days", type=int, default=5)
    run_parser.add_argument("--costs-bps", type=float, default=5.0)
    run_parser.add_argument("--entry-signal", default="ts_momentum", choices=["ts_momentum", "ma_trend", "breakout"])
    run_parser.add_argument("--entry-signal-params", default="{}", help="JSON object with entry signal parameters.")
    run_parser.add_argument("--exit-signal", default="none", choices=["none", "momentum_flip", "trailing_stop", "max_hold"])
    run_parser.add_argument("--exit-signal-params", default="{}", help="JSON object with exit signal parameters.")
    run_parser.add_argument("--signal-rebalance-interval", type=int, default=1, help="Only allow signal changes every N bars.")

    sweep_parser = subparsers.add_parser("sweep", help="Run parameter sweep across signal/core grids.")
    _add_common_args(sweep_parser)
    sweep_parser.add_argument("--entry-grid", required=True, help="JSON mapping signal->list[params].")
    sweep_parser.add_argument("--exit-grid", required=True, help="JSON mapping signal->list[params].")
    sweep_parser.add_argument("--core-grid", required=True, help="JSON mapping core param->list[values].")
    sweep_parser.add_argument("--seed", type=int, default=42)
    sweep_parser.add_argument("--max-workers", type=int, default=None)
    sweep_parser.add_argument("--top-n", type=int, default=10)
    sweep_parser.add_argument("--fail-fast", action="store_true")
    sweep_parser.add_argument("--continue-on-error", action="store_true")

    parser.set_defaults(command="run")
    return parser


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    parser = _build_parser()
    args = parser.parse_args()

    tickers = [part.strip() for part in args.tickers.split(",") if part.strip()]
    start_date = date.fromisoformat(args.start_date)
    end_date = date.fromisoformat(args.end_date)
    cache_root = Path(args.cache_root)

    if args.command == "sweep":
        output = run_parameter_sweep(
            tickers=tickers,
            start_date=start_date,
            end_date=end_date,
            cache_root=cache_root,
            entry_grid={str(k): list(v) for k, v in _parse_json_object(args.entry_grid).items()},
            exit_grid={str(k): list(v) for k, v in _parse_json_object(args.exit_grid).items()},
            core_grid={str(k): list(v) for k, v in _parse_json_object(args.core_grid).items()},
            seed=args.seed,
            max_workers=args.max_workers,
            fail_fast=bool(args.fail_fast),
            continue_on_error=bool(args.continue_on_error),
            top_n=args.top_n,
        )
    else:
        output = run_time_series_momentum_backtest(
            tickers=tickers,
            start_date=start_date,
            end_date=end_date,
            cache_root=cache_root,
            lookback_days=args.lookback_days,
            skip_days=args.skip_days,
            costs_bps=args.costs_bps,
            entry_signal=args.entry_signal,
            entry_signal_params=_parse_json_object(args.entry_signal_params),
            exit_signal=args.exit_signal,
            exit_signal_params=_parse_json_object(args.exit_signal_params),
            signal_rebalance_interval=args.signal_rebalance_interval,
        )
    print(output)


if __name__ == "__main__":
    main()

from __future__ import annotations

import csv
import json
import random
import argparse
import logging
import hashlib
import os
import platform
import socket
import subprocess
import sys
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor, as_completed
from datetime import date, datetime
from itertools import product
from pathlib import Path
from typing import Any, Callable

import numpy as np

from analysis.reporting import (
    build_backtest_robustness_report,
    build_drawdown_rows,
    build_sweep_robustness_report,
    format_backtest_report,
)
from backtesting.execution import (
    AssetClassCarryCost,
    BpsSlippage,
    ParticipationImpactSlippage,
    ShortBorrowCost,
    SpreadSlippage,
    VolatilityScaledSlippage,
)
from backtesting.vectorized import backtest_vectorized
from config import BACKTEST_CACHE_DIR, BACKTEST_OUTPUT_DIR
from data_access.api_client import MassiveApiClient
from data_access.cache import _safe_ticker_name

from data_access.engine_loader import (
    EngineArrayBundle,
    EngineArrayMetadata,
    load_canonical_price_arrays,
    validate_engine_dataset_contracts,
)
from utils.parsing import build_npz_payload, chunk_results_by_year
from backtesting.signals import build_targets, parse_entry_signal_config, parse_exit_signal_config, required_lookback_window
from backtesting.portfolio import PortfolioConstructionConfig, construct_target_weights
from backtesting.strategies import CrossSectionalMomentumConfig, build_cross_sectional_momentum_targets
from backtesting.walk_forward import (
    build_walk_forward_folds,
    persist_walk_forward_outputs,
    run_walk_forward_optimization,
)
from backtesting.optimization import (
    Constraint,
    Objective,
    RandomSampler,
    TPESampler,
    optimize,
)


LOGGER = logging.getLogger(__name__)
CANONICAL_METRIC_SCHEMA_VERSION = "1.0"


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
    execution_model: str = "bps",
    execution_model_params: dict[str, object] | None = None,
    carry_model: str = "short_borrow",
    carry_model_params: dict[str, object] | None = None,
    entry_signal: str = "ts_momentum",
    entry_signal_params: dict[str, object] | None = None,
    exit_signal: str = "none",
    exit_signal_params: dict[str, object] | None = None,
    signal_rebalance_interval: int = 1,
    starting_capital: float = 100_000.0,
    bet_sizing_mode: str = "half_kelly",
    custom_bet_pct: float = 10.0,
    strategy: str = "momentum",
    xsmom_top_quantile: float = 0.2,
    xsmom_bottom_quantile: float = 0.2,
    xsmom_long_only: bool = False,
    xsmom_vol_lookback_days: int = 20,
    timeframe: str = "1m",
    portfolio_method: str = "equal_weight",
    portfolio_vol_lookback_bars: int = 20,
    portfolio_target_volatility: float = 0.10,
    portfolio_max_symbol_weight: float = 0.25,
    portfolio_max_sector_weight: float = 0.60,
    portfolio_max_gross_exposure: float = 1.0,
    portfolio_min_net_exposure: float = -1.0,
    portfolio_max_net_exposure: float = 1.0,
    portfolio_sector_map: dict[str, str] | None = None,
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
    if strategy == "xsmom":
        lookback_window = max(lookback_window, lookback_days + skip_days + 1, int(xsmom_vol_lookback_days) + 2)
    arrays = load_backtest_engine_arrays(
        tickers=tickers,
        start=start_dt.isoformat(),
        end=end_dt.isoformat(),
        cache_root=cache_root,
        timeframe="1m",
        lookback_window=lookback_window,
    )
    arrays = _resample_engine_bundle_from_1m(arrays, timeframe=timeframe)
    dataset_contracts = validate_engine_dataset_contracts(arrays)

    prices = _fill_missing_prices(arrays.close_prices)
    bet_fraction = _resolve_bet_fraction(
        prices=prices,
        mode=bet_sizing_mode,
        custom_bet_pct=custom_bet_pct,
    )

    if strategy == "xsmom":
        xs_cfg = CrossSectionalMomentumConfig(
            lookback_days=lookback_days,
            skip_days=skip_days,
            top_quantile=float(xsmom_top_quantile),
            bottom_quantile=float(xsmom_bottom_quantile),
            long_only=bool(xsmom_long_only),
            vol_lookback_days=int(xsmom_vol_lookback_days),
            rebalance_interval=max(1, int(signal_rebalance_interval)),
        )
        signals = build_cross_sectional_momentum_targets(
            close_prices=prices,
            missing_mask=arrays.missing_mask,
            config=xs_cfg,
        )
        sized_signals = signals * float(bet_fraction)
    else:
        signals = build_targets(
            close_prices=prices,
            missing_mask=arrays.missing_mask,
            entry_config=entry_cfg,
            exit_config=exit_cfg,
        )
        signals = _throttle_signal_changes(signals, interval=max(1, int(signal_rebalance_interval)))
        sized_signals = _apply_discrete_bet_sizing(
            signals=signals,
            prices=prices,
            starting_capital=starting_capital,
            bet_fraction=bet_fraction,
        )

    slippage = _build_slippage_model(
        model_name=execution_model,
        costs_bps=costs_bps,
        params=execution_model_params,
    )
    borrow = _build_carry_model(
        model_name=carry_model,
        params=carry_model_params,
        n_assets=prices.shape[1],
        timeframe=timeframe,
    )

    portfolio_cfg = PortfolioConstructionConfig(
        method=str(portfolio_method),
        vol_lookback_bars=int(portfolio_vol_lookback_bars),
        target_volatility=float(portfolio_target_volatility),
        max_symbol_weight=float(portfolio_max_symbol_weight),
        max_sector_weight=float(portfolio_max_sector_weight),
        max_gross_exposure=float(portfolio_max_gross_exposure),
        min_net_exposure=float(portfolio_min_net_exposure),
        max_net_exposure=float(portfolio_max_net_exposure),
        sector_map=dict(portfolio_sector_map or {}),
    )

    symbol_order = [
        symbol
        for symbol, _idx in sorted(
            arrays.metadata.symbol_to_column.items(), key=lambda item: item[1]
        )
    ]

    portfolio_result = construct_target_weights(
        raw_signals=sized_signals,
        prices=prices,
        symbol_order=symbol_order,
        config=portfolio_cfg,
    )

    borrow_rate_series = None
    borrow_available_flags = None
    if isinstance(carry_model_params, dict):
        if "borrow_rate_series" in carry_model_params:
            borrow_rate_series = np.asarray(carry_model_params.get("borrow_rate_series"), dtype=float)
        if "borrow_available_flags" in carry_model_params:
            borrow_available_flags = np.asarray(carry_model_params.get("borrow_available_flags"), dtype=bool)

    result = backtest_vectorized(
        prices=prices,
        signals=portfolio_result.target_weights,
        slippage_model=slippage,
        borrow_cost_model=borrow,
        initial_equity=float(starting_capital),
        timeframe=timeframe,
        dates=arrays.date_index,
        symbols=symbol_order,
        carry_asset_classes=[arrays.metadata.asset_class_by_symbol[symbol] for symbol in symbol_order],
        carry_expiry_by_asset=[arrays.metadata.expiry_by_symbol[symbol] for symbol in symbol_order],
        carry_multipliers=[arrays.metadata.multiplier_by_symbol[symbol] for symbol in symbol_order],
        carry_borrow_availability_tiers=[
            arrays.metadata.borrow_availability_tier_by_symbol[symbol] for symbol in symbol_order
        ],
        carry_financing_benchmarks=[arrays.metadata.financing_benchmark_by_symbol[symbol] for symbol in symbol_order],
        borrow_rate_series=borrow_rate_series,
        borrow_available_flags=borrow_available_flags,
    )

    timestamps = arrays.date_index

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

    parameter_payload = {
        "tickers": list(tickers),
        "lookback_days": lookback_days,
        "skip_days": skip_days,
        "costs_bps": costs_bps,
        "execution_model": execution_model,
        "execution_model_params": execution_model_params or {},
        "carry_model": carry_model,
        "carry_model_params": carry_model_params or {},
        "entry_signal": entry_signal,
        "entry_signal_params": entry_signal_params or {},
        "exit_signal": exit_signal,
        "exit_signal_params": exit_signal_params or {},
        "start_date": start_date.isoformat(),
        "end_date": end_date.isoformat(),
        "signal_rebalance_interval": signal_rebalance_interval,
        "starting_capital": starting_capital,
        "bet_sizing_mode": bet_sizing_mode,
        "custom_bet_pct": custom_bet_pct,
        "resolved_bet_pct": bet_fraction * 100.0,
        "strategy": strategy,
        "xsmom_top_quantile": xsmom_top_quantile,
        "xsmom_bottom_quantile": xsmom_bottom_quantile,
        "xsmom_long_only": xsmom_long_only,
        "xsmom_vol_lookback_days": xsmom_vol_lookback_days,
        "timeframe": timeframe,
        "portfolio_method": portfolio_method,
        "portfolio_vol_lookback_bars": portfolio_vol_lookback_bars,
        "portfolio_target_volatility": portfolio_target_volatility,
        "portfolio_max_symbol_weight": portfolio_max_symbol_weight,
        "portfolio_max_sector_weight": portfolio_max_sector_weight,
        "portfolio_max_gross_exposure": portfolio_max_gross_exposure,
        "portfolio_min_net_exposure": portfolio_min_net_exposure,
        "portfolio_max_net_exposure": portfolio_max_net_exposure,
        "portfolio_sector_map": dict(portfolio_sector_map or {}),
        "cache_root": str(cache_root),
    }

    fill_rows = list(result.fills)
    robustness_report = build_backtest_robustness_report(
        returns=returns,
        metrics=metrics,
        turnover_stats=turnover_stats,
        cost_totals=cost_totals,
        fills=fill_rows,
    )

    run_dir = _persist_backtest_outputs(
        timestamps=timestamps,
        symbol_order=symbol_order,
        equity=equity,
        returns=returns,
        trades=trades,
        risk_diagnostics=portfolio_result.diagnostics,
        metrics=metrics,
        dataset_contracts=dataset_contracts,
        parameters=parameter_payload,
        data_snapshot=_build_data_snapshot_identifiers(arrays=arrays, cache_root=cache_root, timeframe=timeframe),
        random_seed=None,
        robustness_report=robustness_report,
    )

    trade_log_rows = _build_trade_log_rows(
        timestamps=timestamps,
        symbol_order=symbol_order,
        prices=prices,
        trades=trades,
        costs_bps=costs_bps,
        starting_capital=starting_capital,
    )
    trade_log_summary = _format_trade_log_summary(trade_log_rows)

    report_text = format_backtest_report(
        title=f"{strategy.upper()} Backtest" if strategy != "momentum" else "Time-Series Momentum Backtest",
        params={
            "tickers": ", ".join(tickers),
            **{k: v for k, v in parameter_payload.items() if k not in {"tickers", "cache_root", "portfolio_sector_map"}},
        },
        metrics=metrics,
        drawdown_rows=drawdown_rows,
        turnover_stats=turnover_stats,
        cost_totals=cost_totals,
        robustness_report=robustness_report,
    )
    (run_dir / "trade_log.csv").write_text(_trade_log_csv(trade_log_rows))
    (run_dir / "trade_log.json").write_text(json.dumps(trade_log_rows, indent=2))
    (run_dir / "fills.csv").write_text(_fills_csv(fill_rows))
    (run_dir / "fills.json").write_text(json.dumps(fill_rows, indent=2))

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
        parameters={
            "tickers": list(tickers),
            "start_date": start_date.isoformat(),
            "end_date": end_date.isoformat(),
            "cache_root": str(cache_root),
            "entry_grid": entry_grid,
            "exit_grid": exit_grid,
            "core_grid": core_grid,
            "max_workers": n_workers,
            "fail_fast": fail_fast,
            "continue_on_error": continue_on_error,
            "top_n": top_n,
        },
        random_seed=seed,
    )
    return (
        f"Sweep complete: {len(ranked_rows)} successful combos, "
        f"{len(invalid_rows)} skipped, {len(errors)} failed. "
        f"Saved outputs to: {run_dir}"
    )


def run_walk_forward_backtest(
    *,
    tickers: list[str],
    start_date: date,
    end_date: date,
    cache_root: Path,
    entry_grid: dict[str, list[dict[str, Any]]],
    exit_grid: dict[str, list[dict[str, Any]]],
    core_grid: dict[str, list[Any]],
    train_bars: int | None = None,
    validation_bars: int | None = None,
    test_bars: int | None = None,
    step_bars: int | None = None,
    train_fraction: float | None = None,
    validation_fraction: float | None = None,
    test_fraction: float | None = None,
    step_fraction: float | None = None,
    score_metric: str = "sharpe",
    purge_window_bars: int = 0,
    embargo_window_bars: int = 0,
    label_horizon_bars: int = 1,
    nested_optimization: bool = False,
    inner_train_fraction: float = 0.7,
) -> str:
    combos, invalid_rows = generate_sweep_combinations(
        entry_grid=entry_grid,
        exit_grid=exit_grid,
        core_grid=core_grid,
    )
    if not combos:
        raise ValueError("No valid combinations generated for walk-forward")

    max_lookback = 1
    for combo in combos:
        entry_cfg = parse_entry_signal_config(
            str(combo["entry_signal"]),
            dict(combo.get("entry_signal_params", {})),
            default_lookback_days=int(combo["lookback_days"]),
            default_skip_days=int(combo["skip_days"]),
        )
        exit_cfg = parse_exit_signal_config(
            str(combo["exit_signal"]),
            dict(combo.get("exit_signal_params", {})),
            default_lookback_days=int(combo["lookback_days"]),
            default_skip_days=int(combo["skip_days"]),
        )
        max_lookback = max(max_lookback, required_lookback_window(entry_cfg, exit_cfg))

    arrays = load_backtest_engine_arrays(
        tickers=tickers,
        start=datetime.combine(start_date, datetime.min.time()).isoformat(),
        end=datetime.combine(end_date, datetime.max.time()).isoformat(),
        cache_root=cache_root,
        timeframe="1m",
        lookback_window=max_lookback,
    )
    prices = _fill_missing_prices(arrays.close_prices)

    total_bars = int(prices.shape[0])
    if train_fraction is not None or validation_fraction is not None or test_fraction is not None:
        if train_fraction is None or validation_fraction is None or test_fraction is None:
            raise ValueError("train/validation/test fractions must all be provided together")
        frac_sum = float(train_fraction) + float(validation_fraction) + float(test_fraction)
        if frac_sum <= 0:
            raise ValueError("walk-forward fractions must sum to a positive value")
        required_gap_bars = max(int(purge_window_bars), int(label_horizon_bars)) + max(
            int(embargo_window_bars),
            int(label_horizon_bars),
        )
        effective_total_bars = max(1, total_bars - required_gap_bars)
        train_bars = max(1, int(round(effective_total_bars * float(train_fraction) / frac_sum)))
        validation_bars = max(1, int(round(effective_total_bars * float(validation_fraction) / frac_sum)))
        remaining_for_test = effective_total_bars - train_bars - validation_bars
        if remaining_for_test < 1:
            overflow = 1 - remaining_for_test
            reducible_train = max(0, train_bars - 1)
            reduce_train = min(reducible_train, overflow)
            train_bars -= reduce_train
            overflow -= reduce_train
            if overflow > 0:
                reducible_validation = max(0, validation_bars - 1)
                reduce_validation = min(reducible_validation, overflow)
                validation_bars -= reduce_validation
                overflow -= reduce_validation
            remaining_for_test = effective_total_bars - train_bars - validation_bars
        test_bars = max(1, remaining_for_test)
        step_bars = max(
            1,
            int(round(effective_total_bars * float(step_fraction if step_fraction is not None else test_fraction))),
        )
    elif train_bars is None or validation_bars is None or test_bars is None:
        raise ValueError("Either explicit bar windows or train/validation/test fractions are required")

    folds = build_walk_forward_folds(
        total_bars=total_bars,
        train_bars=int(train_bars),
        validation_bars=int(validation_bars),
        test_bars=int(test_bars),
        step_bars=None if step_bars is None else int(step_bars),
        purge_window_bars=int(purge_window_bars),
        embargo_window_bars=int(embargo_window_bars),
        label_horizon_bars=int(label_horizon_bars),
    )
    if not folds:
        raise ValueError("No folds generated; increase date range or reduce window sizes")

    def evaluate_segment(candidate: dict[str, Any], start_idx: int, end_idx: int) -> dict[str, Any]:
        if end_idx <= start_idx:
            return {"metrics": {}, "equity": []}

        entry_cfg = parse_entry_signal_config(
            str(candidate["entry_signal"]),
            dict(candidate.get("entry_signal_params", {})),
            default_lookback_days=int(candidate["lookback_days"]),
            default_skip_days=int(candidate["skip_days"]),
        )
        exit_cfg = parse_exit_signal_config(
            str(candidate["exit_signal"]),
            dict(candidate.get("exit_signal_params", {})),
            default_lookback_days=int(candidate["lookback_days"]),
            default_skip_days=int(candidate["skip_days"]),
        )
        segment_prices = prices[start_idx:end_idx]
        segment_missing = arrays.missing_mask[start_idx:end_idx]
        signals = build_targets(
            close_prices=segment_prices,
            missing_mask=segment_missing,
            entry_config=entry_cfg,
            exit_config=exit_cfg,
        )
        result = backtest_vectorized(
            prices=segment_prices,
            signals=signals,
            slippage_model=BpsSlippage(float(candidate["costs_bps"])),
            initial_equity=1.0,
            timeframe="1m",
        )
        metrics = {key: float(value) for key, value in dict(result.metrics).items()}
        equity = _to_numpy_1d(result.equity_curve)
        timestamps = arrays.date_index[start_idx:end_idx]
        equity_rows = [
            {"timestamp": _timestamp_to_iso8601(timestamps[idx]), "equity": float(equity[idx])}
            for idx in range(min(len(timestamps), equity.size))
        ]
        return {"metrics": metrics, "equity": equity_rows}

    nested_inner_folds: dict[int, list[Any]] | None = None
    if nested_optimization:
        nested_inner_folds = {}
        for outer_fold in folds:
            outer_train_total = max(0, outer_fold.validation_end - outer_fold.train_start)
            if outer_train_total < 3:
                continue
            inner_train_bars = max(1, int(round(outer_train_total * float(inner_train_fraction))))
            inner_validation_bars = max(1, outer_train_total - inner_train_bars)
            if inner_train_bars + inner_validation_bars > outer_train_total:
                inner_validation_bars = max(1, outer_train_total - inner_train_bars)
            if inner_train_bars <= 0 or inner_validation_bars <= 0:
                continue
            inner_total = outer_fold.validation_end
            inner_folds = build_walk_forward_folds(
                total_bars=inner_total,
                train_bars=inner_train_bars,
                validation_bars=inner_validation_bars,
                test_bars=1,
                step_bars=max(1, inner_validation_bars),
                purge_window_bars=int(purge_window_bars),
                embargo_window_bars=0,
                label_horizon_bars=int(label_horizon_bars),
            )
            filtered_inner_folds = [
                fold
                for fold in inner_folds
                if fold.train_start >= outer_fold.train_start and fold.validation_end <= outer_fold.validation_end
            ]
            if filtered_inner_folds:
                nested_inner_folds[outer_fold.fold_id] = filtered_inner_folds

    wf_result = run_walk_forward_optimization(
        folds=folds,
        parameter_candidates=combos,
        evaluate_segment=evaluate_segment,
        score_metric=score_metric,
        nested_inner_folds=nested_inner_folds,
    )

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    run_dir = BACKTEST_OUTPUT_DIR / f"tsmom_walk_forward_{timestamp}"
    persist_walk_forward_outputs(run_dir=run_dir, result=wf_result)
    (run_dir / "skipped_invalid_combos.json").write_text(json.dumps(invalid_rows, indent=2))
    report_text = _format_walk_forward_report(
        folds=wf_result.folds,
        aggregate_metrics=wf_result.aggregate_metrics,
        stability=wf_result.stability,
        score_metric=score_metric,
        candidate_count=len(combos),
        skipped_invalid_count=len(invalid_rows),
    )
    (run_dir / "report.txt").write_text(report_text)

    return (
        f"Walk-forward complete: {len(wf_result.folds)} folds, "
        f"{len(combos)} candidates, {len(invalid_rows)} skipped invalid combos. "
        f"Saved outputs to: {run_dir}\n"
        f"Report: {run_dir / 'report.txt'}"
    )


def run_strategy_optimization(
    *,
    tickers: list[str],
    start_date: date,
    end_date: date,
    cache_root: Path,
    entry_grid: dict[str, list[dict[str, Any]]],
    exit_grid: dict[str, list[dict[str, Any]]],
    core_grid: dict[str, list[Any]],
    seed: int = 42,
    n_trials: int = 30,
    sampler_name: str = "tpe",
    objectives: list[dict[str, str]] | None = None,
    max_turnover: float | None = None,
    max_drawdown_floor: float | None = None,
    min_trades: float | None = None,
    partial_period_fractions: list[float] | None = None,
) -> str:
    combos, invalid_rows = generate_sweep_combinations(
        entry_grid=entry_grid,
        exit_grid=exit_grid,
        core_grid=core_grid,
    )
    if not combos:
        raise ValueError("No valid combinations generated for optimization")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    run_dir = BACKTEST_OUTPUT_DIR / f"tsmom_optimize_{timestamp}"
    run_dir.mkdir(parents=True, exist_ok=True)

    objective_defs = objectives or [
        {"name": "sharpe", "sense": "maximize"},
        {"name": "turnover_total", "sense": "minimize"},
        {"name": "max_drawdown", "sense": "maximize"},
    ]
    objective_specs = [Objective(name=str(item["name"]), sense=str(item.get("sense", "maximize"))) for item in objective_defs]

    constraints: list[Constraint] = []
    if max_turnover is not None:
        constraints.append(Constraint(metric="turnover_total", max_value=float(max_turnover)))
    if max_drawdown_floor is not None:
        constraints.append(Constraint(metric="max_drawdown", min_value=float(max_drawdown_floor)))
    if min_trades is not None:
        constraints.append(Constraint(metric="trade_count", min_value=float(min_trades)))

    sampler = TPESampler() if sampler_name.lower() == "tpe" else RandomSampler()

    def _evaluate(params: dict[str, Any], period_fraction: float) -> dict[str, float]:
        combo = combos[int(params["combo_index"])]
        row = _execute_sweep_combo(
            {
                "combo_index": int(params["combo_index"]),
                "seed": seed,
                "tickers": tickers,
                "start_date": start_date.isoformat(),
                "end_date": end_date.isoformat(),
                "cache_root": str(cache_root),
                "partial_period_fraction": float(period_fraction),
                **combo,
            }
        )
        return {key: float(value) for key, value in row.items() if isinstance(value, (int, float))}

    result = optimize(
        space={"combo_index": list(range(len(combos)))},
        evaluate=_evaluate,
        objectives=objective_specs,
        constraints=constraints,
        sampler=sampler,
        n_trials=int(n_trials),
        seed=int(seed),
        partial_period_fractions=partial_period_fractions,
        output_dir=run_dir,
    )

    (run_dir / "invalid_combinations.json").write_text(json.dumps(invalid_rows, indent=2), encoding="utf-8")
    (run_dir / "optimizer_manifest.json").write_text(
        json.dumps(
            {
                "tickers": tickers,
                "start_date": start_date.isoformat(),
                "end_date": end_date.isoformat(),
                "entry_grid": entry_grid,
                "exit_grid": exit_grid,
                "core_grid": core_grid,
                "seed": seed,
                "n_trials": n_trials,
                "sampler": sampler_name,
                "objectives": objective_defs,
                "constraints": [constraint.__dict__ for constraint in constraints],
                "partial_period_fractions": partial_period_fractions,
            },
            indent=2,
            sort_keys=True,
        ),
        encoding="utf-8",
    )

    return (
        f"Optimization complete: {result['trial_count']} trials, {result['feasible_count']} feasible, "
        f"{result['pareto_count']} Pareto-optimal. Saved outputs to: {run_dir}"
    )


def _format_walk_forward_report(
    *,
    folds: list[dict[str, Any]],
    aggregate_metrics: dict[str, float],
    stability: dict[str, Any],
    score_metric: str,
    candidate_count: int,
    skipped_invalid_count: int,
) -> str:
    lines = [
        "Walk-Forward Backtest Report",
        "============================",
        "",
        f"Folds: {len(folds)}",
        f"Candidates evaluated per fold: {candidate_count}",
        f"Skipped invalid combinations: {skipped_invalid_count}",
        f"Selection score metric: {score_metric}",
        "",
        "Aggregate OOS Metrics",
        "---------------------",
    ]
    if aggregate_metrics:
        for key in sorted(aggregate_metrics):
            lines.append(f"- {key}: {aggregate_metrics[key]:.6f}")
    else:
        lines.append("- none")

    lines.extend(["", "Stability", "---------"])
    unique_count = int(stability.get("unique_selected_params", 0))
    lines.append(f"- unique_selected_params: {unique_count}")
    lines.append(f"- validation_score_mean: {float(stability.get('validation_score_mean', 0.0)):.6f}")
    lines.append(f"- validation_score_std: {float(stability.get('validation_score_std', 0.0)):.6f}")

    selection_counts = stability.get("selection_counts", {})
    if isinstance(selection_counts, dict) and selection_counts:
        lines.append("- selection_counts:")
        for key, value in sorted(
            selection_counts.items(),
            key=lambda item: (-int(item[1]), str(item[0])),
        ):
            lines.append(f"  - {key}: {int(value)}")

    lines.extend(["", "Fold Picks", "---------"])
    for fold in folds:
        lines.append(
            f"- fold={int(fold.get('fold_id', -1))} validation_score={float(fold.get('validation_score', 0.0)):.6f} "
            f"selected_params={json.dumps(fold.get('selected_params', {}), sort_keys=True)}"
        )

    return "\n".join(lines)


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
    partial_fraction = float(payload.get("partial_period_fraction", 1.0))
    if partial_fraction < 1.0:
        keep_bars = max(2, int(round(prices.shape[0] * partial_fraction)))
        prices = prices[:keep_bars]
        missing_mask = arrays.missing_mask[:keep_bars]
    else:
        missing_mask = arrays.missing_mask
    signals = build_targets(
        close_prices=prices,
        missing_mask=missing_mask,
        entry_config=entry_cfg,
        exit_config=exit_cfg,
    )
    result = backtest_vectorized(
        prices=prices,
        signals=signals,
        slippage_model=BpsSlippage(costs_bps),
        initial_equity=1.0,
        timeframe="1m",
    )

    turnover = _to_numpy_1d(result.turnover)
    trades = _to_numpy_2d(result.trades)
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
        "cagr": float(metrics.get("cagr", 0.0)),
        "max_drawdown": float(metrics.get("max_drawdown", 0.0)),
        "calmar": float(metrics.get("calmar", 0.0)),
        "volatility": float(metrics.get("volatility", 0.0)),
        "sortino": float(metrics.get("sortino", 0.0)),
        "downside_deviation": float(metrics.get("downside_deviation", 0.0)),
        "hit_rate": float(metrics.get("hit_rate", 0.0)),
        "profit_factor": float(metrics.get("profit_factor", 0.0)),
        "exposure_time": float(metrics.get("exposure_time", 0.0)),
        "turnover_adjusted_return": float(metrics.get("turnover_adjusted_return", 0.0)),
        "rolling_sharpe_mean": float(metrics.get("rolling_sharpe_mean", 0.0)),
        "rolling_drawdown_worst": float(metrics.get("rolling_drawdown_worst", 0.0)),
        "turnover_total": float(np.sum(turnover)) if turnover.size else 0.0,
        "trade_count": float(np.sum(np.abs(trades) > 1e-12)) if trades.size else 0.0,
        "cost_total": cost_totals.get("total", 0.0),
    }


def _persist_sweep_outputs(
    *,
    ranked_rows: list[dict[str, Any]],
    invalid_rows: list[dict[str, Any]],
    errors: list[dict[str, Any]],
    top_n: int,
    parameters: dict[str, Any],
    random_seed: int,
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
    robustness_report = build_sweep_robustness_report(ranked_rows=ranked_rows)
    _write_robustness_report(run_dir=run_dir, robustness_report=robustness_report)

    report_lines = ["Top sweep combinations", "======================", ""]
    for idx, row in enumerate(top_rows, start=1):
        report_lines.append(
            f"#{idx}: sharpe={row['sharpe']:.6f} total_return={row['total_return']:.6f} "
            f"entry={row['entry_signal']} exit={row['exit_signal']} "
            f"core=(lookback_days={row['lookback_days']}, skip_days={row['skip_days']}, costs_bps={row['costs_bps']})"
        )
    pbo_style = robustness_report.get("pbo_style", {})
    report_lines.extend([
        "",
        "Robustness Diagnostics",
        "----------------------",
        f"deflated_sharpe_ratio={float(robustness_report.get('deflated_sharpe_ratio', 0.0)):.6f}",
        f"pbo_probability={float(pbo_style.get('probability_of_overfitting', 0.0)):.6f}",
        f"pbo_median_logit={float(pbo_style.get('median_logit', 0.0)):.6f}",
    ])
    (run_dir / "top_n_report.txt").write_text("\n".join(report_lines))

    manifest = {
        "run_type": "parameter_sweep",
        "created_at": datetime.now().isoformat(),
        "code_version": _resolve_git_commit(),
        "metric_schema_version": CANONICAL_METRIC_SCHEMA_VERSION,
        "parameters": parameters,
        "data_snapshot_identifiers": _build_sweep_snapshot_identifiers(parameters),
        "random_seed": int(random_seed),
        "environment": _collect_environment_metadata(),
        "result_summary": {
            "successful_combos": len(ranked_rows),
            "invalid_combos": len(invalid_rows),
            "failed_combos": len(errors),
            "best_sharpe": float(ranked_rows[0]["sharpe"]) if ranked_rows else 0.0,
        },
        "reproducibility_fingerprint": _stable_fingerprint(
            {
                "metric_schema_version": CANONICAL_METRIC_SCHEMA_VERSION,
                "parameters": parameters,
                "random_seed": int(random_seed),
                "top_row": ranked_rows[0] if ranked_rows else {},
            }
        ),
    }
    (run_dir / "manifest.json").write_text(json.dumps(manifest, indent=2))

    _append_experiment_index(
        {
            "timestamp": manifest["created_at"],
            "run_type": "parameter_sweep",
            "run_dir": str(run_dir),
            "code_version": manifest["code_version"],
            "metric_schema_version": CANONICAL_METRIC_SCHEMA_VERSION,
            "random_seed": int(random_seed),
            "primary_metric": "sharpe",
            "primary_metric_value": float(ranked_rows[0]["sharpe"]) if ranked_rows else 0.0,
            "manifest_path": str(run_dir / "manifest.json"),
            "reproducibility_fingerprint": manifest["reproducibility_fingerprint"],
            "parameters": parameters,
            "data_snapshot_identifiers": manifest["data_snapshot_identifiers"],
        }
    )

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
    starting_capital: float = 100_000.0,
    bet_sizing_mode: str = "half_kelly",
    custom_bet_pct: float = 10.0,
    timeframe: str = "1m",
    portfolio_method: str = "equal_weight",
    portfolio_vol_lookback_bars: int = 20,
    portfolio_target_volatility: float = 0.10,
    portfolio_max_symbol_weight: float = 0.25,
    portfolio_max_sector_weight: float = 0.60,
    portfolio_max_gross_exposure: float = 1.0,
    portfolio_min_net_exposure: float = -1.0,
    portfolio_max_net_exposure: float = 1.0,
    portfolio_sector_map: dict[str, str] | None = None,
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
                starting_capital=starting_capital,
                bet_sizing_mode=bet_sizing_mode,
                custom_bet_pct=custom_bet_pct,
                entry_signal=entry_signal,
                entry_signal_params={
                    "min_abs_return": 0.01,
                    "long_only": True,
                } if entry_signal == "ts_momentum" else {},
                exit_signal=exit_signal,
                exit_signal_params={
                    "min_abs_return": 0.01,
                } if exit_signal == "momentum_flip" else {},
                timeframe=timeframe,
                portfolio_method=portfolio_method,
                portfolio_vol_lookback_bars=portfolio_vol_lookback_bars,
                portfolio_target_volatility=portfolio_target_volatility,
                portfolio_max_symbol_weight=portfolio_max_symbol_weight,
                portfolio_max_sector_weight=portfolio_max_sector_weight,
                portfolio_max_gross_exposure=portfolio_max_gross_exposure,
                portfolio_min_net_exposure=portfolio_min_net_exposure,
                portfolio_max_net_exposure=portfolio_max_net_exposure,
                portfolio_sector_map=portfolio_sector_map,
            )
            run_dir = _extract_saved_output_dir(report)
            metrics = _load_metrics_from_run_dir(run_dir)
            rows.append(
                {
                    "entry_signal": entry_signal,
                    "exit_signal": exit_signal,
                    "total_return": float(metrics.get("total_return", 0.0)),
                    "sharpe": float(metrics.get("sharpe", 0.0)),
                    "cagr": float(metrics.get("cagr", 0.0)),
                    "max_drawdown": float(metrics.get("max_drawdown", 0.0)),
                    "calmar": float(metrics.get("calmar", 0.0)),
                    "volatility": float(metrics.get("volatility", 0.0)),
                    "sortino": float(metrics.get("sortino", 0.0)),
                    "hit_rate": float(metrics.get("hit_rate", 0.0)),
                    "profit_factor": float(metrics.get("profit_factor", 0.0)),
                    "turnover_adjusted_return": float(metrics.get("turnover_adjusted_return", 0.0)),
                    "rolling_sharpe_mean": float(metrics.get("rolling_sharpe_mean", 0.0)),
                    "rolling_drawdown_worst": float(metrics.get("rolling_drawdown_worst", 0.0)),
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
        f"Starting capital: {starting_capital:.2f}",
        f"Bet sizing mode: {bet_sizing_mode}",
        "Ranked combinations (by sharpe):",
    ]
    for idx, row in enumerate(ranked_rows, start=1):
        summary_lines.append(
            f"#{idx} entry={row['entry_signal']} exit={row['exit_signal']} "
            f"sharpe={row['sharpe']:.6f} total_return={row['total_return']:.6f} "
            f"cagr={row['cagr']:.6f} calmar={row['calmar']:.6f} "
            f"max_drawdown={row['max_drawdown']:.6f} volatility={row['volatility']:.6f} "
            f"sortino={row['sortino']:.6f} hit_rate={row['hit_rate']:.6f} "
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
        "cagr",
        "max_drawdown",
        "calmar",
        "volatility",
        "sortino",
        "hit_rate",
        "profit_factor",
        "turnover_adjusted_return",
        "rolling_sharpe_mean",
        "rolling_drawdown_worst",
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



def _timeframe_step_from_1m(timeframe: str) -> int:
    key = str(timeframe).strip().lower()
    mapping = {
        "1m": 1,
        "5m": 5,
        "15m": 15,
        "30m": 30,
        "1h": 60,
        "1d": 390,
    }
    if key not in mapping:
        raise ValueError(f"Unsupported timeframe: {timeframe}")
    return mapping[key]


def _resample_engine_bundle_from_1m(arrays: EngineArrayBundle, *, timeframe: str) -> EngineArrayBundle:
    step = _timeframe_step_from_1m(timeframe)
    if step <= 1 or arrays.date_index.size == 0:
        return arrays

    start_idx = step - 1
    if start_idx >= arrays.date_index.size:
        start_idx = arrays.date_index.size - 1
    selector = np.arange(start_idx, arrays.date_index.size, step, dtype=int)
    if selector.size == 0:
        selector = np.array([arrays.date_index.size - 1], dtype=int)

    date_index = arrays.date_index[selector]
    open_prices = arrays.open_prices[selector]
    close_prices = arrays.close_prices[selector]
    missing_mask = arrays.missing_mask[selector]

    symbol_order = sorted(arrays.metadata.symbol_to_column.items(), key=lambda item: item[1])
    missingness_by_symbol = {
        symbol: float(np.mean(missing_mask[:, col])) if missing_mask.size else 1.0
        for symbol, col in symbol_order
    }
    metadata = EngineArrayMetadata(
        symbol_to_column=dict(arrays.metadata.symbol_to_column),
        date_index=date_index,
        missingness_ratio=float(np.mean(missing_mask)) if missing_mask.size else 1.0,
        missingness_by_symbol=missingness_by_symbol,
        coverage_by_symbol=dict(arrays.metadata.coverage_by_symbol),
        tradable_ratio_by_symbol={
            symbol: (1.0 - missingness_by_symbol[symbol])
            for symbol, _ in symbol_order
        },
        excluded_symbols=dict(arrays.metadata.excluded_symbols),
        audit_summary_by_symbol=dict(arrays.metadata.audit_summary_by_symbol),
        asset_class_by_symbol=dict(arrays.metadata.asset_class_by_symbol),
        expiry_by_symbol=dict(arrays.metadata.expiry_by_symbol),
        multiplier_by_symbol=dict(arrays.metadata.multiplier_by_symbol),
        borrow_availability_tier_by_symbol=dict(arrays.metadata.borrow_availability_tier_by_symbol),
        financing_benchmark_by_symbol=dict(arrays.metadata.financing_benchmark_by_symbol),
        pit_membership_violations_by_symbol=dict(arrays.metadata.pit_membership_violations_by_symbol),
        adjustment_violations_by_symbol=dict(arrays.metadata.adjustment_violations_by_symbol),
        delisted_symbols=list(arrays.metadata.delisted_symbols),
        survivorship_bias_flags_by_symbol=dict(arrays.metadata.survivorship_bias_flags_by_symbol),
        leakage_flags_by_symbol=dict(arrays.metadata.leakage_flags_by_symbol),
    )
    return EngineArrayBundle(
        date_index=date_index,
        open_prices=open_prices,
        close_prices=close_prices,
        missing_mask=missing_mask,
        metadata=metadata,
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


def _resolve_bet_fraction(*, prices: np.ndarray, mode: str, custom_bet_pct: float) -> float:
    if mode == "custom":
        return float(np.clip(custom_bet_pct / 100.0, 0.0, 1.0))

    rets = np.diff(prices, axis=0) / np.where(prices[:-1] == 0.0, np.nan, prices[:-1])
    finite = rets[np.isfinite(rets)]
    if finite.size == 0:
        base = 0.1
    else:
        mean_ret = float(np.mean(finite))
        var_ret = float(np.var(finite))
        if var_ret <= 1e-12:
            base = 0.1
        else:
            base = max(0.0, min(1.0, mean_ret / var_ret))
    if mode == "kelly":
        return float(np.clip(base, 0.0, 1.0))
    return float(np.clip(base * 0.5, 0.0, 1.0))


def _apply_discrete_bet_sizing(
    *,
    signals: np.ndarray,
    prices: np.ndarray,
    starting_capital: float,
    bet_fraction: float,
) -> np.ndarray:
    if starting_capital <= 0:
        raise ValueError("starting_capital must be > 0")
    if bet_fraction <= 0:
        return np.zeros_like(signals, dtype=float)
    sized = np.zeros_like(signals, dtype=float)
    for idx in range(signals.shape[0]):
        row = signals[idx]
        active = np.flatnonzero(row != 0.0)
        if active.size == 0:
            continue
        per_asset_budget = (starting_capital * bet_fraction) / float(active.size)
        for col_idx in active:
            px = float(prices[idx, col_idx])
            if px <= 0.0 or not np.isfinite(px):
                continue
            shares = int(np.floor(per_asset_budget / px))
            if shares <= 0:
                continue
            direction = 1.0 if row[col_idx] > 0 else -1.0
            sized[idx, col_idx] = direction * (shares * px / starting_capital)
    return sized


def _save_portfolio_value_chart(run_dir: Path) -> Path | None:
    equity_path = run_dir / "equity.csv"
    if not equity_path.exists():
        return None
    rows: list[tuple[str, float]] = []
    with equity_path.open() as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            ts = str(row.get("timestamp", ""))
            day = ts.split("T")[0]
            eq = float(row.get("equity", 0.0))
            rows.append((day, eq))
    if not rows:
        return None
    daily: dict[str, float] = {}
    for day, eq in rows:
        daily[day] = eq
    days = sorted(daily)
    values = [daily[d] for d in days]
    try:
        import matplotlib.pyplot as plt
    except Exception:
        return None
    plt.figure(figsize=(10, 4))
    plt.plot(days, values, linewidth=1.5)
    plt.title("Portfolio Value Over Time")
    plt.xlabel("day")
    plt.ylabel("portfolio value")
    plt.xticks(rotation=45, ha="right")
    plt.tight_layout()
    out = run_dir / "portfolio_value_over_time.png"
    plt.savefig(out)
    plt.close()
    return out


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
    starting_capital: float,
) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for col_idx, symbol in enumerate(symbol_order):
        side = 0
        entry_price = 0.0
        position_shares = 0
        running_pnl = 0.0
        for row_idx, ts in enumerate(timestamps):
            trade = float(trades[row_idx, col_idx])
            if trade == 0.0:
                continue
            price = float(prices[row_idx, col_idx])
            shares_delta = int(np.floor(abs(trade) * starting_capital / max(price, 1e-12)))
            if shares_delta <= 0:
                continue
            event = "adjust"
            trade_pnl = 0.0
            trade_cost = (shares_delta * price) * (costs_bps / 10_000.0)
            if side == 0 and abs(trade) > 0:
                side = 1 if trade > 0 else -1
                position_shares = shares_delta
                entry_price = price
                event = "entry"
            elif side != 0 and side * trade < 0:
                trade_pnl = ((price - entry_price) * side * position_shares) if entry_price > 0 else 0.0
                running_pnl += trade_pnl - trade_cost
                event = "exit" if abs(trade) == abs(side) else "flip"
                next_side = side + int(np.sign(trade))
                side = 0 if next_side == 0 else (1 if next_side > 0 else -1)
                if side != 0:
                    position_shares = shares_delta
                    entry_price = price
                else:
                    position_shares = 0
                    entry_price = 0.0
            rows.append(
                {
                    "timestamp": datetime.utcfromtimestamp(int(ts) / 1000.0).isoformat(),
                    "symbol": symbol,
                    "event": event,
                    "trade": trade,
                    "shares_delta": float(shares_delta),
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
    lines.append("timestamp | symbol | event | trade | shares | price | trade_pnl | trade_cost | running_pnl")
    for row in rows[:max_rows]:
        lines.append(
            f"{row['timestamp']} | {row['symbol']} | {row['event']} | {float(row['trade']):.2f} | {float(row['shares_delta']):.0f} | "
            f"{float(row['price']):.4f} | {float(row['trade_pnl']):.6f} | {float(row['trade_cost']):.6f} | {float(row['running_pnl']):.6f}"
        )
    if len(rows) > max_rows:
        lines.append(f"... ({len(rows) - max_rows} more trade events)")
    return "\n".join(lines)


def _trade_log_csv(rows: list[dict[str, object]]) -> str:
    header = "timestamp,symbol,event,trade,shares_delta,price,trade_pnl,trade_cost,running_pnl"
    body = [
        f"{row['timestamp']},{row['symbol']},{row['event']},{row['trade']},{row['shares_delta']},{row['price']},{row['trade_pnl']},{row['trade_cost']},{row['running_pnl']}"
        for row in rows
    ]
    return "\n".join([header, *body]) + "\n"



def _fills_csv(rows: list[dict[str, object]]) -> str:
    header = "bar_index,asset_index,requested_size,filled_size,residual_size,participation_rate,available_volume,order_type,latency_bars,latency_ms,queue_rank_proxy"
    body = [
        f"{int(row.get('bar_index', 0))},{int(row.get('asset_index', 0))},{float(row.get('requested_size', 0.0))},{float(row.get('filled_size', 0.0))},{float(row.get('residual_size', 0.0))},{float(row.get('participation_rate', 0.0))},{float(row.get('available_volume', 0.0))},{str(row.get('order_type', 'market'))},{int(row.get('latency_bars', 0))},{int(row.get('latency_ms', 0))},{float(row.get('queue_rank_proxy', 0.5))}"
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
    risk_diagnostics: dict[str, np.ndarray],
    metrics: dict[str, float],
    dataset_contracts: object | None = None,
    parameters: dict[str, Any] | None = None,
    data_snapshot: dict[str, Any] | None = None,
    random_seed: int | None = None,
    robustness_report: dict[str, Any] | None = None,
) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    run_dir = BACKTEST_OUTPUT_DIR / f"tsmom_backtest_{timestamp}"
    run_dir.mkdir(parents=True, exist_ok=True)

    time_strings = [_timestamp_to_iso8601(ts) for ts in timestamps]

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


    _write_risk_diagnostics_csv_json(
        run_dir=run_dir,
        timestamps=time_strings,
        symbol_order=symbol_order,
        diagnostics=risk_diagnostics,
    )

    metrics_rows = [{"metric": key, "value": float(value)} for key, value in metrics.items()]
    metrics_csv = run_dir / "metrics.csv"
    with metrics_csv.open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["metric", "value"])
        writer.writeheader()
        writer.writerows(metrics_rows)
    (run_dir / "metrics.json").write_text(json.dumps(metrics_rows, indent=2))
    (run_dir / "metric_schema_version.txt").write_text(f"{CANONICAL_METRIC_SCHEMA_VERSION}\n")


    if robustness_report is not None:
        _write_robustness_report(run_dir=run_dir, robustness_report=robustness_report)

    if dataset_contracts is not None:
        audit_payload = {
            "coverage_by_symbol": dataset_contracts.coverage_by_symbol,
            "missingness_by_symbol": dataset_contracts.missingness_by_symbol,
            "excluded_symbols": dataset_contracts.excluded_symbols,
            "reasons_by_symbol": dataset_contracts.reasons_by_symbol,
            "survivorship_bias_flags_by_symbol": dataset_contracts.survivorship_bias_flags_by_symbol,
            "leakage_flags_by_symbol": dataset_contracts.leakage_flags_by_symbol,
        }
        (run_dir / "dataset_quality_audit.json").write_text(json.dumps(audit_payload, indent=2))

    manifest = {
        "run_type": "backtest",
        "created_at": datetime.now().isoformat(),
        "code_version": _resolve_git_commit(),
        "metric_schema_version": CANONICAL_METRIC_SCHEMA_VERSION,
        "parameters": parameters or {},
        "data_snapshot_identifiers": data_snapshot or {},
        "random_seed": random_seed,
        "environment": _collect_environment_metadata(),
        "reproducibility_fingerprint": _stable_fingerprint(
            {
                "metric_schema_version": CANONICAL_METRIC_SCHEMA_VERSION,
                "parameters": parameters or {},
                "random_seed": random_seed,
                "data_snapshot_identifiers": data_snapshot or {},
            }
        ),
    }
    (run_dir / "manifest.json").write_text(json.dumps(manifest, indent=2))
    _append_experiment_index(
        {
            "timestamp": manifest["created_at"],
            "run_type": "backtest",
            "run_dir": str(run_dir),
            "code_version": manifest["code_version"],
            "metric_schema_version": CANONICAL_METRIC_SCHEMA_VERSION,
            "random_seed": random_seed,
            "primary_metric": "sharpe",
            "primary_metric_value": float(metrics.get("sharpe", 0.0)),
            "manifest_path": str(run_dir / "manifest.json"),
            "reproducibility_fingerprint": manifest["reproducibility_fingerprint"],
            "parameters": parameters or {},
            "data_snapshot_identifiers": data_snapshot or {},
        }
    )

    return run_dir



def _write_robustness_report(*, run_dir: Path, robustness_report: dict[str, Any]) -> None:
    (run_dir / "robustness_report.json").write_text(json.dumps(robustness_report, indent=2))

    rows: list[dict[str, object]] = []

    def _flatten(prefix: str, value: object) -> None:
        if isinstance(value, dict):
            for key, inner in value.items():
                _flatten(f"{prefix}.{key}" if prefix else str(key), inner)
            return
        if isinstance(value, list):
            for idx, inner in enumerate(value):
                _flatten(f"{prefix}[{idx}]", inner)
            return
        rows.append({"key": prefix, "value": value})

    _flatten("", robustness_report)
    with (run_dir / "robustness_report.csv").open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["key", "value"])
        writer.writeheader()
        writer.writerows(rows)


def _resolve_git_commit() -> str:
    try:
        return subprocess.check_output(["git", "rev-parse", "HEAD"], text=True).strip()
    except Exception:
        return "unknown"


def _collect_environment_metadata() -> dict[str, Any]:
    return {
        "python_version": sys.version,
        "platform": platform.platform(),
        "hostname": socket.gethostname(),
        "pid": os.getpid(),
        "numpy_version": np.__version__,
    }


def _stable_fingerprint(payload: dict[str, Any]) -> str:
    encoded = json.dumps(payload, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def _build_data_snapshot_identifiers(*, arrays: EngineArrayBundle, cache_root: Path, timeframe: str) -> dict[str, Any]:
    symbols = [symbol for symbol, _ in sorted(arrays.metadata.symbol_to_column.items(), key=lambda item: item[1])]
    range_start = _timestamp_to_iso8601(arrays.date_index[0]) if arrays.date_index.size else None
    range_end = _timestamp_to_iso8601(arrays.date_index[-1]) if arrays.date_index.size else None
    payload = {
        "symbols": symbols,
        "timeframe": timeframe,
        "cache_root": str(cache_root),
        "range_start": range_start,
        "range_end": range_end,
        "coverage_by_symbol": arrays.metadata.coverage_by_symbol,
        "missingness_by_symbol": arrays.metadata.missingness_by_symbol,
    }
    payload["dataset_fingerprint"] = _stable_fingerprint(payload)
    return payload


def _build_sweep_snapshot_identifiers(parameters: dict[str, Any]) -> dict[str, Any]:
    payload = {
        "tickers": list(parameters.get("tickers", [])),
        "start_date": parameters.get("start_date"),
        "end_date": parameters.get("end_date"),
        "cache_root": parameters.get("cache_root"),
        "core_grid": parameters.get("core_grid", {}),
    }
    payload["dataset_fingerprint"] = _stable_fingerprint(payload)
    return payload


def _append_experiment_index(entry: dict[str, Any]) -> None:
    BACKTEST_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    jsonl_path = BACKTEST_OUTPUT_DIR / "experiment_index.jsonl"
    with jsonl_path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(entry, sort_keys=True) + "\n")

    csv_path = BACKTEST_OUTPUT_DIR / "experiment_index.csv"
    row = {
        "timestamp": entry.get("timestamp"),
        "run_type": entry.get("run_type"),
        "run_dir": entry.get("run_dir"),
        "code_version": entry.get("code_version"),
        "metric_schema_version": entry.get("metric_schema_version"),
        "random_seed": entry.get("random_seed"),
        "primary_metric": entry.get("primary_metric"),
        "primary_metric_value": entry.get("primary_metric_value"),
        "manifest_path": entry.get("manifest_path"),
        "reproducibility_fingerprint": entry.get("reproducibility_fingerprint"),
        "parameters_json": json.dumps(entry.get("parameters", {}), sort_keys=True),
        "data_snapshot_json": json.dumps(entry.get("data_snapshot_identifiers", {}), sort_keys=True),
    }
    fieldnames = list(row.keys())
    write_header = not csv_path.exists()
    with csv_path.open("a", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        if write_header:
            writer.writeheader()
        writer.writerow(row)




def _timestamp_to_iso8601(value: object) -> str:
    if isinstance(value, datetime):
        return value.isoformat()

    if isinstance(value, np.datetime64):
        nanos = value.astype("datetime64[ns]").astype(np.int64)
        return datetime.utcfromtimestamp(float(nanos) / 1_000_000_000.0).isoformat()

    if isinstance(value, (np.integer, int, np.floating, float)):
        numeric = float(value)
        abs_numeric = abs(numeric)
        # heuristic based on magnitude: seconds/ms/us/ns epoch
        if abs_numeric >= 1e17:  # ns
            seconds = numeric / 1_000_000_000.0
        elif abs_numeric >= 1e14:  # us
            seconds = numeric / 1_000_000.0
        elif abs_numeric >= 1e11:  # ms
            seconds = numeric / 1_000.0
        else:  # seconds
            seconds = numeric
        return datetime.utcfromtimestamp(seconds).isoformat()

    if hasattr(value, "isoformat"):
        try:
            return str(value.isoformat())
        except Exception:
            pass
    return str(value)

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

def _write_risk_diagnostics_csv_json(
    *,
    run_dir: Path,
    timestamps: list[str],
    symbol_order: list[str],
    diagnostics: dict[str, np.ndarray],
) -> None:
    gross = np.asarray(diagnostics.get("gross_exposure", np.zeros(len(timestamps))), dtype=float)
    net = np.asarray(diagnostics.get("net_exposure", np.zeros(len(timestamps))), dtype=float)
    concentration = np.asarray(diagnostics.get("concentration", np.zeros(len(timestamps))), dtype=float)
    leverage = np.asarray(diagnostics.get("leverage_usage", np.zeros(len(timestamps))), dtype=float)
    turnover = np.asarray(diagnostics.get("turnover", np.zeros(len(timestamps))), dtype=float)
    turnover_by_symbol = np.asarray(
        diagnostics.get("turnover_by_symbol", np.zeros((len(timestamps), len(symbol_order)))),
        dtype=float,
    )

    rows = []
    for idx, ts in enumerate(timestamps):
        row = {
            "timestamp": ts,
            "gross_exposure": float(gross[idx]) if idx < gross.size else 0.0,
            "net_exposure": float(net[idx]) if idx < net.size else 0.0,
            "concentration": float(concentration[idx]) if idx < concentration.size else 0.0,
            "leverage_usage": float(leverage[idx]) if idx < leverage.size else 0.0,
            "turnover": float(turnover[idx]) if idx < turnover.size else 0.0,
        }
        rows.append(row)

    csv_path = run_dir / "risk_diagnostics.csv"
    with csv_path.open("w", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "timestamp",
                "gross_exposure",
                "net_exposure",
                "concentration",
                "leverage_usage",
                "turnover",
            ],
        )
        writer.writeheader()
        writer.writerows(rows)
    (run_dir / "risk_diagnostics.json").write_text(json.dumps(rows, indent=2))

    symbol_rows: list[dict[str, object]] = []
    for row_idx, ts in enumerate(timestamps):
        for col_idx, symbol in enumerate(symbol_order):
            value = 0.0
            if turnover_by_symbol.ndim == 2 and row_idx < turnover_by_symbol.shape[0] and col_idx < turnover_by_symbol.shape[1]:
                value = float(turnover_by_symbol[row_idx, col_idx])
            symbol_rows.append({"timestamp": ts, "symbol": symbol, "turnover": value})

    csv_sym = run_dir / "turnover_by_symbol.csv"
    with csv_sym.open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["timestamp", "symbol", "turnover"])
        writer.writeheader()
        writer.writerows(symbol_rows)
    (run_dir / "turnover_by_symbol.json").write_text(json.dumps(symbol_rows, indent=2))


def _build_slippage_model(*, model_name: str, costs_bps: float, params: dict[str, object] | None) -> object:
    cfg = dict(params or {})
    name = str(model_name).strip().lower()
    if name == "bps":
        return BpsSlippage(float(cfg.get("bps", costs_bps)))
    if name == "spread":
        return SpreadSlippage(float(cfg.get("spread_bps", costs_bps)))
    if name == "participation":
        calibration_path = cfg.get("calibration_path")
        if calibration_path is not None:
            from backtesting.execution import load_impact_calibration_buckets

            buckets = load_impact_calibration_buckets(str(calibration_path))
            return ParticipationImpactSlippage.from_calibration_buckets(
                buckets,
                base_bps=float(cfg.get("base_bps", 0.0)),
                participation_exponent=float(cfg.get("participation_exponent", 1.0)),
                max_participation=float(cfg.get("max_participation", 1.0)),
            )
        return ParticipationImpactSlippage(
            base_bps=float(cfg.get("base_bps", 0.0)),
            impact_coefficient_bps=float(cfg.get("impact_coefficient_bps", cfg.get("impact_bps", 20.0))),
            participation_exponent=float(cfg.get("participation_exponent", 1.0)),
            max_participation=float(cfg.get("max_participation", 1.0)),
        )
    if name == "volatility_scaled":
        return VolatilityScaledSlippage(
            base_bps=float(cfg.get("base_bps", costs_bps)),
            target_volatility=float(cfg.get("target_volatility", 0.01)),
            volatility_exponent=float(cfg.get("volatility_exponent", 1.0)),
            min_volatility=float(cfg.get("min_volatility", 1e-6)),
        )
    raise ValueError(f"Unknown execution model: {model_name}")


def _build_carry_model(*, model_name: str, params: dict[str, object] | None, n_assets: int, timeframe: str) -> object:
    cfg = dict(params or {})
    name = str(model_name).strip().lower()
    periods_per_year = 252.0 if timeframe.endswith("d") else 252.0 * 390.0
    if name == "short_borrow":
        return ShortBorrowCost(
            annual_borrow_rate=float(cfg.get("annual_borrow_rate", 0.0)),
            periods_per_year=float(cfg.get("periods_per_year", periods_per_year)),
        )
    if name == "asset_class":
        asset_classes = cfg.get("asset_classes")
        if not isinstance(asset_classes, list):
            asset_classes = ["equity"] * int(n_assets)
        return AssetClassCarryCost(
            asset_classes=asset_classes,
            annual_short_borrow_rates=cfg.get("annual_short_borrow_rates") if isinstance(cfg.get("annual_short_borrow_rates"), dict) else {"equity": 0.0},
            annual_long_financing_rates=cfg.get("annual_long_financing_rates") if isinstance(cfg.get("annual_long_financing_rates"), dict) else {},
            periods_per_year=float(cfg.get("periods_per_year", periods_per_year)),
        )
    raise ValueError(f"Unknown carry model: {model_name}")
def _parse_json_object(raw: str | None) -> dict[str, object]:
    if raw is None or not raw.strip():
        return {}
    parsed = json.loads(raw)
    if not isinstance(parsed, dict):
        raise ValueError("signal params must be a JSON object")
    return parsed


def _parse_json_array(raw: str | None) -> list[object]:
    if raw is None or not raw.strip():
        return []
    parsed = json.loads(raw)
    if not isinstance(parsed, list):
        raise ValueError("value must be a JSON array")
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
    run_parser.add_argument("--execution-model", choices=["bps", "spread", "participation", "volatility_scaled"], default="bps")
    run_parser.add_argument("--execution-model-params", default="{}", help="JSON object with execution model params.")
    run_parser.add_argument("--carry-model", choices=["short_borrow", "asset_class"], default="short_borrow")
    run_parser.add_argument("--carry-model-params", default="{}", help="JSON object with carry model params.")
    run_parser.add_argument("--entry-signal", default="ts_momentum", choices=["ts_momentum", "ma_trend", "breakout"])
    run_parser.add_argument("--entry-signal-params", default="{}", help="JSON object with entry signal parameters.")
    run_parser.add_argument("--exit-signal", default="none", choices=["none", "momentum_flip", "trailing_stop", "max_hold"])
    run_parser.add_argument("--exit-signal-params", default="{}", help="JSON object with exit signal parameters.")
    run_parser.add_argument("--signal-rebalance-interval", type=int, default=1, help="Only allow signal changes every N bars.")
    run_parser.add_argument("--starting-capital", type=float, default=100000.0)
    run_parser.add_argument("--bet-sizing-mode", choices=["kelly", "half_kelly", "custom"], default="half_kelly")
    run_parser.add_argument("--custom-bet-pct", type=float, default=10.0)
    run_parser.add_argument("--timeframe", default="1m", help="Bar resolution (e.g. 1m, 5m, 15m, 30m, 1h, 1d).")
    run_parser.add_argument("--strategy", choices=["momentum", "xsmom"], default="momentum")
    run_parser.add_argument("--xsmom-top-quantile", type=float, default=0.2)
    run_parser.add_argument("--xsmom-bottom-quantile", type=float, default=0.2)
    run_parser.add_argument("--xsmom-long-only", action="store_true")
    run_parser.add_argument("--xsmom-vol-lookback-days", type=int, default=20)
    run_parser.add_argument("--portfolio-method", choices=["equal_weight", "vol_target", "inverse_vol", "capped_optimization"], default="equal_weight")
    run_parser.add_argument("--portfolio-vol-lookback-bars", type=int, default=20)
    run_parser.add_argument("--portfolio-target-volatility", type=float, default=0.10)
    run_parser.add_argument("--portfolio-max-symbol-weight", type=float, default=0.25)
    run_parser.add_argument("--portfolio-max-sector-weight", type=float, default=0.60)
    run_parser.add_argument("--portfolio-max-gross-exposure", type=float, default=1.0)
    run_parser.add_argument("--portfolio-min-net-exposure", type=float, default=-1.0)
    run_parser.add_argument("--portfolio-max-net-exposure", type=float, default=1.0)

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

    wf_parser = subparsers.add_parser("walk_forward", help="Run rolling walk-forward optimization/evaluation.")
    _add_common_args(wf_parser)
    wf_parser.add_argument("--entry-grid", required=True, help="JSON mapping signal->list[params].")
    wf_parser.add_argument("--exit-grid", required=True, help="JSON mapping signal->list[params].")
    wf_parser.add_argument("--core-grid", required=True, help="JSON mapping core param->list[values].")
    wf_parser.add_argument("--train-bars", type=int, required=False)
    wf_parser.add_argument("--validation-bars", type=int, required=False)
    wf_parser.add_argument("--test-bars", type=int, required=False)
    wf_parser.add_argument("--step-bars", type=int, default=None)
    wf_parser.add_argument("--train-fraction", type=float, default=None)
    wf_parser.add_argument("--validation-fraction", type=float, default=None)
    wf_parser.add_argument("--test-fraction", type=float, default=None)
    wf_parser.add_argument("--step-fraction", type=float, default=None)
    wf_parser.add_argument("--score-metric", default="sharpe")
    wf_parser.add_argument("--purge-window-bars", type=int, default=0)
    wf_parser.add_argument("--embargo-window-bars", type=int, default=0)
    wf_parser.add_argument("--label-horizon-bars", type=int, default=1)
    wf_parser.add_argument("--nested-optimization", action="store_true")
    wf_parser.add_argument("--inner-train-fraction", type=float, default=0.7)

    opt_parser = subparsers.add_parser("optimize", help="Run constrained multi-objective optimization.")
    _add_common_args(opt_parser)
    opt_parser.add_argument("--entry-grid", required=True, help="JSON mapping signal->list[params].")
    opt_parser.add_argument("--exit-grid", required=True, help="JSON mapping signal->list[params].")
    opt_parser.add_argument("--core-grid", required=True, help="JSON mapping core param->list[values].")
    opt_parser.add_argument("--seed", type=int, default=42)
    opt_parser.add_argument("--n-trials", type=int, default=30)
    opt_parser.add_argument("--sampler", choices=["tpe", "random"], default="tpe")
    opt_parser.add_argument("--objectives", default='[{"name":"sharpe","sense":"maximize"},{"name":"turnover_total","sense":"minimize"},{"name":"max_drawdown","sense":"maximize"}]')
    opt_parser.add_argument("--max-turnover", type=float, default=None)
    opt_parser.add_argument("--max-drawdown-floor", type=float, default=None)
    opt_parser.add_argument("--min-trades", type=float, default=None)
    opt_parser.add_argument("--partial-period-fractions", default='[0.33,0.66,1.0]')

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
    elif args.command == "walk_forward":
        output = run_walk_forward_backtest(
            tickers=tickers,
            start_date=start_date,
            end_date=end_date,
            cache_root=cache_root,
            entry_grid={str(k): list(v) for k, v in _parse_json_object(args.entry_grid).items()},
            exit_grid={str(k): list(v) for k, v in _parse_json_object(args.exit_grid).items()},
            core_grid={str(k): list(v) for k, v in _parse_json_object(args.core_grid).items()},
            train_bars=None if args.train_bars is None else int(args.train_bars),
            validation_bars=None if args.validation_bars is None else int(args.validation_bars),
            test_bars=None if args.test_bars is None else int(args.test_bars),
            step_bars=args.step_bars,
            train_fraction=args.train_fraction,
            validation_fraction=args.validation_fraction,
            test_fraction=args.test_fraction,
            step_fraction=args.step_fraction,
            score_metric=str(args.score_metric),
            purge_window_bars=int(args.purge_window_bars),
            embargo_window_bars=int(args.embargo_window_bars),
            label_horizon_bars=int(args.label_horizon_bars),
            nested_optimization=bool(args.nested_optimization),
            inner_train_fraction=float(args.inner_train_fraction),
        )
    elif args.command == "optimize":
        output = run_strategy_optimization(
            tickers=tickers,
            start_date=start_date,
            end_date=end_date,
            cache_root=cache_root,
            entry_grid={str(k): list(v) for k, v in _parse_json_object(args.entry_grid).items()},
            exit_grid={str(k): list(v) for k, v in _parse_json_object(args.exit_grid).items()},
            core_grid={str(k): list(v) for k, v in _parse_json_object(args.core_grid).items()},
            seed=int(args.seed),
            n_trials=int(args.n_trials),
            sampler_name=str(args.sampler),
            objectives=list(_parse_json_array(args.objectives)),
            max_turnover=args.max_turnover,
            max_drawdown_floor=args.max_drawdown_floor,
            min_trades=args.min_trades,
            partial_period_fractions=[float(v) for v in _parse_json_array(args.partial_period_fractions)],
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
            execution_model=str(args.execution_model),
            execution_model_params=_parse_json_object(args.execution_model_params),
            carry_model=str(args.carry_model),
            carry_model_params=_parse_json_object(args.carry_model_params),
            entry_signal=args.entry_signal,
            entry_signal_params=_parse_json_object(args.entry_signal_params),
            exit_signal=args.exit_signal,
            exit_signal_params=_parse_json_object(args.exit_signal_params),
            signal_rebalance_interval=args.signal_rebalance_interval,
            starting_capital=args.starting_capital,
            bet_sizing_mode=args.bet_sizing_mode,
            custom_bet_pct=args.custom_bet_pct,
            strategy=args.strategy,
            xsmom_top_quantile=args.xsmom_top_quantile,
            xsmom_bottom_quantile=args.xsmom_bottom_quantile,
            xsmom_long_only=bool(args.xsmom_long_only),
            xsmom_vol_lookback_days=args.xsmom_vol_lookback_days,
            timeframe=str(args.timeframe),
            portfolio_method=str(args.portfolio_method),
            portfolio_vol_lookback_bars=int(args.portfolio_vol_lookback_bars),
            portfolio_target_volatility=float(args.portfolio_target_volatility),
            portfolio_max_symbol_weight=float(args.portfolio_max_symbol_weight),
            portfolio_max_sector_weight=float(args.portfolio_max_sector_weight),
            portfolio_max_gross_exposure=float(args.portfolio_max_gross_exposure),
            portfolio_min_net_exposure=float(args.portfolio_min_net_exposure),
            portfolio_max_net_exposure=float(args.portfolio_max_net_exposure),
        )
    print(output)


if __name__ == "__main__":
    main()

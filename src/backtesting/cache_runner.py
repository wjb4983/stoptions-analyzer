from __future__ import annotations

import csv
import json
import random
import argparse
import logging
import hashlib
import importlib.metadata
import time
import os
import platform
import socket
import subprocess
import sys
import threading
from concurrent.futures import FIRST_COMPLETED, Future, ProcessPoolExecutor, ThreadPoolExecutor, as_completed, wait
from dataclasses import dataclass
from datetime import date, datetime, timezone
from itertools import product
from pathlib import Path
from typing import Any, Callable

import numpy as np

from analysis.attribution import build_attribution_payload
from analysis.explainability import build_trade_explainability
from analysis.factor_exposure import build_factor_exposure_model
from analysis.reporting import (
    build_scenario_attribution_and_guardrails,
    build_backtest_robustness_report,
    build_drawdown_rows,
    build_sweep_robustness_report,
    format_backtest_report,
)
from backtesting.execution import (
    AssetClassCarryCost,
    BpsSlippage,
    BrokerFeeModel,
    CompositeSlippage,
    LatencyQueueDriftSlippage,
    ParticipationImpactSlippage,
    SquareRootImpactSlippage,
    ShortBorrowCost,
    SlippageCalibrationSelection,
    SpreadSlippage,
    VolatilityScaledSlippage,
    load_slippage_calibration_snapshots,
    select_slippage_calibration_snapshot,
)
from backtesting.vectorized import backtest_vectorized
from backtesting.perf import deflated_sharpe_ratio, probabilistic_sharpe_ratio
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
from backtesting.regimes import (
    RegimeScaledSlippageModel,
    apply_regime_risk_overlays,
    attribute_pnl_by_regime,
    compute_regime_labels,
    resolve_regime_parameters,
)
from backtesting.strategies import CrossSectionalMomentumConfig, build_cross_sectional_momentum_targets
from backtesting.strategies.ensemble import RegimeMetaPolicyConfig, build_regime_weight_schedule
from backtesting.walk_forward import (
    build_cpcv_walk_forward_folds,
    build_walk_forward_folds,
    persist_walk_forward_outputs,
    run_walk_forward_optimization,
)
from backtesting.validation import generate_combinatorial_purged_cv_splits, generate_purged_kfold_splits
from backtesting.optimization import (
    Constraint,
    Objective,
    BayesianSampler,
    CMASampler,
    GridSampler,
    OverfittingPenaltyConfig,
    RandomSampler,
    TPESampler,
    optimize,
)
from backtesting.experiment_registry import append_experiment_entry
from backtesting.monitoring import evaluate_drift_monitoring
from backtesting.attribution import write_attribution_artifacts
from backtesting.scenario_toolkit import list_scenario_pack_templates, resolve_scenario_pack_templates
from backtesting.regime_backtest_adapter import RegimeBacktestContract, RegimeBacktestOption, load_regime_backtest_contract


LOGGER = logging.getLogger(__name__)
CANONICAL_METRIC_SCHEMA_VERSION = "1.0"
RUN_MANIFEST_SCHEMA_VERSION = "2.0"
METRIC_TABLE_SCHEMA_VERSIONS: dict[str, str] = {
    "metrics": CANONICAL_METRIC_SCHEMA_VERSION,
    "attribution_timeseries": CANONICAL_METRIC_SCHEMA_VERSION,
    "attribution_summary": CANONICAL_METRIC_SCHEMA_VERSION,
    "leaderboard": CANONICAL_METRIC_SCHEMA_VERSION,
    "per_combo_summary": CANONICAL_METRIC_SCHEMA_VERSION,
}
PROMOTION_STATES = ("research", "paper", "shadow", "production")
PROMOTION_REQUIRED_CHECKS: dict[str, list[str]] = {
    "research": ["dataset_lock", "signal_diagnostics"],
    "paper": ["dataset_lock", "signal_diagnostics", "oos_periods", "stability_threshold", "drift_monitoring", "friction_adjusted_edge", "causal_robustness", "deflated_sharpe_reality_check", "parameter_stability_penalty", "train_validation_test_drift"],
    "shadow": ["dataset_lock", "signal_diagnostics", "oos_periods", "stability_threshold", "turnover_capacity", "drift_monitoring", "friction_adjusted_edge", "causal_robustness", "deflated_sharpe_reality_check", "parameter_stability_penalty", "train_validation_test_drift", "experiment_id"],
    "production": ["dataset_lock", "signal_diagnostics", "oos_periods", "stability_threshold", "turnover_capacity", "drift_monitoring", "friction_adjusted_edge", "causal_robustness", "deflated_sharpe_reality_check", "parameter_stability_penalty", "train_validation_test_drift", "experiment_id", "approval"],
}

BENCHMARK_BUY_HOLD = "buy_hold"
BENCHMARK_EQUAL_WEIGHT_MOMENTUM = "equal_weight_momentum"
BENCHMARK_VOLATILITY_PARITY = "volatility_parity"
DEFAULT_BENCHMARK_SELECTION: tuple[str, ...] = (
    BENCHMARK_BUY_HOLD,
    BENCHMARK_EQUAL_WEIGHT_MOMENTUM,
    BENCHMARK_VOLATILITY_PARITY,
)


class TaskCancellationError(RuntimeError):
    """Raised when a cooperative cancellation request interrupts a workflow."""


class CancellationToken:
    def __init__(self) -> None:
        self._event = threading.Event()
        self._lock = threading.Lock()
        self._reason: str | None = None
        self._requested_at: str | None = None

    def cancel(self, reason: str = "Cancellation requested") -> None:
        with self._lock:
            if self._event.is_set():
                return
            self._reason = str(reason)
            self._requested_at = datetime.now(timezone.utc).isoformat()
            self._event.set()

    @property
    def is_canceled(self) -> bool:
        return self._event.is_set()

    def checkpoint(self, location: str) -> None:
        if not self._event.is_set():
            return
        reason = self._reason or "Cancellation requested"
        raise TaskCancellationError(f"{reason} at {location}")

    def snapshot(self) -> dict[str, Any]:
        return {
            "requested": self.is_canceled,
            "requested_at": self._requested_at,
            "reason": self._reason,
        }


def _resolve_benchmark_selection(selected: list[str] | None) -> list[str]:
    allowed = set(DEFAULT_BENCHMARK_SELECTION)
    if not selected:
        return list(DEFAULT_BENCHMARK_SELECTION)
    normalized: list[str] = []
    for name in selected:
        key = str(name).strip().lower().replace("-", "_").replace(" ", "_")
        if key in allowed and key not in normalized:
            normalized.append(key)
    return normalized or list(DEFAULT_BENCHMARK_SELECTION)


def _build_benchmark_signals(*, benchmark: str, prices: np.ndarray, lookback_days: int, skip_days: int) -> np.ndarray:
    n_periods, n_assets = prices.shape
    if benchmark == BENCHMARK_BUY_HOLD:
        return np.ones((n_periods, n_assets), dtype=float)
    if benchmark == BENCHMARK_EQUAL_WEIGHT_MOMENTUM:
        entry_cfg = parse_entry_signal_config(
            "ts_momentum",
            {"lookback_days": int(lookback_days), "skip_days": int(skip_days), "long_only": True},
            default_lookback_days=int(lookback_days),
            default_skip_days=int(skip_days),
        )
        exit_cfg = parse_exit_signal_config("none", {}, default_lookback_days=int(lookback_days), default_skip_days=int(skip_days))
        return build_targets(close_prices=prices, missing_mask=np.zeros_like(prices, dtype=bool), entry_config=entry_cfg, exit_config=exit_cfg)
    if benchmark == BENCHMARK_VOLATILITY_PARITY:
        returns = np.zeros_like(prices, dtype=float)
        returns[1:] = prices[1:] / np.where(prices[:-1] == 0.0, 1.0, prices[:-1]) - 1.0
        lookback = max(5, int(lookback_days))
        signals = np.zeros_like(prices, dtype=float)
        for idx in range(lookback, n_periods):
            window = returns[max(1, idx - lookback + 1): idx + 1]
            vol = np.std(window, axis=0)
            inv_vol = np.where(vol > 1e-8, 1.0 / vol, 0.0)
            denom = float(np.sum(inv_vol))
            if denom <= 1e-12:
                signals[idx] = np.full(n_assets, 1.0 / max(1, n_assets), dtype=float)
            else:
                signals[idx] = inv_vol / denom
        if n_periods:
            signals[:lookback] = np.full((min(lookback, n_periods), n_assets), 1.0 / max(1, n_assets), dtype=float)
        return signals
    raise ValueError(f"Unsupported benchmark: {benchmark}")


def _compute_relative_alpha_ir(candidate_returns: np.ndarray, benchmark_returns: np.ndarray) -> dict[str, float]:
    candidate = np.asarray(candidate_returns, dtype=float).reshape(-1)
    benchmark = np.asarray(benchmark_returns, dtype=float).reshape(-1)
    size = min(candidate.size, benchmark.size)
    if size == 0:
        return {"alpha": 0.0, "information_ratio": 0.0, "tracking_error": 0.0}
    active = candidate[:size] - benchmark[:size]
    alpha = float(np.mean(active))
    tracking_error = float(np.std(active))
    ir = 0.0 if tracking_error <= 1e-12 else float(alpha / tracking_error)
    return {"alpha": alpha, "information_ratio": ir, "tracking_error": tracking_error}


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




def run_trained_regime_backtest(
    *,
    tickers: list[str],
    start_date: date,
    end_date: date,
    cache_root: Path,
    timeframe: str,
    regime_contract: RegimeBacktestContract | None = None,
    regime_manifest_path: str | Path | None = None,
    regime_source: str | None = None,
    governance_metadata: dict[str, Any] | None = None,
    stress_controls: dict[str, Any] | None = None,
    scenario_packs: list[str] | None = None,
    selected_test_suite: str = "custom",
    suite_composition: dict[str, Any] | None = None,
) -> str:
    contract = regime_contract
    if contract is None:
        if regime_manifest_path is None:
            raise ValueError("Either regime_contract or regime_manifest_path must be provided.")
        manifest_path = Path(regime_manifest_path)
        inferred_source = regime_source or ("bundle" if manifest_path.name == "bundle_manifest.json" else "training_run")
        option = RegimeBacktestOption(
            option_id=f"direct:{manifest_path.stem}",
            label=str(manifest_path),
            source=inferred_source,
            manifest_path=str(manifest_path),
        )
        contract = load_regime_backtest_contract(option)

    defaults = dict(contract.defaults)
    strategy = str(defaults.get("strategy", "momentum")).strip() or "momentum"
    lookback_days = int(float(defaults.get("lookback_days", 90) or 90))
    skip_days = int(float(defaults.get("skip_days", 5) or 5))

    output_text = run_time_series_momentum_backtest(
        tickers=tickers,
        start_date=start_date,
        end_date=end_date,
        cache_root=cache_root,
        lookback_days=lookback_days,
        skip_days=skip_days,
        costs_bps=5.0,
        strategy=strategy,
        timeframe=timeframe,
        portfolio_max_gross_exposure=float(defaults.get("portfolio_max_gross_exposure", 1.0) or 1.0),
        portfolio_min_net_exposure=float(defaults.get("portfolio_min_net_exposure", -1.0) or -1.0),
        portfolio_max_net_exposure=float(defaults.get("portfolio_max_net_exposure", 1.0) or 1.0),
        portfolio_max_symbol_weight=float(defaults.get("portfolio_max_symbol_weight", 0.25) or 0.25),
        portfolio_max_sector_weight=float(defaults.get("portfolio_max_sector_weight", 0.60) or 0.60),
        governance_metadata=dict(governance_metadata or {}),
        stress_controls=dict(stress_controls or {}),
        scenario_packs=list(scenario_packs or []),
        selected_test_suite=selected_test_suite,
        suite_composition=dict(suite_composition or {}),
    )
    artifacts = contract.execution_artifacts
    run_dir = _extract_run_dir_from_output(output_text)
    if run_dir is not None:
        _write_suite_artifact_bundle(
            run_dir=run_dir,
            suite_key=selected_test_suite,
            suite_composition=dict(suite_composition or {}),
            governance_metadata=dict(governance_metadata or {}),
            stress_controls=dict(stress_controls or {}),
        )
    return (
        output_text
        + "\nRegime contract: " + contract.regime_name
        + f" ({contract.source})"
        + "\nManifest: " + contract.manifest_path
        + "\nChampion model IDs: " + json.dumps(artifacts.get("champion_model_ids", {}), sort_keys=True)
        + "\nModel paths: " + json.dumps(artifacts.get("model_paths", {}), sort_keys=True)
        + "\nCalibration paths: " + json.dumps(artifacts.get("calibration_paths", {}), sort_keys=True)
        + "\nTest suite: " + str(selected_test_suite)
        + "\nSuite composition: " + json.dumps(dict(suite_composition or {}), sort_keys=True)
    )

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
    portfolio_rebalance_frequency_bars: int = 1,
    portfolio_clustering_linkage: str = "single",
    portfolio_covariance_shrinkage: float = 0.15,
    portfolio_max_gross_exposure: float = 1.0,
    portfolio_min_net_exposure: float = -1.0,
    portfolio_max_net_exposure: float = 1.0,
    portfolio_max_net_gamma: float | None = None,
    portfolio_max_abs_vega_bucket: float | None = None,
    portfolio_max_abs_delta_per_underlying: float | None = None,
    portfolio_sector_map: dict[str, str] | None = None,
    regime_parameter_map: dict[str, dict[str, object]] | None = None,
    regime_risk_map: dict[str, dict[str, float]] | None = None,
    regime_leverage_multipliers: dict[str, float] | None = None,
    regime_cost_multipliers: dict[str, float] | None = None,
    governance_metadata: dict[str, Any] | None = None,
    stress_controls: dict[str, Any] | None = None,
    scenario_packs: list[str] | None = None,
    capacity_aum_scales: list[float] | None = None,
    max_participation_rate: float | None = None,
    random_seed: int = 42,
    preflight_config: PreflightValidationConfig | None = None,
    benchmarks: list[str] | None = None,
    selected_test_suite: str = "custom",
    suite_composition: dict[str, Any] | None = None,
) -> str:
    BACKTEST_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    random.seed(int(random_seed))
    np.random.seed(int(random_seed))
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
    _run_preflight_or_raise(
        arrays=arrays,
        requested_tickers=tickers,
        start_dt=start_dt,
        end_dt=end_dt,
        timeframe=timeframe,
        config=preflight_config or PreflightValidationConfig(),
        workflow_label="run",
    )

    prices = _fill_missing_prices(arrays.close_prices)
    bet_fraction = _resolve_bet_fraction(
        prices=prices,
        mode=bet_sizing_mode,
        custom_bet_pct=custom_bet_pct,
    )

    regime_state = compute_regime_labels(prices)
    regime_labels = np.asarray(regime_state["labels"], dtype=object)
    regime_probabilities = np.asarray(
        regime_state.get("regime_probabilities", np.zeros((prices.shape[0], 0), dtype=float)),
        dtype=float,
    )
    regime_states = np.asarray(regime_state.get("regime_states", np.array([], dtype=object)), dtype=object)

    if strategy == "xsmom":
        base_params = {
            "lookback_days": lookback_days,
            "skip_days": skip_days,
            "top_quantile": float(xsmom_top_quantile),
            "bottom_quantile": float(xsmom_bottom_quantile),
            "long_only": bool(xsmom_long_only),
            "vol_lookback_days": int(xsmom_vol_lookback_days),
            "rebalance_interval": max(1, int(signal_rebalance_interval)),
        }
        signals = np.zeros_like(prices, dtype=float)
        unique_regimes = sorted(set(str(item) for item in regime_labels.tolist()))
        for regime in unique_regimes:
            params = resolve_regime_parameters(
                base_params=base_params,
                regime_label=regime,
                parameter_map=regime_parameter_map,
            )
            xs_cfg = CrossSectionalMomentumConfig(
                lookback_days=int(params["lookback_days"]),
                skip_days=int(params["skip_days"]),
                top_quantile=float(params["top_quantile"]),
                bottom_quantile=float(params["bottom_quantile"]),
                long_only=bool(params["long_only"]),
                vol_lookback_days=int(params["vol_lookback_days"]),
                rebalance_interval=int(params["rebalance_interval"]),
            )
            candidate = build_cross_sectional_momentum_targets(
                close_prices=prices,
                missing_mask=arrays.missing_mask,
                config=xs_cfg,
            )
            mask = regime_labels == regime
            signals[mask] = candidate[mask]
        sized_signals = signals * float(bet_fraction)
    else:
        base_entry = dict(entry_signal_params or {})
        base_exit = dict(exit_signal_params or {})
        signals = np.zeros_like(prices, dtype=float)
        unique_regimes = sorted(set(str(item) for item in regime_labels.tolist()))
        for regime in unique_regimes:
            resolved = resolve_regime_parameters(
                base_params={**base_entry, **base_exit},
                regime_label=regime,
                parameter_map=regime_parameter_map,
            )
            entry_params = {**base_entry, **{k: v for k, v in resolved.items() if k in {"lookback_days", "skip_days", "threshold", "window", "fast_days", "slow_days", "zscore_threshold", "atr_mult"}}}
            exit_params = {**base_exit, **{k: v for k, v in resolved.items() if k in {"lookback_days", "skip_days", "threshold", "window", "fast_days", "slow_days", "zscore_threshold", "atr_mult", "max_hold_days"}}}
            regime_entry_cfg = parse_entry_signal_config(
                entry_signal,
                entry_params,
                default_lookback_days=lookback_days,
                default_skip_days=skip_days,
            )
            regime_exit_cfg = parse_exit_signal_config(
                exit_signal,
                exit_params,
                default_lookback_days=lookback_days,
                default_skip_days=skip_days,
            )
            candidate = build_targets(
                close_prices=prices,
                missing_mask=arrays.missing_mask,
                entry_config=regime_entry_cfg,
                exit_config=regime_exit_cfg,
            )
            mask = regime_labels == regime
            signals[mask] = candidate[mask]

        signals = _throttle_signal_changes(signals, interval=max(1, int(signal_rebalance_interval)))

        trend_cfg = parse_entry_signal_config("ma_trend", {"ma_window": 30}, default_lookback_days=lookback_days, default_skip_days=skip_days)
        meanrev_cfg = parse_entry_signal_config("mean_reversion", {"lookback_days": 20, "zscore_threshold": 1.0}, default_lookback_days=lookback_days, default_skip_days=skip_days)
        no_exit_cfg = parse_exit_signal_config("none", {}, default_lookback_days=lookback_days, default_skip_days=skip_days)
        trend_targets = build_targets(close_prices=prices, missing_mask=arrays.missing_mask, entry_config=trend_cfg, exit_config=no_exit_cfg)
        meanrev_targets = build_targets(close_prices=prices, missing_mask=arrays.missing_mask, entry_config=meanrev_cfg, exit_config=no_exit_cfg)
        base_sleeve = np.nanmean(signals, axis=1)
        trend_sleeve = np.nanmean(trend_targets, axis=1)
        meanrev_sleeve = np.nanmean(meanrev_targets, axis=1)
        regime_confidence = np.asarray(regime_state.get("regime_confidence", np.ones(prices.shape[0], dtype=float)), dtype=float)
        routed = np.zeros_like(base_sleeve)
        for idx in range(base_sleeve.size):
            label = str(regime_labels[idx])
            confidence = float(np.clip(regime_confidence[idx], 0.0, 1.0))
            if "macro_risk_off" in label or "vol_high" in label:
                routed[idx] = (1.0 - confidence) * base_sleeve[idx] + confidence * meanrev_sleeve[idx]
            elif "trend_up" in label and "macro_risk_on" in label:
                routed[idx] = (1.0 - confidence) * base_sleeve[idx] + confidence * trend_sleeve[idx]
            else:
                routed[idx] = 0.5 * base_sleeve[idx] + 0.25 * trend_sleeve[idx] + 0.25 * meanrev_sleeve[idx]
        signals = np.repeat(routed[:, None], prices.shape[1], axis=1)

        sized_signals = _apply_discrete_bet_sizing(
            signals=signals,
            prices=prices,
            starting_capital=starting_capital,
            bet_fraction=bet_fraction,
        )

    slippage, calibration_selection = _build_slippage_model(
        model_name=execution_model,
        costs_bps=costs_bps,
        params=execution_model_params,
        as_of_date=end_date.isoformat(),
    )
    slippage = RegimeScaledSlippageModel(slippage, regime_labels, regime_cost_multipliers)
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
        rebalance_frequency_bars=int(portfolio_rebalance_frequency_bars),
        clustering_linkage=str(portfolio_clustering_linkage),
        covariance_shrinkage=float(portfolio_covariance_shrinkage),
        max_gross_exposure=float(portfolio_max_gross_exposure),
        min_net_exposure=float(portfolio_min_net_exposure),
        max_net_exposure=float(portfolio_max_net_exposure),
        max_net_gamma=None if portfolio_max_net_gamma is None else float(portfolio_max_net_gamma),
        max_abs_vega_bucket=None if portfolio_max_abs_vega_bucket is None else float(portfolio_max_abs_vega_bucket),
        max_abs_delta_per_underlying=None if portfolio_max_abs_delta_per_underlying is None else float(portfolio_max_abs_delta_per_underlying),
        sector_map=dict(portfolio_sector_map or {}),
        use_residual_signals=True,
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
        underlying_by_symbol={
            symbol: getattr(arrays.metadata, "underlying_by_symbol", {}).get(symbol, symbol)
            for symbol in symbol_order
        },
    )

    borrow_rate_series = None
    borrow_available_flags = None
    if isinstance(carry_model_params, dict):
        if "borrow_rate_series" in carry_model_params:
            borrow_rate_series = np.asarray(carry_model_params.get("borrow_rate_series"), dtype=float)
        if "borrow_available_flags" in carry_model_params:
            borrow_available_flags = np.asarray(carry_model_params.get("borrow_available_flags"), dtype=bool)

    adjusted_weights, regime_diag = apply_regime_risk_overlays(
        weights=portfolio_result.target_weights,
        regime_labels=regime_labels,
        risk_map=regime_risk_map,
        regime_probabilities=regime_probabilities,
        regime_states=regime_states,
        state_to_label=np.asarray(regime_state.get("regime_state_to_legacy_label", np.array([], dtype=object)), dtype=object),
        leverage_multipliers=regime_leverage_multipliers,
    )
    if (
        regime_probabilities.ndim == 2
        and regime_probabilities.shape[0] == prices.shape[0]
        and regime_states.size == regime_probabilities.shape[1]
    ):
        for idx, state_name in enumerate(regime_states.tolist()):
            regime_diag[f"regime_probability_{state_name}"] = regime_probabilities[:, idx]
    regime_diag["regime_confidence"] = np.asarray(regime_state.get("regime_confidence", np.zeros(prices.shape[0], dtype=float)), dtype=float)
    portfolio_result.diagnostics.update(regime_diag)

    fee_model = BrokerFeeModel(
        fee_bps=float((execution_model_params or {}).get("broker_fee_bps", 0.0)),
        fee_per_unit=float((execution_model_params or {}).get("broker_fee_per_unit", 0.0)),
        minimum_fee=float((execution_model_params or {}).get("broker_min_fee", 0.0)),
    )

    result = backtest_vectorized(
        prices=prices,
        signals=adjusted_weights,
        slippage_model=slippage,
        fee_model=fee_model,
        borrow_cost_model=borrow,
        volumes=getattr(arrays, "volume", np.maximum(arrays.close_prices, 1.0)),
        adv=getattr(arrays, "adv", np.maximum(arrays.close_prices, 1.0)),
        volatility=getattr(arrays, "realized_volatility", np.zeros_like(arrays.close_prices)),
        spread_bps=getattr(arrays, "spread_bps", np.full_like(arrays.close_prices, float(costs_bps))),
        order_type=str((execution_model_params or {}).get("order_type", "market")),
        latency_bars=getattr(arrays, "latency_bars", np.zeros_like(arrays.close_prices)),
        latency_ms=getattr(arrays, "latency_ms", np.zeros_like(arrays.close_prices)),
        queue_rank_proxy=getattr(arrays, "queue_rank_proxy", np.full_like(arrays.close_prices, 0.5)),
        available_bar_volume=getattr(arrays, "available_bar_volume", np.maximum(arrays.close_prices, 1.0)),
        max_participation_per_bar=getattr(arrays, "max_participation", np.ones_like(arrays.close_prices)),
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
        corporate_action_splits=arrays.split_factors,
        corporate_action_dividends=arrays.dividends,
    )

    friction_off_result = backtest_vectorized(
        prices=prices,
        signals=adjusted_weights,
        slippage_model=BpsSlippage(0.0),
        fee_model=BrokerFeeModel(fee_bps=0.0, fee_per_unit=0.0, minimum_fee=0.0),
        borrow_cost_model=ShortBorrowCost(annual_borrow_rate=0.0),
        volumes=getattr(arrays, "volume", np.maximum(arrays.close_prices, 1.0)),
        adv=getattr(arrays, "adv", np.maximum(arrays.close_prices, 1.0)),
        volatility=getattr(arrays, "realized_volatility", np.zeros_like(arrays.close_prices)),
        spread_bps=getattr(arrays, "spread_bps", np.full_like(arrays.close_prices, float(costs_bps))),
        order_type=str((execution_model_params or {}).get("order_type", "market")),
        latency_bars=getattr(arrays, "latency_bars", np.zeros_like(arrays.close_prices)),
        latency_ms=getattr(arrays, "latency_ms", np.zeros_like(arrays.close_prices)),
        queue_rank_proxy=getattr(arrays, "queue_rank_proxy", np.full_like(arrays.close_prices, 0.5)),
        available_bar_volume=getattr(arrays, "available_bar_volume", np.maximum(arrays.close_prices, 1.0)),
        max_participation_per_bar=getattr(arrays, "max_participation", np.ones_like(arrays.close_prices)),
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
        corporate_action_splits=arrays.split_factors,
        corporate_action_dividends=arrays.dividends,
    )

    selected_benchmarks = _resolve_benchmark_selection(benchmarks)
    benchmark_rows: list[dict[str, Any]] = []
    benchmark_lookup: dict[str, dict[str, float]] = {}
    for benchmark_name in selected_benchmarks:
        benchmark_signals = _build_benchmark_signals(
            benchmark=benchmark_name,
            prices=prices,
            lookback_days=lookback_days,
            skip_days=skip_days,
        )
        benchmark_result = backtest_vectorized(
            prices=prices,
            signals=benchmark_signals,
            slippage_model=BpsSlippage(0.0),
            fee_model=BrokerFeeModel(fee_bps=0.0, fee_per_unit=0.0, minimum_fee=0.0),
            borrow_cost_model=ShortBorrowCost(annual_borrow_rate=0.0),
            volumes=getattr(arrays, "volume", np.maximum(arrays.close_prices, 1.0)),
            adv=getattr(arrays, "adv", np.maximum(arrays.close_prices, 1.0)),
            volatility=getattr(arrays, "realized_volatility", np.zeros_like(arrays.close_prices)),
            spread_bps=getattr(arrays, "spread_bps", np.full_like(arrays.close_prices, float(costs_bps))),
            order_type=str((execution_model_params or {}).get("order_type", "market")),
            latency_bars=getattr(arrays, "latency_bars", np.zeros_like(arrays.close_prices)),
            latency_ms=getattr(arrays, "latency_ms", np.zeros_like(arrays.close_prices)),
            queue_rank_proxy=getattr(arrays, "queue_rank_proxy", np.full_like(arrays.close_prices, 0.5)),
            available_bar_volume=getattr(arrays, "available_bar_volume", np.maximum(arrays.close_prices, 1.0)),
            max_participation_per_bar=getattr(arrays, "max_participation", np.ones_like(arrays.close_prices)),
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
            corporate_action_splits=arrays.split_factors,
            corporate_action_dividends=arrays.dividends,
        )
        rel = _compute_relative_alpha_ir(
            candidate_returns=_to_numpy_1d(result.returns),
            benchmark_returns=_to_numpy_1d(benchmark_result.returns),
        )
        benchmark_row = {
            "benchmark": benchmark_name,
            "total_return": float(benchmark_result.metrics.get("total_return", 0.0)),
            "sharpe": float(benchmark_result.metrics.get("sharpe", 0.0)),
            "volatility": float(benchmark_result.metrics.get("volatility", 0.0)),
            "alpha": float(rel["alpha"]),
            "information_ratio": float(rel["information_ratio"]),
            "tracking_error": float(rel["tracking_error"]),
        }
        benchmark_rows.append(benchmark_row)
        benchmark_lookup[benchmark_name] = benchmark_row

    governance_payload = _build_governance_metadata(governance_metadata)

    friction_edge = _friction_adjusted_edge_checks(
        friction_on_metrics={k: float(v) for k, v in result.metrics.items() if isinstance(v, (int, float))},
        friction_off_metrics={k: float(v) for k, v in friction_off_result.metrics.items() if isinstance(v, (int, float))},
        governance_payload=governance_payload,
    )
    governance_payload["gate_checks"]["friction_adjusted_edge"] = bool(friction_edge["pass"])

    timestamps = arrays.date_index

    equity = _to_numpy_1d(result.equity_curve)
    returns = _to_numpy_1d(result.returns)
    pnl = _to_numpy_1d(result.pnl)
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
    factor_model = build_factor_exposure_model(prices=prices)
    attribution_payload = build_attribution_payload(
        timestamps=np.asarray(timestamps),
        prices=prices,
        positions=_to_numpy_2d(result.positions),
        slippage_drag=_to_numpy_1d(result.cost_breakdown.get("slippage", np.zeros_like(returns))),
        fee_drag=_to_numpy_1d(result.cost_breakdown.get("fees", np.zeros_like(returns))),
        borrow_drag=_to_numpy_1d(result.cost_breakdown.get("borrow", np.zeros_like(returns))),
        factor_exposures=factor_model.exposures_by_asset,
        factor_returns=factor_model.factor_returns,
    )

    metrics = dict(result.metrics)
    metrics["turnover_total"] = turnover_stats["total"]
    metrics["cost_total"] = cost_totals.get("total", 0.0)
    metrics["friction_adjusted_edge"] = float(friction_edge["friction_adjusted_edge"])
    metrics["friction_edge_retention"] = float(friction_edge["friction_edge_retention"])
    for benchmark_name, benchmark_row in benchmark_lookup.items():
        metrics[f"alpha_vs_{benchmark_name}"] = float(benchmark_row.get("alpha", 0.0))
        metrics[f"ir_vs_{benchmark_name}"] = float(benchmark_row.get("information_ratio", 0.0))

    regime_pnl_attribution = attribute_pnl_by_regime(pnl=pnl, regime_labels=regime_labels)
    regime_ensemble_report = _build_regime_ensemble_comparison(
        prices=prices,
        missing_mask=arrays.missing_mask,
        regime_labels=regime_labels,
    )

    parameter_payload = {
        "tickers": list(tickers),
        "lookback_days": lookback_days,
        "skip_days": skip_days,
        "costs_bps": costs_bps,
        "execution_model": execution_model,
        "execution_model_params": execution_model_params or {},
        "benchmarks": benchmark_rows,
        "friction_backtests": {
            "friction_on": {"metrics": {k: float(v) for k, v in result.metrics.items() if isinstance(v, (int, float))}},
            "friction_off": {"metrics": {k: float(v) for k, v in friction_off_result.metrics.items() if isinstance(v, (int, float))}},
            "edge": friction_edge,
        },
        "execution_model_calibration": {
            "source": calibration_selection.source,
            "effective_date": calibration_selection.effective_date,
            "warning_flags": calibration_selection.warning_flags,
            "resolved_params": calibration_selection.params,
        },
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
        "portfolio_rebalance_frequency_bars": portfolio_rebalance_frequency_bars,
        "portfolio_clustering_linkage": portfolio_clustering_linkage,
        "portfolio_covariance_shrinkage": portfolio_covariance_shrinkage,
        "portfolio_max_gross_exposure": portfolio_max_gross_exposure,
        "portfolio_min_net_exposure": portfolio_min_net_exposure,
        "portfolio_max_net_exposure": portfolio_max_net_exposure,
        "portfolio_max_net_gamma": portfolio_max_net_gamma,
        "portfolio_max_abs_vega_bucket": portfolio_max_abs_vega_bucket,
        "portfolio_max_abs_delta_per_underlying": portfolio_max_abs_delta_per_underlying,
        "portfolio_sector_map": dict(portfolio_sector_map or {}),
        "regime_parameter_map": dict(regime_parameter_map or {}),
        "regime_risk_map": dict(regime_risk_map or {}),
        "regime_leverage_multipliers": dict(regime_leverage_multipliers or {}),
        "regime_cost_multipliers": dict(regime_cost_multipliers or {}),
        "cache_root": str(cache_root),
        "governance": governance_payload,
        "stress_controls": dict(stress_controls or {}),
        "scenario_packs": [str(pack) for pack in (scenario_packs or [])],
        "selected_test_suite": str(selected_test_suite or "custom"),
        "suite_composition": dict(suite_composition or {}),
    }

    fill_rows = list(result.fills)
    robustness_report = build_backtest_robustness_report(
        returns=returns,
        metrics=metrics,
        turnover_stats=turnover_stats,
        cost_totals=cost_totals,
        fills=fill_rows,
        capacity_scales=None if capacity_aum_scales is None else np.asarray(capacity_aum_scales, dtype=float),
        max_participation_rate=max_participation_rate,
    )

    scenario_definitions = _build_stress_scenario_definitions(
        timestamps=timestamps,
        returns=returns,
        controls=stress_controls,
        scenario_packs=scenario_packs,
    )
    scenario_results = _run_stress_scenario_wrappers(
        returns=returns,
        prices=prices,
        scenario_definitions=scenario_definitions,
    )
    scenario_payload = build_scenario_attribution_and_guardrails(
        baseline_metrics=metrics,
        scenario_results=scenario_results,
    )
    scenario_payload.update(_stress_gate_summary(scenario_payload, controls=stress_controls))

    account_state = result.cost_breakdown.get("account_state", {})
    merged_risk_diagnostics = dict(portfolio_result.diagnostics)
    for key in (
        "cash",
        "margin_requirement",
        "excess_liquidity",
        "buying_power",
        "margin_utilization",
        "forced_liquidation",
        "deleveraging_scale",
    ):
        if key in account_state:
            merged_risk_diagnostics[key] = _to_numpy_1d(account_state[key])

    data_snapshot = _build_data_snapshot_identifiers(arrays=arrays, cache_root=cache_root, timeframe=timeframe)
    run_dir = _persist_backtest_outputs(
        timestamps=timestamps,
        symbol_order=symbol_order,
        equity=equity,
        returns=returns,
        trades=trades,
        risk_diagnostics=merged_risk_diagnostics,
        metrics=metrics,
        dataset_contracts=dataset_contracts,
        parameters=parameter_payload,
        data_snapshot=data_snapshot,
        random_seed=int(random_seed),
        robustness_report=robustness_report,
        scenario_payload=scenario_payload,
        regime_labels=regime_labels,
        regime_probabilities=regime_probabilities,
        regime_states=regime_states,
        regime_pnl_attribution=regime_pnl_attribution,
        regime_ensemble_report=regime_ensemble_report,
        governance=governance_payload,
        corporate_action_splits=arrays.split_factors,
        corporate_action_dividends=arrays.dividends,
        attribution_payload={"time_series": attribution_payload.time_series, "summary": attribution_payload.summary},
        fill_rows=fill_rows,
        slippage_calibration_selection={
            "source": calibration_selection.source,
            "effective_date": calibration_selection.effective_date,
            "warning_flags": list(calibration_selection.warning_flags),
        },
        selected_test_suite=str(selected_test_suite or "custom"),
        suite_composition=dict(suite_composition or {}),
    )

    trade_log_rows = _build_trade_log_rows(
        timestamps=timestamps,
        symbol_order=symbol_order,
        prices=prices,
        trades=trades,
        costs_bps=costs_bps,
        starting_capital=starting_capital,
    )
    explainability_rows = _build_trade_explainability_rows(
        trade_log_rows=trade_log_rows,
        fill_rows=fill_rows,
        risk_diagnostics=merged_risk_diagnostics,
        costs_bps=costs_bps,
    )
    explainability_report = _format_explainability_report(explainability_rows)
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
        run_details=_build_report_run_details(run_dir),
        robustness_report=robustness_report,
    )
    (run_dir / "trade_log.csv").write_text(_trade_log_csv(trade_log_rows))
    (run_dir / "trade_log.json").write_text(json.dumps(trade_log_rows, indent=2))
    (run_dir / "trade_explainability.json").write_text(json.dumps(explainability_rows, indent=2))
    (run_dir / "trade_explainability_report.txt").write_text(explainability_report)
    (run_dir / "fills.csv").write_text(_fills_csv(fill_rows))
    (run_dir / "fills.json").write_text(json.dumps(fill_rows, indent=2))

    ensemble_lines = ["Regime Ensemble Comparison:"]
    for row in regime_ensemble_report.get("comparison", []):
        ensemble_lines.append(
            f"- {row['policy']}: total_return={row['total_return']:.6f} sharpe={row['sharpe']:.6f} max_drawdown={row['max_drawdown']:.6f}"
        )
    if regime_ensemble_report.get("uplift_by_regime_bucket"):
        ensemble_lines.append("Regime routing uplift by bucket (weighted vs baseline):")
        for row in regime_ensemble_report.get("uplift_by_regime_bucket", []):
            if not isinstance(row, dict):
                continue
            ensemble_lines.append(
                f"- {row.get('regime', '')}: uplift_pnl_mean={float(row.get('uplift_pnl_mean', 0.0)):.8f} bars={int(row.get('bars', 0))}"
            )
    benchmark_lines = ["Benchmark comparison:"]
    if benchmark_rows:
        for row in benchmark_rows:
            benchmark_lines.append(
                f"- {row['benchmark']}: total_return={row['total_return']:.6f} sharpe={row['sharpe']:.6f} alpha={row['alpha']:.6f} ir={row['information_ratio']:.6f}"
            )
    else:
        benchmark_lines.append("- none")

    final_report = report_text + "\n\n" + trade_log_summary + "\n\n" + explainability_report + "\n\n" + "\n".join(ensemble_lines) + "\n\n" + "\n".join(benchmark_lines)
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


@dataclass(frozen=True)
class SweepRetryPolicy:
    max_attempts: int = 2
    stale_worker_timeout_seconds: float = 120.0


@dataclass(frozen=True)
class SweepWorkerConfig:
    mode: str = "local"
    max_workers: int = 1
    remote_endpoints: tuple[str, ...] = ()


@dataclass(frozen=True)
class PreflightValidationConfig:
    """Configurable checks run before backtest workflow execution."""

    max_missing_bars_ratio: float = 1.0
    min_symbol_coverage_ratio: float = 1.0
    critical_checks: tuple[str, ...] = (
        "timestamp_consistency",
        "adjustment_flags",
        "symbol_coverage",
    )
    block_on_critical: bool = True


class _SweepExecutorAdapter:
    def submit(self, fn: Callable[[dict[str, Any]], dict[str, Any]], payload: dict[str, Any]) -> Future[dict[str, Any]]:
        raise NotImplementedError

    def shutdown(self) -> None:
        raise NotImplementedError


class _LocalSweepExecutorAdapter(_SweepExecutorAdapter):
    def __init__(self, *, max_workers: int, use_process_pool: bool) -> None:
        cls = ProcessPoolExecutor if use_process_pool else ThreadPoolExecutor
        self._executor = cls(max_workers=max_workers)

    def submit(self, fn: Callable[[dict[str, Any]], dict[str, Any]], payload: dict[str, Any]) -> Future[dict[str, Any]]:
        return self._executor.submit(fn, payload)

    def shutdown(self) -> None:
        self._executor.shutdown(wait=True)


class _RemoteSweepExecutorAdapter(_SweepExecutorAdapter):
    """Threaded adapter that simulates endpoint-aware dispatch for remote workers."""

    def __init__(self, *, max_workers: int, endpoints: tuple[str, ...]) -> None:
        self._executor = ThreadPoolExecutor(max_workers=max_workers)
        self._endpoints = endpoints or ("remote-default",)

    def submit(self, fn: Callable[[dict[str, Any]], dict[str, Any]], payload: dict[str, Any]) -> Future[dict[str, Any]]:
        endpoint = self._endpoints[int(payload["combo_index"]) % len(self._endpoints)]
        remote_payload = dict(payload)
        remote_payload["worker_endpoint"] = endpoint
        return self._executor.submit(fn, remote_payload)

    def shutdown(self) -> None:
        self._executor.shutdown(wait=True)


def _load_resume_state(path: Path) -> dict[str, Any] | None:
    if not path.exists():
        return None
    return json.loads(path.read_text())


def _persist_resume_state(
    *,
    state_path: Path,
    job_id: str,
    seed: int,
    queued_indices: list[int],
    completed_rows: list[dict[str, Any]],
    invalid_rows: list[dict[str, Any]],
    errors: list[dict[str, Any]],
    retry_counts: dict[int, int],
) -> None:
    state_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "job_id": job_id,
        "seed": int(seed),
        "updated_at": datetime.now().isoformat(),
        "queued_indices": list(queued_indices),
        "completed_rows": list(completed_rows),
        "invalid_rows": list(invalid_rows),
        "errors": list(errors),
        "retry_counts": {str(k): int(v) for k, v in retry_counts.items()},
    }
    state_path.write_text(json.dumps(payload, indent=2))


def _persist_partial_leaderboard(path: Path, rows: list[dict[str, Any]]) -> None:
    ranked = sorted(rows, key=lambda row: float(row.get("sharpe", 0.0)), reverse=True)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(ranked, indent=2))


def _resolve_executor_adapter(
    *,
    worker_config: SweepWorkerConfig,
    use_process_pool: bool,
) -> _SweepExecutorAdapter:
    if worker_config.mode == "remote":
        return _RemoteSweepExecutorAdapter(
            max_workers=max(1, int(worker_config.max_workers)),
            endpoints=worker_config.remote_endpoints,
        )
    return _LocalSweepExecutorAdapter(
        max_workers=max(1, int(worker_config.max_workers)),
        use_process_pool=use_process_pool,
    )


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
    resume_state_path: Path | None = None,
    worker_config: SweepWorkerConfig | None = None,
    retry_policy: SweepRetryPolicy | None = None,
    lineage_parent_manifest: str | None = None,
    preflight_config: PreflightValidationConfig | None = None,
    benchmarks: list[str] | None = None,
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
    worker_cfg = worker_config or SweepWorkerConfig(mode="local", max_workers=n_workers)
    retry_cfg = retry_policy or SweepRetryPolicy()

    resolved_preflight = preflight_config or PreflightValidationConfig()
    base_payload = {
        "seed": int(seed),
        "tickers": tickers,
        "start_date": start_date.isoformat(),
        "end_date": end_date.isoformat(),
        "cache_root": str(cache_root),
        "preflight_max_missing_bars_ratio": float(resolved_preflight.max_missing_bars_ratio),
        "preflight_min_symbol_coverage_ratio": float(resolved_preflight.min_symbol_coverage_ratio),
        "preflight_critical_checks": ",".join(resolved_preflight.critical_checks),
        "preflight_block_on_critical": bool(resolved_preflight.block_on_critical),
    }
    queued_indices = list(range(len(combos)))
    rows: list[dict[str, Any]] = []
    errors: list[dict[str, Any]] = []
    retry_counts: dict[int, int] = {}
    job_id = _stable_fingerprint({
        "tickers": tickers,
        "start_date": start_date.isoformat(),
        "end_date": end_date.isoformat(),
        "entry_grid": entry_grid,
        "exit_grid": exit_grid,
        "core_grid": core_grid,
        "seed": seed,
    })
    state_path = resume_state_path or (BACKTEST_OUTPUT_DIR / f"sweep_job_state_{job_id}.json")
    state_path.parent.mkdir(parents=True, exist_ok=True)
    partial_results_path = state_path.with_suffix(".partial.json")

    prior_state = _load_resume_state(state_path)
    resumed_from_manifest: str | None = None
    if prior_state and str(prior_state.get("job_id")) == job_id:
        queued_indices = [int(i) for i in prior_state.get("queued_indices", [])]
        rows = list(prior_state.get("completed_rows", []))
        invalid_rows = list(prior_state.get("invalid_rows", invalid_rows))
        errors = list(prior_state.get("errors", []))
        retry_counts = {int(k): int(v) for k, v in dict(prior_state.get("retry_counts", {})).items()}
        resumed_from_manifest = str(state_path)

    LOGGER.info(
        "Starting sweep job=%s with %s queued combinations (%s workers, mode=%s)",
        job_id,
        len(queued_indices),
        worker_cfg.max_workers,
        worker_cfg.mode,
    )

    adapter = _resolve_executor_adapter(worker_config=worker_cfg, use_process_pool=use_process_pool)
    in_flight: dict[Future[dict[str, Any]], tuple[int, float]] = {}
    completed = len(rows)
    try:
        while queued_indices or in_flight:
            while queued_indices and len(in_flight) < max(1, int(worker_cfg.max_workers)):
                idx = queued_indices.pop(0)
                payload = {
                    "combo_index": idx,
                    "worker_seed": int(seed) + int(idx),
                    "job_id": job_id,
                    **base_payload,
                    **combos[idx],
                }
                fut = adapter.submit(worker, payload)
                in_flight[fut] = (idx, time.monotonic())

            if not in_flight:
                continue

            done, _ = wait(list(in_flight.keys()), timeout=0.25, return_when=FIRST_COMPLETED)
            now = time.monotonic()
            stale = [f for f, (_, started) in in_flight.items() if (now - started) > retry_cfg.stale_worker_timeout_seconds]
            for future in stale:
                idx, _ = in_flight.pop(future)
                attempts = retry_counts.get(idx, 0) + 1
                retry_counts[idx] = attempts
                errors.append({"error": "stale_worker_timeout", "combo": combos[idx], "attempt": attempts})
                if attempts < retry_cfg.max_attempts:
                    queued_indices.append(idx)

            for future in done:
                idx, _ = in_flight.pop(future)
                completed += 1
                try:
                    rows.append(future.result())
                except Exception as exc:
                    attempts = retry_counts.get(idx, 0) + 1
                    retry_counts[idx] = attempts
                    errors.append({"error": str(exc), "combo": combos[idx], "attempt": attempts})
                    LOGGER.exception("Sweep combo failed (%s/%s)", completed, len(combos))
                    if attempts < retry_cfg.max_attempts:
                        queued_indices.append(idx)
                    elif fail_fast:
                        raise
                    elif not continue_on_error:
                        raise
                LOGGER.info("Sweep progress: %s/%s", completed, len(combos))

            _persist_partial_leaderboard(partial_results_path, rows)
            _persist_resume_state(
                state_path=state_path,
                job_id=job_id,
                seed=seed,
                queued_indices=queued_indices,
                completed_rows=rows,
                invalid_rows=invalid_rows,
                errors=errors,
                retry_counts=retry_counts,
            )
    finally:
        adapter.shutdown()

    ranked_rows = sorted(rows, key=lambda row: (bool(row.get("stress_passed", False)), float(row["sharpe"])), reverse=True)
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
            "worker_mode": worker_cfg.mode,
            "remote_endpoints": list(worker_cfg.remote_endpoints),
            "retry_policy": {
                "max_attempts": retry_cfg.max_attempts,
                "stale_worker_timeout_seconds": retry_cfg.stale_worker_timeout_seconds,
            },
            "fail_fast": fail_fast,
            "continue_on_error": continue_on_error,
            "top_n": top_n,
            "data_fingerprint": {},
        },
        random_seed=seed,
        governance=None,
        lineage={
            "job_id": job_id,
            "resume_state_path": str(state_path),
            "resumed_from": resumed_from_manifest,
            "lineage_parent_manifest": lineage_parent_manifest,
            "partial_results_path": str(partial_results_path),
            "retry_counts": retry_counts,
        },
    )
    if state_path.exists():
        state_path.unlink()
    return (
        f"Sweep complete: {len(ranked_rows)} successful combos, "
        f"{len(invalid_rows)} skipped, {len(errors)} failed. "
        f"Saved outputs to: {run_dir}"
    )


def _normalize_model_grid(model_grid: dict[str, list[dict[str, Any]]]) -> list[dict[str, Any]]:
    variants: list[dict[str, Any]] = []
    for model_name, param_rows in sorted(model_grid.items()):
        rows = list(param_rows or [{}])
        if not rows:
            rows = [{}]
        for params in rows:
            variants.append({"model_name": str(model_name), "model_params": dict(params or {})})
    return variants


def _persist_deduped_artifacts(*, rows: list[dict[str, Any]], artifact_store_dir: Path) -> None:
    artifact_store_dir.mkdir(parents=True, exist_ok=True)
    hash_to_ref: dict[str, str] = {}
    for row in rows:
        artifacts = row.pop("artifacts", None)
        if not isinstance(artifacts, dict) or not artifacts:
            continue
        refs: dict[str, str] = {}
        for name, payload in artifacts.items():
            blob = json.dumps(payload, sort_keys=True, default=str)
            digest = hashlib.sha256(blob.encode("utf-8")).hexdigest()
            ref = hash_to_ref.get(digest)
            if ref is None:
                target = artifact_store_dir / f"{digest}.json"
                if not target.exists():
                    target.write_text(blob)
                ref = str(target)
                hash_to_ref[digest] = ref
            refs[str(name)] = ref
        if refs:
            row["artifact_refs"] = refs


def get_experiment_grid_status(*, state_path: Path) -> dict[str, Any]:
    state = _load_resume_state(state_path)
    if state is None:
        return {
            "state_path": str(state_path),
            "status": "missing",
            "queued": 0,
            "completed": 0,
            "errors": 0,
            "retry_counts": {},
        }
    queued = len(list(state.get("queued_indices", [])))
    completed = len(list(state.get("completed_rows", [])))
    errors = len(list(state.get("errors", [])))
    return {
        "state_path": str(state_path),
        "status": "running" if queued else "finalizing",
        "job_id": str(state.get("job_id", "")),
        "queued": queued,
        "completed": completed,
        "errors": errors,
        "retry_counts": dict(state.get("retry_counts", {})),
        "updated_at": str(state.get("updated_at", "")),
    }


def run_experiment_grid(
    *,
    tickers: list[str],
    start_date: date,
    end_date: date,
    cache_root: Path,
    entry_grid: dict[str, list[dict[str, Any]]],
    exit_grid: dict[str, list[dict[str, Any]]],
    core_grid: dict[str, list[Any]],
    model_grid: dict[str, list[dict[str, Any]]],
    seed: int = 42,
    max_workers: int | None = None,
    evaluator: Callable[[dict[str, Any]], dict[str, Any]] | None = None,
    resume_state_path: Path | None = None,
    worker_config: SweepWorkerConfig | None = None,
    retry_policy: SweepRetryPolicy | None = None,
    fail_fast: bool = False,
    continue_on_error: bool = True,
) -> str:
    """Run model-comparison experiment grids with resumable orchestration metadata."""

    combos, invalid_rows = generate_sweep_combinations(entry_grid=entry_grid, exit_grid=exit_grid, core_grid=core_grid)
    if not combos:
        raise ValueError("No valid combinations generated for experiment grid")
    model_variants = _normalize_model_grid(model_grid)
    if not model_variants:
        raise ValueError("No model variants provided")

    all_tasks: list[dict[str, Any]] = []
    for model in model_variants:
        for combo in combos:
            all_tasks.append({**model, **combo})

    n_workers = max_workers or min(8, max(1, os.cpu_count() or 1))
    worker_cfg = worker_config or SweepWorkerConfig(mode="local", max_workers=n_workers)
    retry_cfg = retry_policy or SweepRetryPolicy(max_attempts=2, stale_worker_timeout_seconds=120.0)
    worker = evaluator or _execute_sweep_combo

    job_id = _stable_fingerprint(
        {
            "kind": "experiment_grid",
            "tickers": list(tickers),
            "start_date": start_date.isoformat(),
            "end_date": end_date.isoformat(),
            "entry_grid": entry_grid,
            "exit_grid": exit_grid,
            "core_grid": core_grid,
            "model_grid": model_grid,
            "seed": int(seed),
        }
    )
    state_path = resume_state_path or (BACKTEST_OUTPUT_DIR / f"experiment_grid_state_{job_id}.json")
    partial_results_path = state_path.with_suffix(".partial.json")
    prior_state = _load_resume_state(state_path)

    queued_indices = list(range(len(all_tasks)))
    rows: list[dict[str, Any]] = []
    errors: list[dict[str, Any]] = []
    retry_counts: dict[int, int] = {}
    resumed_from_manifest: str | None = None
    if prior_state and str(prior_state.get("job_id")) == job_id:
        queued_indices = [int(i) for i in prior_state.get("queued_indices", [])]
        rows = list(prior_state.get("completed_rows", []))
        errors = list(prior_state.get("errors", []))
        retry_counts = {int(k): int(v) for k, v in dict(prior_state.get("retry_counts", {})).items()}
        resumed_from_manifest = str(state_path)

    adapter = _resolve_executor_adapter(worker_config=worker_cfg, use_process_pool=False)
    in_flight: dict[Future[dict[str, Any]], tuple[int, float]] = {}
    total_task_runtime = 0.0
    started_at = time.monotonic()
    try:
        while queued_indices or in_flight:
            while queued_indices and len(in_flight) < max(1, int(worker_cfg.max_workers)):
                idx = queued_indices.pop(0)
                payload = {
                    "combo_index": idx,
                    "worker_seed": int(seed) + int(idx),
                    "job_id": job_id,
                    "tickers": list(tickers),
                    "start_date": start_date,
                    "end_date": end_date,
                    "cache_root": cache_root,
                    **all_tasks[idx],
                }
                fut = adapter.submit(worker, payload)
                in_flight[fut] = (idx, time.monotonic())

            done, _ = wait(list(in_flight.keys()), timeout=0.25, return_when=FIRST_COMPLETED)
            now = time.monotonic()
            stale = [f for f, (_, started) in in_flight.items() if (now - started) > retry_cfg.stale_worker_timeout_seconds]
            for future in stale:
                idx, _ = in_flight.pop(future)
                attempts = retry_counts.get(idx, 0) + 1
                retry_counts[idx] = attempts
                errors.append({"error": "stale_worker_timeout", "task": all_tasks[idx], "attempt": attempts})
                if attempts < retry_cfg.max_attempts:
                    queued_indices.append(idx)

            for future in done:
                idx, task_started = in_flight.pop(future)
                elapsed = max(0.0, time.monotonic() - task_started)
                total_task_runtime += elapsed
                try:
                    row = future.result()
                    row["model_name"] = str(all_tasks[idx]["model_name"])
                    row["model_params"] = json.dumps(all_tasks[idx]["model_params"], sort_keys=True)
                    row["task_runtime_seconds"] = elapsed
                    rows.append(row)
                except Exception as exc:
                    attempts = retry_counts.get(idx, 0) + 1
                    retry_counts[idx] = attempts
                    errors.append({"error": str(exc), "task": all_tasks[idx], "attempt": attempts})
                    if attempts < retry_cfg.max_attempts:
                        queued_indices.append(idx)
                    elif fail_fast or not continue_on_error:
                        raise

            _persist_partial_leaderboard(partial_results_path, rows)
            _persist_resume_state(
                state_path=state_path,
                job_id=job_id,
                seed=seed,
                queued_indices=queued_indices,
                completed_rows=rows,
                invalid_rows=invalid_rows,
                errors=errors,
                retry_counts=retry_counts,
            )
    finally:
        adapter.shutdown()

    elapsed_parallel = max(1e-9, time.monotonic() - started_at)
    _persist_deduped_artifacts(rows=rows, artifact_store_dir=BACKTEST_OUTPUT_DIR / "artifact_store")
    ranked_rows = sorted(rows, key=lambda row: float(row.get("sharpe", 0.0)), reverse=True)
    run_dir = _persist_sweep_outputs(
        ranked_rows=ranked_rows,
        invalid_rows=invalid_rows,
        errors=errors,
        top_n=10,
        parameters={
            "tickers": list(tickers),
            "start_date": start_date.isoformat(),
            "end_date": end_date.isoformat(),
            "cache_root": str(cache_root),
            "entry_grid": entry_grid,
            "exit_grid": exit_grid,
            "core_grid": core_grid,
            "model_grid": model_grid,
            "max_workers": int(worker_cfg.max_workers),
        },
        random_seed=seed,
        governance=None,
        lineage={
            "job_id": job_id,
            "resume_state_path": str(state_path),
            "resumed_from": resumed_from_manifest,
            "partial_results_path": str(partial_results_path),
            "retry_counts": retry_counts,
            "speedup_metrics": {
                "parallel_elapsed_seconds": elapsed_parallel,
                "single_node_baseline_seconds": total_task_runtime,
                "speedup_vs_single_node": total_task_runtime / elapsed_parallel,
            },
        },
    )
    if state_path.exists():
        state_path.unlink()
    return f"Experiment grid complete: {len(ranked_rows)} tasks finished. Saved outputs to: {run_dir}"


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
    cv_scheme: str = "walk_forward",
    cpcv_n_groups: int = 6,
    cpcv_n_test_groups: int = 2,
    cv_seed: int = 42,
    split_policy: str = "calendar-based",
    governance_metadata: dict[str, Any] | None = None,
    lineage_parent_manifest: str | None = None,
    stress_controls: dict[str, Any] | None = None,
    benchmarks: list[str] | None = None,
    objective_weights: dict[str, float] | None = None,
    overfitting_penalty: dict[str, float] | None = None,
    strategy_key: str | None = None,
    prior_strategy_keys: list[str] | None = None,
    history_path: Path | None = None,
    cancellation_token: CancellationToken | None = None,
    run_namespace: str | None = None,
) -> str:
    cancellation = cancellation_token or CancellationToken()
    cancellation.checkpoint("run_walk_forward_backtest:start")
    random.seed(int(cv_seed))
    np.random.seed(int(cv_seed))
    combos, invalid_rows = generate_sweep_combinations(
        entry_grid=entry_grid,
        exit_grid=exit_grid,
        core_grid=core_grid,
    )
    if not combos:
        raise ValueError("No valid combinations generated for walk-forward")

    governance_payload = _build_governance_metadata(governance_metadata)

    max_lookback = 1
    for combo_idx, combo in enumerate(combos):
        if combo_idx % 10 == 0:
            cancellation.checkpoint("run_walk_forward_backtest:combo_validation")
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

    cancellation.checkpoint("run_walk_forward_backtest:before_data_load")
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

    cancellation.checkpoint("run_walk_forward_backtest:before_fold_generation")
    if cv_scheme == "cpcv":
        folds = build_cpcv_walk_forward_folds(
            total_bars=total_bars,
            n_groups=int(cpcv_n_groups),
            n_test_groups=int(cpcv_n_test_groups),
            purge_window_bars=int(purge_window_bars),
            embargo_window_bars=int(embargo_window_bars),
            label_horizon_bars=int(label_horizon_bars),
        )
    else:
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

    supported_split_policies = {
        "calendar-based",
        "volatility-regime-stratified",
        "event-exclusion windows",
    }
    normalized_split_policy = str(split_policy).strip().lower()
    if normalized_split_policy not in supported_split_policies:
        raise ValueError(
            "split_policy must be one of calendar-based, volatility-regime-stratified, event-exclusion windows"
        )

    if cv_scheme == "cpcv":
        split_rows = [
            split.to_dict()
            for split in generate_combinatorial_purged_cv_splits(
                n_samples=total_bars,
                n_groups=int(cpcv_n_groups),
                n_test_groups=int(cpcv_n_test_groups),
                purge_window_bars=int(purge_window_bars),
                embargo_window_bars=int(embargo_window_bars),
                label_horizon_bars=int(label_horizon_bars),
                seed=int(cv_seed),
            )
        ]
    else:
        split_rows = [
            split.to_dict()
            for split in generate_purged_kfold_splits(
                n_samples=total_bars,
                n_splits=max(2, len(folds)),
                purge_window_bars=int(purge_window_bars),
                embargo_window_bars=int(embargo_window_bars),
                label_horizon_bars=int(label_horizon_bars),
            )
        ]

    _apply_split_policy_metadata(
        split_rows=split_rows,
        split_policy=normalized_split_policy,
        date_index=arrays.date_index,
        prices=prices,
        stress_controls=stress_controls,
    )

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

    cancellation.checkpoint("run_walk_forward_backtest:before_walk_forward_optimization")
    wf_result = run_walk_forward_optimization(
        folds=folds,
        parameter_candidates=combos,
        evaluate_segment=evaluate_segment,
        score_metric=score_metric,
        nested_inner_folds=nested_inner_folds,
    )

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    namespace_prefix = f"{run_namespace}_" if run_namespace else ""
    run_dir = BACKTEST_OUTPUT_DIR / f"{namespace_prefix}tsmom_walk_forward_{timestamp}"
    persist_walk_forward_outputs(run_dir=run_dir, result=wf_result)
    (run_dir / "skipped_invalid_combos.json").write_text(json.dumps(invalid_rows, indent=2))
    (run_dir / "audit_inputs.json").write_text(
        json.dumps(
            {
                "schema_version": "1.0",
                "tickers": tickers,
                "start_date": start_date.isoformat(),
                "end_date": end_date.isoformat(),
                "entry_grid": entry_grid,
                "exit_grid": exit_grid,
                "core_grid": core_grid,
                "score_metric": score_metric,
                "cv_scheme": cv_scheme,
                "cpcv_n_groups": int(cpcv_n_groups),
                "cpcv_n_test_groups": int(cpcv_n_test_groups),
                "cv_seed": int(cv_seed),
                "split_policy": normalized_split_policy,
                "random_seeds": {"run_seed": int(cv_seed), "python_random_seed": int(cv_seed), "numpy_random_seed": int(cv_seed)},
                "windowing": {
                    "train_bars": train_bars,
                    "validation_bars": validation_bars,
                    "test_bars": test_bars,
                    "step_bars": step_bars,
                    "train_fraction": train_fraction,
                    "validation_fraction": validation_fraction,
                    "test_fraction": test_fraction,
                    "step_fraction": step_fraction,
                },
            },
            indent=2,
        )
    )
    (run_dir / "audit_outputs.json").write_text(
        json.dumps(
            {
                "aggregate_metrics": wf_result.aggregate_metrics,
                "stability": wf_result.stability,
                "validation_report": wf_result.validation_report,
                "fold_count": len(wf_result.folds),
            },
            indent=2,
        )
    )
    (run_dir / "split_metadata.json").write_text(json.dumps(split_rows, indent=2))
    with (run_dir / "fold_performance.csv").open("w", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=["fold_id", "total_return", "sharpe", "max_drawdown", "turnover_total", "hit_rate"],
        )
        writer.writeheader()
        for row in wf_result.folds:
            metrics = row.get("oos_metrics", {}) if isinstance(row.get("oos_metrics"), dict) else {}
            writer.writerow(
                {
                    "fold_id": int(row.get("fold_id", -1)),
                    "total_return": float(metrics.get("total_return", 0.0)),
                    "sharpe": float(metrics.get("sharpe", 0.0)),
                    "max_drawdown": float(metrics.get("max_drawdown", 0.0)),
                    "turnover_total": float(metrics.get("turnover_total", metrics.get("turnover", 0.0))),
                    "hit_rate": float(metrics.get("hit_rate", 0.0)),
                }
            )
    with (run_dir / "fold_boundaries.csv").open("w", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=["split_id", "scheme", "train_ranges", "test_ranges", "purge_window_bars", "embargo_window_bars"],
        )
        writer.writeheader()
        for row in split_rows:
            metadata = row.get("metadata", {}) if isinstance(row.get("metadata"), dict) else {}
            writer.writerow(
                {
                    "split_id": int(row.get("split_id", -1)),
                    "scheme": str(metadata.get("scheme", cv_scheme)),
                    "train_ranges": json.dumps(row.get("train_ranges", [])),
                    "test_ranges": json.dumps(row.get("test_ranges", [])),
                    "purge_window_bars": int(row.get("purge_window_bars", 0)),
                    "embargo_window_bars": int(row.get("embargo_window_bars", 0)),
                }
            )
    report_text = _format_walk_forward_report(
        folds=wf_result.folds,
        aggregate_metrics=wf_result.aggregate_metrics,
        stability=wf_result.stability,
        score_metric=score_metric,
        candidate_count=len(combos),
        skipped_invalid_count=len(invalid_rows),
    )
    (run_dir / "report.txt").write_text(report_text)
    (run_dir / "artifact_metadata.json").write_text(json.dumps({
        "schema_version": "1.0",
        "run_type": "walk_forward",
        "random_seeds": {"run_seed": int(cv_seed), "python_random_seed": int(cv_seed), "numpy_random_seed": int(cv_seed)},
        "data_fingerprint": {},
    }, indent=2))

    computed_checks = _evaluate_governance_gate_checks(
        metrics={k: float(v) for k, v in wf_result.aggregate_metrics.items()},
        fold_rows=wf_result.folds,
        governance=governance_payload,
    )
    governance_payload["gate_checks"].update(computed_checks)
    required_checks = governance_payload.get("promotion_required_checks", [])
    governance_payload["missing_required_checks"] = [
        name for name in required_checks if not governance_payload["gate_checks"].get(name, False)
    ]
    governance_payload["is_promotion_ready"] = not governance_payload["missing_required_checks"]
    governance_payload["audit_trail"].append({
        "timestamp": datetime.now().isoformat(),
        "event": "gate_checks_evaluated",
        "gate_checks": dict(governance_payload["gate_checks"]),
        "missing_required_checks": list(governance_payload["missing_required_checks"]),
    })

    cancellation.checkpoint("run_walk_forward_backtest:before_manifest")
    manifest_parameters = {
        "tickers": list(tickers),
        "start_date": start_date.isoformat(),
        "end_date": end_date.isoformat(),
        "entry_grid": entry_grid,
        "exit_grid": exit_grid,
        "core_grid": core_grid,
        "score_metric": score_metric,
        "nested_optimization": nested_optimization,
        "cv_scheme": cv_scheme,
        "cpcv_n_groups": int(cpcv_n_groups),
        "cpcv_n_test_groups": int(cpcv_n_test_groups),
        "cv_seed": int(cv_seed),
        "split_policy": normalized_split_policy,
        "stress_controls": dict(stress_controls or {}),
        "governance_metadata": governance_payload,
        "cancellation": cancellation.snapshot(),
    }
    manifest = _build_run_manifest(
        run_type="walk_forward",
        parameters=manifest_parameters,
        data_snapshot=_build_sweep_snapshot_identifiers(
            {
                "tickers": tickers,
                "start_date": start_date.isoformat(),
                "end_date": end_date.isoformat(),
                "cache_root": str(cache_root),
                "core_grid": core_grid,
                "data_fingerprint": {},
            }
        ),
        random_seed=int(cv_seed),
        governance=governance_payload,
        lineage={
            "lineage_parent_manifest": lineage_parent_manifest,
            "merged_leaderboard_sources": [str(run_dir / "fold_scores.csv"), str(run_dir / "fold_details.json")],
            "cancellation": cancellation.snapshot(),
        },
        extra_fingerprint_payload={"lineage_parent_manifest": lineage_parent_manifest},
    )
    (run_dir / "manifest.json").write_text(json.dumps(manifest, indent=2))
    _append_experiment_index(
        {
            "timestamp": manifest["created_at"],
            "run_type": "walk_forward",
            "run_dir": str(run_dir),
            "code_version": manifest["code_version"],
            "metric_schema_version": CANONICAL_METRIC_SCHEMA_VERSION,
            "random_seed": int(cv_seed),
            "primary_metric": score_metric,
            "primary_metric_value": float(wf_result.aggregate_metrics.get(score_metric, 0.0)),
            "manifest_path": str(run_dir / "manifest.json"),
            "reproducibility_fingerprint": manifest["reproducibility_fingerprint"],
            "run_id": manifest["run_id"],
            "config_hash": manifest["config_hash"],
            "config_checksum": manifest["config_checksum"],
            "data_snapshot_checksum": manifest["data_snapshot_checksum"],
            "manifest_checksum": manifest["manifest_checksum"],
            "parameters": manifest["parameters"],
            "data_snapshot_identifiers": manifest["data_snapshot_identifiers"],
            "metrics": {k: float(v) for k, v in wf_result.aggregate_metrics.items()},
            "significance": {},
            "governance": governance_payload,
            "model_artifacts": _collect_artifact_inventory(run_dir).get("model_artifacts", []),
            "plot_artifacts": _collect_artifact_inventory(run_dir).get("plot_artifacts", []),
            "metric_artifacts": _collect_artifact_inventory(run_dir).get("metric_artifacts", []),
            "reproducibility_metadata": manifest.get("reproducibility_metadata", {}),
        }
    )

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
    search_space: dict[str, Any] | None = None,
    objectives: list[dict[str, str]] | None = None,
    max_turnover: float | None = None,
    max_drawdown_floor: float | None = None,
    min_trades: float | None = None,
    partial_period_fractions: list[float] | None = None,
    governance_metadata: dict[str, Any] | None = None,
    stress_controls: dict[str, Any] | None = None,
    benchmarks: list[str] | None = None,
    objective_weights: dict[str, float] | None = None,
    overfitting_penalty: dict[str, float] | None = None,
    strategy_key: str | None = None,
    prior_strategy_keys: list[str] | None = None,
    history_path: Path | None = None,
    enable_pruning: bool = True,
    prune_on_constraint_violation: bool = True,
    prune_on_lcb: bool = True,
    min_completed_for_pruning: int = 5,
    staged_budgets: list[dict[str, Any]] | None = None,
    cancellation_token: CancellationToken | None = None,
    run_namespace: str | None = None,
) -> str:
    cancellation = cancellation_token or CancellationToken()
    cancellation.checkpoint("run_strategy_optimization:start")
    random.seed(int(seed))
    np.random.seed(int(seed))
    combos, invalid_rows = generate_sweep_combinations(
        entry_grid=entry_grid,
        exit_grid=exit_grid,
        core_grid=core_grid,
    )
    if not combos:
        raise ValueError("No valid combinations generated for optimization")

    governance_payload = _build_governance_metadata(governance_metadata)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    namespace_prefix = f"{run_namespace}_" if run_namespace else ""
    run_dir = BACKTEST_OUTPUT_DIR / f"{namespace_prefix}tsmom_optimize_{timestamp}"
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

    def _build_sampler(name: str):
        sampler_key = str(name).strip().lower()
        if sampler_key == "bayesian":
            return BayesianSampler()
        if sampler_key == "tpe":
            return TPESampler()
        if sampler_key in {"cma-es", "cma_es", "cma"}:
            return CMASampler()
        if sampler_key == "grid":
            return GridSampler()
        return RandomSampler()

    def _evaluate(params: dict[str, Any], period_fraction: float) -> dict[str, float]:
        cancellation.checkpoint("run_strategy_optimization:evaluate_trial")
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
                "stress_controls": dict(stress_controls or {}),
                **combo,
            }
        )
        return {key: float(value) for key, value in row.items() if isinstance(value, (int, float))}

    effective_space = search_space or {"combo_index": {"type": "discrete", "values": list(range(len(combos)))}}

    optimization_history_path = history_path if history_path is not None else run_dir / "optimization_history.jsonl"
    stages = list(staged_budgets or [])
    if not stages:
        stages = [{
            "label": "single_stage",
            "n_trials": int(n_trials),
            "sampler": sampler_name,
            "partial_period_fractions": partial_period_fractions,
        }]

    stage_summaries: list[dict[str, Any]] = []
    result: dict[str, Any] = {}
    for stage_idx, stage in enumerate(stages):
        cancellation.checkpoint(f"run_strategy_optimization:stage_{stage_idx}")
        stage_label = str(stage.get("label", f"stage_{stage_idx + 1}"))
        stage_trials = int(stage.get("n_trials", n_trials))
        if stage_trials <= 0:
            continue
        stage_sampler_name = str(stage.get("sampler", sampler_name))
        stage_sampler = _build_sampler(stage_sampler_name)
        stage_space = stage.get("search_space") if isinstance(stage.get("search_space"), dict) else effective_space
        stage_fractions = stage.get("partial_period_fractions")
        if isinstance(stage_fractions, list):
            parsed_stage_fractions = [float(v) for v in stage_fractions]
        else:
            parsed_stage_fractions = partial_period_fractions

        cancellation.checkpoint(f"run_strategy_optimization:before_optimize_stage_{stage_idx}")
        result = optimize(
            space=stage_space,
            evaluate=_evaluate,
            objectives=objective_specs,
            constraints=constraints,
            sampler=stage_sampler,
            n_trials=stage_trials,
            seed=int(seed) + stage_idx,
            partial_period_fractions=parsed_stage_fractions,
            output_dir=run_dir,
            objective_weights=objective_weights,
            overfitting_penalty=OverfittingPenaltyConfig(**(overfitting_penalty or {})),
            use_walk_forward_objective_metrics=True,
            history_path=optimization_history_path,
            strategy_key=strategy_key,
            prior_strategy_keys=prior_strategy_keys,
            enable_pruning=bool(stage.get("enable_pruning", enable_pruning)),
            prune_on_constraint_violation=bool(stage.get("prune_on_constraint_violation", prune_on_constraint_violation)),
            prune_on_lcb=bool(stage.get("prune_on_lcb", prune_on_lcb)),
            min_completed_for_pruning=int(stage.get("min_completed_for_pruning", min_completed_for_pruning)),
        )
        stage_summaries.append({
            "label": stage_label,
            "n_trials": stage_trials,
            "sampler": stage_sampler_name,
            "partial_period_fractions": parsed_stage_fractions,
            "search_space": stage_space,
            "trial_count": int(result.get("trial_count", 0)),
            "feasible_count": int(result.get("feasible_count", 0)),
            "pareto_count": int(result.get("pareto_count", 0)),
            "best_scalar": float(result.get("best_scalar", float("-inf"))),
        })

    best_metrics: dict[str, float] = {}
    pareto_trials = result.get("pareto_trials", [])
    if isinstance(pareto_trials, list) and pareto_trials:
        first = pareto_trials[0]
        if isinstance(first, dict) and isinstance(first.get("metrics"), dict):
            best_metrics = {k: float(v) for k, v in first.get("metrics", {}).items() if isinstance(v, (int, float))}
    computed_checks = _evaluate_governance_gate_checks(
        metrics=best_metrics,
        fold_rows=None,
        governance=governance_payload,
    )
    governance_payload["gate_checks"].update(computed_checks)
    required_checks = governance_payload.get("promotion_required_checks", [])
    governance_payload["missing_required_checks"] = [
        name for name in required_checks if not governance_payload["gate_checks"].get(name, False)
    ]
    governance_payload["is_promotion_ready"] = not governance_payload["missing_required_checks"]
    governance_payload["audit_trail"].append({
        "timestamp": datetime.now().isoformat(),
        "event": "gate_checks_evaluated",
        "gate_checks": dict(governance_payload["gate_checks"]),
        "missing_required_checks": list(governance_payload["missing_required_checks"]),
    })

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
                "random_seeds": {"run_seed": int(seed), "python_random_seed": int(seed), "numpy_random_seed": int(seed)},
                "n_trials": n_trials,
                "sampler": sampler_name,
                "staged_budgets": stages,
                "stage_summaries": stage_summaries,
                "pruning": {"enabled": bool(enable_pruning), "prune_on_constraint_violation": bool(prune_on_constraint_violation), "prune_on_lcb": bool(prune_on_lcb), "min_completed_for_pruning": int(min_completed_for_pruning)},
                "search_space": effective_space,
                "objectives": objective_defs,
                "constraints": [constraint.__dict__ for constraint in constraints],
                "partial_period_fractions": partial_period_fractions,
                "governance": governance_payload,
                "stress_controls": dict(stress_controls or {}),
                "objective_weights": dict(objective_weights or {}),
                "overfitting_penalty": dict(overfitting_penalty or {}),
                "strategy_key": strategy_key,
                "cancellation": cancellation.snapshot(),
                "prior_strategy_keys": list(prior_strategy_keys or []),
                "history_path": str(optimization_history_path),
            },
            indent=2,
            sort_keys=True,
        ),
        encoding="utf-8",
    )

    manifest = _build_run_manifest(
        run_type="optimization",
        parameters={
            "tickers": tickers,
            "start_date": start_date.isoformat(),
            "end_date": end_date.isoformat(),
            "entry_grid": entry_grid,
            "exit_grid": exit_grid,
            "core_grid": core_grid,
            "seed": seed,
            "n_trials": n_trials,
            "sampler": sampler_name,
                "staged_budgets": stages,
                "stage_summaries": stage_summaries,
                "pruning": {"enabled": bool(enable_pruning), "prune_on_constraint_violation": bool(prune_on_constraint_violation), "prune_on_lcb": bool(prune_on_lcb), "min_completed_for_pruning": int(min_completed_for_pruning)},
            "search_space": effective_space,
            "objectives": objective_defs,
            "constraints": [constraint.__dict__ for constraint in constraints],
            "partial_period_fractions": partial_period_fractions,
            "stress_controls": dict(stress_controls or {}),
            "governance_metadata": governance_payload,
            "cancellation": cancellation.snapshot(),
        },
        data_snapshot=_build_sweep_snapshot_identifiers({
            "tickers": tickers,
            "start_date": start_date.isoformat(),
            "end_date": end_date.isoformat(),
            "cache_root": str(cache_root),
            "core_grid": core_grid,
            "data_fingerprint": {},
        }),
        random_seed=int(seed),
        governance=governance_payload,
        extra_fingerprint_payload={"best_trials": result.get("pareto_trials", []), "cancellation": cancellation.snapshot()},
    )
    (run_dir / "manifest.json").write_text(json.dumps(manifest, indent=2))
    (run_dir / "artifact_metadata.json").write_text(json.dumps({
        "schema_version": "1.0",
        "run_type": "optimization",
        "random_seeds": {"run_seed": int(seed), "python_random_seed": int(seed), "numpy_random_seed": int(seed)},
        "data_fingerprint": {},
    }, indent=2))
    _append_experiment_index({
        "timestamp": manifest["created_at"],
        "run_type": "optimization",
        "run_dir": str(run_dir),
        "code_version": manifest["code_version"],
        "metric_schema_version": CANONICAL_METRIC_SCHEMA_VERSION,
        "random_seed": int(seed),
        "primary_metric": "pareto_count",
        "primary_metric_value": float(result.get("pareto_count", 0.0)),
        "manifest_path": str(run_dir / "manifest.json"),
        "reproducibility_fingerprint": manifest["reproducibility_fingerprint"],
        "run_id": manifest["run_id"],
        "config_hash": manifest["config_hash"],
        "config_checksum": manifest["config_checksum"],
        "data_snapshot_checksum": manifest["data_snapshot_checksum"],
        "manifest_checksum": manifest["manifest_checksum"],
        "parameters": manifest["parameters"],
        "data_snapshot_identifiers": manifest["data_snapshot_identifiers"],
        "metrics": best_metrics,
        "significance": {},
        "governance": governance_payload,
        "model_artifacts": _collect_artifact_inventory(run_dir).get("model_artifacts", []),
        "plot_artifacts": _collect_artifact_inventory(run_dir).get("plot_artifacts", []),
        "metric_artifacts": _collect_artifact_inventory(run_dir).get("metric_artifacts", []),
        "reproducibility_metadata": manifest.get("reproducibility_metadata", {}),
    })

    return (
        f"Optimization complete: {result['trial_count']} trials, {result['feasible_count']} feasible, "
        f"{result['pareto_count']} Pareto-optimal across {len(stage_summaries)} stage(s). "
        f"Top robust sets: {len(result.get('best_robust_params', []))}. Saved outputs to: {run_dir}"
    )


def _apply_split_policy_metadata(
    *,
    split_rows: list[dict[str, Any]],
    split_policy: str,
    date_index: Any,
    prices: np.ndarray,
    stress_controls: dict[str, Any] | None,
) -> None:
    event_windows = _resolve_event_exclusion_windows(date_index=date_index, stress_controls=stress_controls)
    regime_boundaries = _resolve_volatility_regime_boundaries(prices=prices)

    for row in split_rows:
        metadata = row.get("metadata")
        if not isinstance(metadata, dict):
            metadata = {}
            row["metadata"] = metadata
        metadata["split_policy"] = split_policy

        if split_policy == "volatility-regime-stratified":
            metadata["volatility_regime_boundaries"] = regime_boundaries
        elif split_policy == "event-exclusion windows":
            metadata["event_exclusion_windows"] = event_windows
            test_ranges = row.get("test_ranges", [])
            metadata["excluded_event_windows"] = _intersect_ranges(test_ranges, event_windows)


def _resolve_volatility_regime_boundaries(*, prices: np.ndarray) -> list[list[int]]:
    if prices.ndim != 2 or prices.shape[0] < 3:
        return []
    anchor = prices[:, 0]
    returns = np.diff(np.log(np.clip(anchor, 1e-12, None)))
    rolling = np.zeros(anchor.shape[0], dtype=float)
    if returns.size > 0:
        for idx in range(anchor.shape[0]):
            left = max(0, idx - 20)
            right = max(0, idx - 1)
            if right > left:
                window = returns[left:right]
                rolling[idx] = float(np.std(window, ddof=0)) if window.size else 0.0
    valid = rolling[rolling > 0]
    if valid.size == 0:
        return []
    low = float(np.quantile(valid, 0.33))
    high = float(np.quantile(valid, 0.66))
    low_idx = np.where(rolling <= low)[0]
    mid_idx = np.where((rolling > low) & (rolling <= high))[0]
    high_idx = np.where(rolling > high)[0]
    return [
        _indices_to_span(low_idx),
        _indices_to_span(mid_idx),
        _indices_to_span(high_idx),
    ]


def _indices_to_span(indices: np.ndarray) -> list[int]:
    if indices.size == 0:
        return [0, 0]
    return [int(indices[0]), int(indices[-1]) + 1]


def _resolve_event_exclusion_windows(*, date_index: Any, stress_controls: dict[str, Any] | None) -> list[list[int]]:
    controls = stress_controls or {}
    configured = controls.get("event_exclusion_windows", [])
    windows: list[list[int]] = []
    if isinstance(configured, list):
        for item in configured:
            if not isinstance(item, (list, tuple)) or len(item) != 2:
                continue
            start = int(item[0])
            end = int(item[1])
            if end > start:
                windows.append([start, end])
    if windows:
        return windows

    # Fallback synthetic event windows based on calendar month boundaries.
    month_starts: list[int] = []
    prev_key: tuple[int, int] | None = None
    for idx, ts in enumerate(date_index):
        dt = _timestamp_to_datetime(ts)
        key = (dt.year, dt.month)
        if key != prev_key:
            month_starts.append(idx)
            prev_key = key
    for start in month_starts[:3]:
        windows.append([int(start), int(start + 1)])
    return windows


def _intersect_ranges(primary: list[Any], secondary: list[list[int]]) -> list[list[int]]:
    intersections: list[list[int]] = []
    for left in primary:
        if not isinstance(left, list) or len(left) != 2:
            continue
        a0, a1 = int(left[0]), int(left[1])
        for b0, b1 in secondary:
            start = max(a0, int(b0))
            end = min(a1, int(b1))
            if end > start:
                intersections.append([start, end])
    return intersections


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
    fold_reuse = stability.get("fold_reuse", {})
    if isinstance(fold_reuse, dict) and fold_reuse:
        lines.append(f"- fold_reuse_train_avg: {float(fold_reuse.get('train_avg_reuse', 0.0)):.4f}")
        lines.append(f"- fold_reuse_validation_avg: {float(fold_reuse.get('validation_avg_reuse', 0.0)):.4f}")
        lines.append(f"- fold_reuse_test_avg: {float(fold_reuse.get('test_avg_reuse', 0.0)):.4f}")

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
    combo_seed = int(payload.get("worker_seed", int(payload["seed"]) + int(payload["combo_index"])))
    random.seed(combo_seed)
    np.random.seed(combo_seed)

    tickers = [str(ticker) for ticker in payload["tickers"]]
    start_date = date.fromisoformat(str(payload["start_date"]))
    end_date = date.fromisoformat(str(payload["end_date"]))
    lookback_days = int(payload["lookback_days"])
    skip_days = int(payload["skip_days"])
    costs_bps = float(payload["costs_bps"])
    universe_filters = payload.get("universe_filters")
    if universe_filters is not None and not isinstance(universe_filters, dict):
        raise ValueError("universe_filters must be a mapping when provided")
    aum_scale = max(float(payload.get("aum_scale", payload.get("notional_scale", 1.0))), 1e-6)
    max_participation_rate = payload.get("max_participation_rate")

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
    _run_preflight_or_raise(
        arrays=arrays,
        requested_tickers=tickers,
        start_dt=datetime.combine(start_date, datetime.min.time()),
        end_dt=datetime.combine(end_date, datetime.max.time()),
        timeframe="1m",
        config=PreflightValidationConfig(
            max_missing_bars_ratio=float(payload.get("preflight_max_missing_bars_ratio", 1.0)),
            min_symbol_coverage_ratio=float(payload.get("preflight_min_symbol_coverage_ratio", 1.0)),
            critical_checks=_parse_preflight_critical_checks(payload.get("preflight_critical_checks")),
            block_on_critical=bool(payload.get("preflight_block_on_critical", True)),
        ),
        workflow_label="sweep",
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
    friction_on = backtest_vectorized(
        prices=prices,
        signals=signals,
        slippage_model=BpsSlippage(costs_bps),
        fee_model=BrokerFeeModel(
            fee_bps=float(payload.get("broker_fee_bps", 0.0)),
            fee_per_unit=float(payload.get("broker_fee_per_unit", 0.0)),
            minimum_fee=float(payload.get("broker_min_fee", 0.0)),
        ),
        order_type=str(payload.get("order_type", "market")),
        available_bar_volume=np.maximum(prices, 1.0),
        max_participation_per_bar=float(payload.get("max_participation_per_bar", 1.0)),
        latency_bars=int(payload.get("latency_bars", 0)),
        latency_ms=int(payload.get("latency_ms", 0)),
        initial_equity=1.0,
        timeframe="1m",
    )
    friction_off = backtest_vectorized(
        prices=prices,
        signals=signals,
        slippage_model=BpsSlippage(0.0),
        fee_model=BrokerFeeModel(fee_bps=0.0, fee_per_unit=0.0, minimum_fee=0.0),
        borrow_cost_model=ShortBorrowCost(annual_borrow_rate=0.0),
        order_type=str(payload.get("order_type", "market")),
        available_bar_volume=np.maximum(prices, 1.0),
        max_participation_per_bar=float(payload.get("max_participation_per_bar", 1.0)),
        latency_bars=int(payload.get("latency_bars", 0)),
        latency_ms=int(payload.get("latency_ms", 0)),
        initial_equity=1.0,
        timeframe="1m",
    )
    result = friction_on

    turnover = _to_numpy_1d(result.turnover)
    trades = _to_numpy_2d(result.trades)
    fills = list(result.fills)
    cost_totals = {
        key: float(value)
        for key, value in result.cost_breakdown.get("totals", {}).items()
    }
    metrics = dict(result.metrics)
    friction_edge = float(friction_on.metrics.get("sharpe", 0.0))
    metrics["friction_adjusted_edge"] = friction_edge
    metrics["friction_edge_retention"] = 0.0 if abs(float(friction_off.metrics.get("sharpe", 0.0))) < 1e-12 else friction_edge / float(friction_off.metrics.get("sharpe", 0.0))
    turnover_total = float(np.sum(turnover)) if turnover.size else 0.0
    avg_participation = float(np.mean(np.array([float(row.get("participation_rate", 0.0)) for row in fills], dtype=float))) if fills else 0.0
    scaled_participation = avg_participation * aum_scale
    if max_participation_rate is not None and scaled_participation > float(max_participation_rate):
        raise ValueError(
            f"Capacity fail-fast triggered in sweep combo: scaled_participation={scaled_participation:.6f} exceeds max_participation_rate={float(max_participation_rate):.6f}"
        )
    impact_scale = float(np.sqrt(max(aum_scale, 1e-12)))
    scaled_slippage = float(cost_totals.get("slippage", 0.0)) * aum_scale * impact_scale
    non_slippage = max(float(cost_totals.get("total", 0.0)) - float(cost_totals.get("slippage", 0.0)), 0.0)
    scaled_cost_total = scaled_slippage + non_slippage * aum_scale
    vol = max(float(metrics.get("volatility", 0.0)), 1e-12)
    post_cost_sharpe = float(metrics.get("sharpe", 0.0)) - ((scaled_cost_total - float(cost_totals.get("total", 0.0))) / vol)
    scenario_defs = _build_stress_scenario_definitions(timestamps=np.arange(result.returns.size, dtype=np.int64), returns=_to_numpy_1d(result.returns), controls=dict(payload.get("stress_controls", {})), scenario_packs=list(payload.get("scenario_packs", [])))
    scenario_rows = _run_stress_scenario_wrappers(returns=_to_numpy_1d(result.returns), prices=prices, scenario_definitions=scenario_defs)
    scenario_payload = build_scenario_attribution_and_guardrails(baseline_metrics={k: float(v) for k, v in metrics.items() if isinstance(v, (int, float))}, scenario_results=scenario_rows)
    stress_gate = _stress_gate_summary(scenario_payload, controls=dict(payload.get("stress_controls", {})))
    return {
        "entry_signal": payload["entry_signal"],
        "entry_signal_params": json.dumps(payload.get("entry_signal_params", {}), sort_keys=True),
        "exit_signal": payload["exit_signal"],
        "exit_signal_params": json.dumps(payload.get("exit_signal_params", {}), sort_keys=True),
        "lookback_days": lookback_days,
        "skip_days": skip_days,
        "costs_bps": costs_bps,
        "aum_scale": aum_scale,
        "notional_scale": aum_scale,
        "total_return": float(metrics.get("total_return", 0.0)),
        "sharpe": float(metrics.get("sharpe", 0.0)),
        "post_cost_sharpe": float(post_cost_sharpe),
        "friction_off_sharpe": float(friction_off.metrics.get("sharpe", 0.0)),
        "friction_adjusted_edge": float(metrics.get("friction_adjusted_edge", 0.0)),
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
        "turnover_total": turnover_total,
        "turnover_total_scaled": float(turnover_total * aum_scale),
        "average_participation_rate": avg_participation,
        "scaled_participation_rate": scaled_participation,
        "slippage_total_scaled": scaled_slippage,
        "trade_count": float(np.sum(np.abs(trades) > 1e-12)) if trades.size else 0.0,
        "cost_total": cost_totals.get("total", 0.0),
        "cost_total_scaled": scaled_cost_total,
        **stress_gate,
    }


def _apply_sweep_statistical_annotations(
    *,
    ranked_rows: list[dict[str, Any]],
    robustness_report: dict[str, Any],
) -> None:
    if not ranked_rows:
        return

    multiple_testing = robustness_report.get("multiple_testing", {})
    raw_pvalues = multiple_testing.get("raw_pvalues", []) if isinstance(multiple_testing, dict) else []
    adjusted_pvalues = multiple_testing.get("bh_adjusted_pvalues", []) if isinstance(multiple_testing, dict) else []
    corrected_pvalues = multiple_testing.get("corrected_pvalues", []) if isinstance(multiple_testing, dict) else []
    correction_components = multiple_testing.get("correction_components", {}) if isinstance(multiple_testing, dict) else {}
    white_component = float(correction_components.get("white_reality_check_pvalue", 1.0)) if isinstance(correction_components, dict) else 1.0
    spa_component = float(correction_components.get("spa_pvalue", 1.0)) if isinstance(correction_components, dict) else 1.0

    sharpes = np.array([float(row.get("sharpe", 0.0)) for row in ranked_rows], dtype=float)
    centered = sharpes - float(np.mean(sharpes))
    m2 = float(np.mean(centered**2)) if sharpes.size > 1 else 0.0
    skew = float(np.mean(centered**3) / (m2 ** 1.5)) if sharpes.size > 2 and m2 > 0 else 0.0
    kurt = float(np.mean(centered**4) / (m2**2)) if sharpes.size > 3 and m2 > 0 else 3.0

    n_returns = max(3, int(sharpes.size))
    for idx, row in enumerate(ranked_rows):
        sharpe = float(row.get("sharpe", 0.0))
        row["deflated_sharpe_ratio"] = deflated_sharpe_ratio(
            observed_sharpe=sharpe,
            n_returns=n_returns,
            skew=skew,
            kurtosis=kurt,
            n_trials=max(1, len(ranked_rows)),
        )
        row["probabilistic_sharpe_ratio"] = probabilistic_sharpe_ratio(
            observed_sharpe=sharpe,
            benchmark_sharpe=0.0,
            n_returns=n_returns,
            skew=skew,
            kurtosis=kurt,
        )
        raw_p = float(raw_pvalues[idx]) if idx < len(raw_pvalues) else 1.0
        adj_p = float(adjusted_pvalues[idx]) if idx < len(adjusted_pvalues) else 1.0
        row["nominal_pvalue"] = raw_p
        row["bh_adjusted_pvalue"] = adj_p
        corrected_p = float(corrected_pvalues[idx]) if idx < len(corrected_pvalues) else max(adj_p, white_component, spa_component)
        row["corrected_pvalue"] = corrected_p
        row["white_reality_check_pvalue"] = white_component
        row["spa_pvalue"] = spa_component
        row["is_significant_nominal_5pct"] = raw_p <= 0.05
        row["is_significant_fdr_5pct"] = adj_p <= 0.05
        row["is_significant_corrected_5pct"] = corrected_p <= 0.05
        row["is_significant"] = bool(row["is_significant_corrected_5pct"] and float(row["deflated_sharpe_ratio"]) >= 0.95)



def _persist_sweep_outputs(
    *,
    ranked_rows: list[dict[str, Any]],
    invalid_rows: list[dict[str, Any]],
    errors: list[dict[str, Any]],
    top_n: int,
    parameters: dict[str, Any],
    random_seed: int,
    governance: dict[str, Any] | None = None,
    lineage: dict[str, Any] | None = None,
) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    run_dir = BACKTEST_OUTPUT_DIR / f"tsmom_sweep_{timestamp}"
    run_dir.mkdir(parents=True, exist_ok=True)

    governance_payload = _build_governance_metadata(governance)
    top_metrics = ranked_rows[0] if ranked_rows else {}
    computed_checks = _evaluate_governance_gate_checks(
        metrics={k: float(v) for k, v in top_metrics.items() if isinstance(v, (int, float))},
        fold_rows=None,
        governance=governance_payload,
    )
    governance_payload["gate_checks"].update(computed_checks)
    required_checks = governance_payload.get("promotion_required_checks", [])
    governance_payload["missing_required_checks"] = [
        name for name in required_checks if not governance_payload["gate_checks"].get(name, False)
    ]
    governance_payload["is_promotion_ready"] = not governance_payload["missing_required_checks"]
    governance_payload["audit_trail"].append({
        "timestamp": datetime.now().isoformat(),
        "event": "gate_checks_evaluated",
        "gate_checks": dict(governance_payload["gate_checks"]),
        "missing_required_checks": list(governance_payload["missing_required_checks"]),
    })

    robustness_report = build_sweep_robustness_report(ranked_rows=ranked_rows)
    _apply_sweep_statistical_annotations(ranked_rows=ranked_rows, robustness_report=robustness_report)

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
    (run_dir / "audit_inputs.json").write_text(json.dumps({"parameters": parameters, "random_seed": int(random_seed), "random_seeds": {"python_random_seed": int(random_seed), "numpy_random_seed": int(random_seed), "run_seed": int(random_seed)}}, indent=2))
    (run_dir / "audit_outputs.json").write_text(
        json.dumps(
            {
                "ranked_rows": ranked_rows,
                "invalid_rows": invalid_rows,
                "errors": errors,
            },
            indent=2,
        )
    )

    top_rows = ranked_rows[: max(0, top_n)]
    _write_robustness_report(run_dir=run_dir, robustness_report=robustness_report)

    report_lines = ["Top sweep combinations", "======================", ""]
    for idx, row in enumerate(top_rows, start=1):
        report_lines.append(
            f"#{idx}: sharpe={row['sharpe']:.6f} total_return={row['total_return']:.6f} "
            f"entry={row['entry_signal']} exit={row['exit_signal']} "
            f"core=(lookback_days={row['lookback_days']}, skip_days={row['skip_days']}, costs_bps={row['costs_bps']})"
        )
    pbo_style = robustness_report.get("pbo_style", {})
    white = robustness_report.get("white_reality_check", {})
    spa = robustness_report.get("spa", {})
    report_lines.extend([
        "",
        "Robustness Diagnostics",
        "----------------------",
        f"deflated_sharpe_ratio={float(robustness_report.get('deflated_sharpe_ratio', 0.0)):.6f}",
        f"probabilistic_sharpe_ratio={float(robustness_report.get('probabilistic_sharpe_ratio', 0.0)):.6f}",
        f"pbo_probability={float(pbo_style.get('probability_of_overfitting', 0.0)):.6f}",
        f"pbo_median_logit={float(pbo_style.get('median_logit', 0.0)):.6f}",
        f"white_reality_check_pvalue={float(white.get('p_value', 1.0)):.6f}",
        f"spa_pvalue={float(spa.get('p_value', 1.0)):.6f}",
        f"min_nominal_pvalue={float((robustness_report.get('multiple_testing', {}) or {}).get('min_raw_pvalue', 1.0)):.6f}",
        f"min_bh_adjusted_pvalue={float((robustness_report.get('multiple_testing', {}) or {}).get('min_bh_adjusted_pvalue', 1.0)):.6f}",
        f"min_corrected_pvalue={float((robustness_report.get('multiple_testing', {}) or {}).get('min_corrected_pvalue', 1.0)):.6f}",
    ])
    (run_dir / "top_n_report.txt").write_text("\n".join(report_lines))

    metric_tables = _write_metric_table_manifest(run_dir=run_dir, run_type="parameter_sweep", table_names=["leaderboard", "per_combo_summary"])
    (run_dir / "artifact_metadata.json").write_text(json.dumps({
        "schema_version": "1.0",
        "run_type": "parameter_sweep",
        "random_seeds": {"run_seed": int(random_seed), "python_random_seed": int(random_seed), "numpy_random_seed": int(random_seed)},
        "data_fingerprint": dict(parameters.get("data_fingerprint", {})) if isinstance(parameters, dict) else {},
    }, indent=2))
    manifest = _build_run_manifest(
        run_type="parameter_sweep",
        parameters=parameters,
        data_snapshot=_build_sweep_snapshot_identifiers(parameters),
        random_seed=int(random_seed),
        governance=governance_payload,
        lineage=lineage or {},
        result_summary={
            "successful_combos": len(ranked_rows),
            "invalid_combos": len(invalid_rows),
            "failed_combos": len(errors),
            "best_sharpe": float(ranked_rows[0]["sharpe"]) if ranked_rows else 0.0,
        },
        extra_fingerprint_payload={
            "top_row": ranked_rows[0] if ranked_rows else {},
        },
        metric_tables=metric_tables,
    )
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
            "run_id": manifest["run_id"],
            "config_hash": manifest["config_hash"],
            "config_checksum": manifest["config_checksum"],
            "data_snapshot_checksum": manifest["data_snapshot_checksum"],
            "manifest_checksum": manifest["manifest_checksum"],
            "parameters": parameters,
            "data_snapshot_identifiers": manifest["data_snapshot_identifiers"],
            "metrics": ranked_rows[0] if ranked_rows else {},
            "significance": {"robustness": robustness_report} if robustness_report else {},
            "governance": governance_payload,
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
        universe_filters = combo.get("universe_filters")
        if universe_filters is not None and not isinstance(universe_filters, dict):
            return False
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
    execution_model: str = "bps",
    execution_model_params: dict[str, object] | None = None,
    starting_capital: float = 100_000.0,
    bet_sizing_mode: str = "half_kelly",
    custom_bet_pct: float = 10.0,
    timeframe: str = "1m",
    portfolio_method: str = "equal_weight",
    portfolio_vol_lookback_bars: int = 20,
    portfolio_target_volatility: float = 0.10,
    portfolio_max_symbol_weight: float = 0.25,
    portfolio_max_sector_weight: float = 0.60,
    portfolio_rebalance_frequency_bars: int = 1,
    portfolio_clustering_linkage: str = "single",
    portfolio_covariance_shrinkage: float = 0.15,
    portfolio_max_gross_exposure: float = 1.0,
    portfolio_min_net_exposure: float = -1.0,
    portfolio_max_net_exposure: float = 1.0,
    portfolio_max_net_gamma: float | None = None,
    portfolio_max_abs_vega_bucket: float | None = None,
    portfolio_max_abs_delta_per_underlying: float | None = None,
    max_participation_rate: float | None = None,
    portfolio_sector_map: dict[str, str] | None = None,
    governance_metadata: dict[str, Any] | None = None,
    stress_controls: dict[str, Any] | None = None,
    scenario_packs: list[str] | None = None,
    benchmarks: list[str] | None = None,
    cancellation_token: CancellationToken | None = None,
    run_namespace: str | None = None,
) -> str:
    """Run all selected entry/exit combinations with shared core parameters."""

    cancellation = cancellation_token or CancellationToken()
    cancellation.checkpoint("run_multi_signal_backtest:start")

    if not entry_signals:
        raise ValueError("At least one entry signal must be selected")
    if not exit_signals:
        raise ValueError("At least one exit signal must be selected")

    rows: list[dict[str, Any]] = []
    combo_reports: list[str] = []

    selected_scenario_packs = [str(pack) for pack in (scenario_packs or [])]
    resolved_pack_templates = resolve_scenario_pack_templates(selected_scenario_packs)

    for entry_idx, entry_signal in enumerate(entry_signals):
        cancellation.checkpoint(f"run_multi_signal_backtest:entry_{entry_idx}")
        for exit_idx, exit_signal in enumerate(exit_signals):
            cancellation.checkpoint(f"run_multi_signal_backtest:entry_{entry_idx}_exit_{exit_idx}")
            report = run_time_series_momentum_backtest(
                tickers=tickers,
                start_date=start_date,
                end_date=end_date,
                cache_root=cache_root,
                lookback_days=lookback_days,
                skip_days=skip_days,
                costs_bps=costs_bps,
                execution_model=execution_model,
                execution_model_params=execution_model_params,
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
                portfolio_rebalance_frequency_bars=portfolio_rebalance_frequency_bars,
                portfolio_clustering_linkage=portfolio_clustering_linkage,
                portfolio_covariance_shrinkage=portfolio_covariance_shrinkage,
                portfolio_max_gross_exposure=portfolio_max_gross_exposure,
                portfolio_min_net_exposure=portfolio_min_net_exposure,
                portfolio_max_net_exposure=portfolio_max_net_exposure,
                portfolio_max_net_gamma=portfolio_max_net_gamma,
                portfolio_max_abs_vega_bucket=portfolio_max_abs_vega_bucket,
                portfolio_max_abs_delta_per_underlying=portfolio_max_abs_delta_per_underlying,
                max_participation_rate=max_participation_rate,
                portfolio_sector_map=portfolio_sector_map,
                governance_metadata=governance_metadata,
                stress_controls=stress_controls,
                scenario_packs=selected_scenario_packs,
                benchmarks=benchmarks,
            )
            run_dir = _extract_saved_output_dir(report)
            metrics = _load_metrics_from_run_dir(run_dir)
            stress_payload = _load_stress_payload_from_run_dir(run_dir)
            stress_gate = _stress_gate_summary(stress_payload, controls=stress_controls)
            benchmark_rows = _load_benchmark_rows_from_run_dir(run_dir)
            alpha_vs_bench = {str(row.get("benchmark", "")): float(row.get("alpha", 0.0)) for row in benchmark_rows}
            ir_vs_bench = {str(row.get("benchmark", "")): float(row.get("information_ratio", 0.0)) for row in benchmark_rows}
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
                    "alpha_vs_buy_hold": float(alpha_vs_bench.get(BENCHMARK_BUY_HOLD, 0.0)),
                    "alpha_vs_equal_weight_momentum": float(alpha_vs_bench.get(BENCHMARK_EQUAL_WEIGHT_MOMENTUM, 0.0)),
                    "alpha_vs_volatility_parity": float(alpha_vs_bench.get(BENCHMARK_VOLATILITY_PARITY, 0.0)),
                    "ir_vs_buy_hold": float(ir_vs_bench.get(BENCHMARK_BUY_HOLD, 0.0)),
                    "ir_vs_equal_weight_momentum": float(ir_vs_bench.get(BENCHMARK_EQUAL_WEIGHT_MOMENTUM, 0.0)),
                    "ir_vs_volatility_parity": float(ir_vs_bench.get(BENCHMARK_VOLATILITY_PARITY, 0.0)),
                    "run_dir": str(run_dir),
                    **stress_gate,
                }
            )
            combo_reports.append(
                f"entry={entry_signal} exit={exit_signal}\n{report}\n"
            )

    ranked_rows = sorted(rows, key=lambda row: (bool(row.get("stress_passed", False)), float(row["sharpe"])), reverse=True)
    leaderboard_dir = _persist_multi_signal_outputs(ranked_rows, run_namespace=run_namespace)
    (leaderboard_dir / "cancellation.json").write_text(json.dumps(cancellation.snapshot(), indent=2), encoding="utf-8")

    summary_lines = [
        "Multi-signal backtest completed.",
        f"Combinations: {len(rows)}",
        f"Leaderboard outputs: {leaderboard_dir}",
        "",
        f"Starting capital: {starting_capital:.2f}",
        f"Bet sizing mode: {bet_sizing_mode}",
        (
            "Scenario packs: "
            + (", ".join(selected_scenario_packs) if selected_scenario_packs else "none")
            + f" (available: {', '.join(list_scenario_pack_templates())})"
        ),
        f"Scenario pack templates expanded: {len(resolved_pack_templates)}",
        f"Cancellation requested: {cancellation.snapshot().get('requested', False)}",
        "Ranked combinations (by sharpe):",
    ]
    for idx, row in enumerate(ranked_rows, start=1):
        summary_lines.append(
            f"#{idx} entry={row['entry_signal']} exit={row['exit_signal']} "
            f"sharpe={row['sharpe']:.6f} total_return={row['total_return']:.6f} "
            f"cagr={row['cagr']:.6f} calmar={row['calmar']:.6f} "
            f"max_drawdown={row['max_drawdown']:.6f} volatility={row['volatility']:.6f} "
            f"sortino={row['sortino']:.6f} hit_rate={row['hit_rate']:.6f} "
            f"turnover_total={row['turnover_total']:.6f} cost_total={row['cost_total']:.6f} "
            f"alpha_vs_bh={row['alpha_vs_buy_hold']:.6f} ir_vs_bh={row['ir_vs_buy_hold']:.6f}"
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
    _assert_metric_table_compatibility(run_dir)
    metrics_path = run_dir / "metrics.json"
    rows = json.loads(metrics_path.read_text())
    metrics: dict[str, float] = {}
    for row in rows:
        metric = str(row.get("metric", ""))
        value = float(row.get("value", 0.0))
        metrics[metric] = value
    return metrics


def _load_benchmark_rows_from_run_dir(run_dir: Path) -> list[dict[str, Any]]:
    manifest_path = run_dir / "manifest.json"
    if not manifest_path.exists():
        return []
    try:
        payload = json.loads(manifest_path.read_text())
    except Exception:
        return []
    parameters = payload.get("parameters", {}) if isinstance(payload, dict) else {}
    rows = parameters.get("benchmarks", []) if isinstance(parameters, dict) else []
    if not isinstance(rows, list):
        return []
    out: list[dict[str, Any]] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        out.append({
            "benchmark": str(row.get("benchmark", "")),
            "alpha": float(row.get("alpha", 0.0)),
            "information_ratio": float(row.get("information_ratio", 0.0)),
            "tracking_error": float(row.get("tracking_error", 0.0)),
            "sharpe": float(row.get("sharpe", 0.0)),
            "total_return": float(row.get("total_return", 0.0)),
        })
    return out




def _load_stress_payload_from_run_dir(run_dir: Path) -> dict[str, Any]:
    path = run_dir / "stress_scenarios.json"
    if not path.exists():
        return {}
    try:
        payload = json.loads(path.read_text())
        return payload if isinstance(payload, dict) else {}
    except Exception:
        return {}

def _persist_multi_signal_outputs(rows: list[dict[str, Any]], *, run_namespace: str | None = None) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    namespace_prefix = f"{run_namespace}_" if run_namespace else ""
    run_dir = BACKTEST_OUTPUT_DIR / f"{namespace_prefix}tsmom_multi_signal_{timestamp}"
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
        "alpha_vs_buy_hold",
        "alpha_vs_equal_weight_momentum",
        "alpha_vs_volatility_parity",
        "ir_vs_buy_hold",
        "ir_vs_equal_weight_momentum",
        "ir_vs_volatility_parity",
        "stress_passed",
        "stress_total_scenarios",
        "stress_failed_scenarios",
        "stress_pass_rate",
        "stress_failed_scenario_names",
        "stress_fragility_index",
        "stress_survivability_score",
        "stress_survivability_min",
        "stress_model_gate_passed",
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



def _parse_preflight_critical_checks(raw: str | None) -> tuple[str, ...]:
    if raw is None or not str(raw).strip():
        return ("timestamp_consistency", "adjustment_flags", "symbol_coverage")
    checks = [item.strip() for item in str(raw).split(",") if item.strip()]
    return tuple(dict.fromkeys(checks))


def _run_preflight_or_raise(
    *,
    arrays: EngineArrayBundle,
    requested_tickers: list[str],
    start_dt: datetime,
    end_dt: datetime,
    timeframe: str,
    config: PreflightValidationConfig,
    workflow_label: str,
) -> dict[str, Any]:
    normalized_requested = [str(t).strip().upper() for t in requested_tickers if str(t).strip()]
    symbol_order = sorted(arrays.metadata.symbol_to_column.items(), key=lambda item: item[1])

    bars_total = int(arrays.date_index.size)
    missing_counts: dict[str, int] = {}
    missing_ratios: dict[str, float] = {}
    coverage_ratios: dict[str, float] = {}
    for symbol, col in symbol_order:
        ratio = float(arrays.metadata.missingness_by_symbol.get(symbol, 1.0))
        missing_ratios[symbol] = ratio
        missing_counts[symbol] = int(np.sum(arrays.missing_mask[:, col])) if arrays.missing_mask.size else 0
        coverage_ratios[symbol] = float(arrays.metadata.coverage_by_symbol.get(symbol, 0.0))

    monotonic_diffs = np.diff(arrays.date_index) if arrays.date_index.size > 1 else np.array([], dtype=np.int64)
    duplicate_count = int(np.sum(monotonic_diffs == 0)) if monotonic_diffs.size else 0
    non_monotonic_count = int(np.sum(monotonic_diffs < 0)) if monotonic_diffs.size else 0
    timestamp_ok = duplicate_count == 0 and non_monotonic_count == 0

    adjustment_violations = dict(arrays.metadata.adjustment_violations_by_symbol)
    symbols_with_adjustment_flags = sorted([sym for sym, count in adjustment_violations.items() if int(count) > 0])

    missing_bars_failed = sorted([sym for sym, ratio in missing_ratios.items() if ratio > float(config.max_missing_bars_ratio)])
    coverage_failed = sorted([sym for sym in normalized_requested if coverage_ratios.get(sym, 0.0) < float(config.min_symbol_coverage_ratio)])
    missing_symbols = sorted([sym for sym in normalized_requested if sym not in arrays.metadata.symbol_to_column])

    checks = {
        "missing_bars": {
            "status": "fail" if missing_bars_failed else "pass",
            "threshold": float(config.max_missing_bars_ratio),
            "failed_symbols": missing_bars_failed,
        },
        "timestamp_consistency": {
            "status": "pass" if timestamp_ok else "fail",
            "duplicate_timestamps": duplicate_count,
            "non_monotonic_steps": non_monotonic_count,
        },
        "adjustment_flags": {
            "status": "fail" if symbols_with_adjustment_flags else "pass",
            "failed_symbols": symbols_with_adjustment_flags,
            "violations_by_symbol": adjustment_violations,
        },
        "symbol_coverage": {
            "status": "fail" if (coverage_failed or missing_symbols) else "pass",
            "min_required": float(config.min_symbol_coverage_ratio),
            "failed_symbols": coverage_failed,
            "missing_symbols": missing_symbols,
        },
    }

    critical = list(config.critical_checks)
    critical_failures = [name for name in critical if checks.get(name, {}).get("status") == "fail"]
    status = "blocked" if (critical_failures and config.block_on_critical) else ("warn" if critical_failures else "pass")

    report = {
        "schema_version": "1.0",
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "workflow": workflow_label,
        "timeframe": timeframe,
        "window": {"start": start_dt.isoformat(), "end": end_dt.isoformat()},
        "requested_symbols": normalized_requested,
        "loaded_symbols": [sym for sym, _ in symbol_order],
        "bars_total": bars_total,
        "checks": checks,
        "critical_checks": critical,
        "critical_failures": critical_failures,
        "status": status,
        "summary": {
            "missing_counts_by_symbol": missing_counts,
            "missing_ratios_by_symbol": missing_ratios,
            "coverage_ratios_by_symbol": coverage_ratios,
            "excluded_symbols": dict(arrays.metadata.excluded_symbols),
        },
    }

    report_path = BACKTEST_OUTPUT_DIR / f"preflight_report_{workflow_label}_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.json"
    report_path.write_text(json.dumps(report, indent=2))
    if critical_failures and config.block_on_critical:
        raise ValueError(
            f"Preflight validation blocked workflow ({', '.join(critical_failures)}). Report: {report_path}"
        )
    return report


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
    raw_open_prices = arrays.raw_open_prices[selector]
    raw_close_prices = arrays.raw_close_prices[selector]
    split_factors = arrays.split_factors[selector]
    dividends = arrays.dividends[selector]

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
        strike_by_symbol=dict(arrays.metadata.strike_by_symbol),
        option_type_by_symbol=dict(arrays.metadata.option_type_by_symbol),
        multiplier_by_symbol=dict(arrays.metadata.multiplier_by_symbol),
        settlement_style_by_symbol=dict(arrays.metadata.settlement_style_by_symbol),
        borrow_availability_tier_by_symbol=dict(arrays.metadata.borrow_availability_tier_by_symbol),
        financing_benchmark_by_symbol=dict(arrays.metadata.financing_benchmark_by_symbol),
        pit_membership_violations_by_symbol=dict(arrays.metadata.pit_membership_violations_by_symbol),
        adjustment_violations_by_symbol=dict(arrays.metadata.adjustment_violations_by_symbol),
        delisted_symbols=list(arrays.metadata.delisted_symbols),
        survivorship_bias_flags_by_symbol=dict(arrays.metadata.survivorship_bias_flags_by_symbol),
        leakage_flags_by_symbol=dict(arrays.metadata.leakage_flags_by_symbol),
        data_fingerprint={**dict(getattr(arrays.metadata, "data_fingerprint", {}) or {}), "resampled_timeframe": timeframe},
    )
    return EngineArrayBundle(
        date_index=date_index,
        open_prices=open_prices,
        close_prices=close_prices,
        raw_open_prices=raw_open_prices,
        raw_close_prices=raw_close_prices,
        split_factors=split_factors,
        dividends=dividends,
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




def _build_trade_explainability_rows(
    *,
    trade_log_rows: list[dict[str, object]],
    fill_rows: list[dict[str, object]],
    risk_diagnostics: dict[str, np.ndarray],
    costs_bps: float,
) -> list[dict[str, object]]:
    explain_rows: list[dict[str, object]] = []
    fills_by_bar: dict[int, list[dict[str, object]]] = {}
    for fill in fill_rows:
        bar = int(fill.get("bar_index", 0))
        fills_by_bar.setdefault(bar, []).append(fill)

    margin = _to_numpy_1d(risk_diagnostics.get("margin_utilization", np.zeros(0, dtype=float)))
    model_conf = _to_numpy_1d(risk_diagnostics.get("model_confidence", np.zeros(0, dtype=float)))
    var_95 = _to_numpy_1d(risk_diagnostics.get("var_95", np.zeros(0, dtype=float)))

    for idx, row in enumerate(trade_log_rows):
        symbol = str(row.get("symbol", ""))
        trade = float(row.get("trade", 0.0))
        shares_delta = float(row.get("shares_delta", 0.0))
        price = float(row.get("price", 0.0))
        trade_notional = abs(shares_delta * price)

        feature_values = {
            "trade_signal": trade,
            "trade_notional": trade_notional,
            "trade_cost_bps": float(costs_bps),
            "running_pnl": float(row.get("running_pnl", 0.0)),
        }
        if idx < margin.size:
            feature_values["margin_utilization"] = float(margin[idx])
        if idx < model_conf.size:
            feature_values["model_confidence"] = float(model_conf[idx])
        if idx < var_95.size:
            feature_values["var_95"] = float(var_95[idx])

        risk_constraints = {
            "max_abs_trade": max(1.0, abs(trade) * 1.1),
            "max_notional": max(1.0, trade_notional * 1.05),
        }
        if idx < margin.size:
            risk_constraints["max_margin_utilization"] = max(0.01, float(margin[idx]))

        artifact = build_trade_explainability(
            trade_id=f"trade_{idx}",
            timestamp=str(row.get("timestamp", "")),
            symbol=symbol,
            requested_size=trade,
            sized_trade=trade,
            feature_values=feature_values,
            baseline_values={"trade_signal": 0.0, "running_pnl": 0.0, "trade_cost_bps": 0.0},
            feature_uncertainty={"trade_signal": 0.2, "running_pnl": 0.1, "trade_notional": max(1.0, trade_notional * 0.2)},
            risk_constraints=risk_constraints,
        )
        explain_rows.append(
            {
                "trade_id": f"trade_{idx}",
                "timestamp": row.get("timestamp", ""),
                "symbol": symbol,
                "top_drivers": artifact.top_drivers,
                "uncertainty": artifact.uncertainty,
                "risk_constraints": artifact.risk_constraints,
                "decision_trace": artifact.decision_trace,
                "counterfactual": artifact.counterfactual,
                "red_flags": artifact.red_flags,
                "fill_context": fills_by_bar.get(idx, []),
            }
        )
    return explain_rows


def _format_explainability_report(rows: list[dict[str, object]], max_rows: int = 30) -> str:
    lines = ["Explainability Governance Report", "------------------------------"]
    if not rows:
        lines.append("No explainability records available.")
        return "\n".join(lines)
    lines.append("Template: model governance + debugging")
    lines.append("fields: top_drivers, uncertainty, risk_constraints, decision_trace, counterfactual, red_flags")
    red_flagged = 0
    for row in rows[:max_rows]:
        flags = row.get("red_flags", [])
        if isinstance(flags, list) and flags:
            red_flagged += 1
        lines.append(
            f"- {row.get('timestamp', '')} {row.get('symbol', '')}: drivers={len(row.get('top_drivers', []))} red_flags={len(flags) if isinstance(flags, list) else 0}"
        )
    if len(rows) > max_rows:
        lines.append(f"... ({len(rows) - max_rows} more rows)")
    lines.append(
        f"Red-flag detector summary: flagged_trades={red_flagged}, flag_rate={(red_flagged / max(1, len(rows))):.2%}"
    )
    return "\n".join(lines)

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

def _build_stress_scenario_definitions(
    *,
    timestamps: np.ndarray,
    returns: np.ndarray,
    controls: dict[str, Any] | None = None,
    scenario_packs: list[str] | None = None,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    n_obs = int(np.asarray(returns).size)
    if n_obs <= 0:
        return rows

    cfg = dict(controls or {})
    window_frac = float(cfg.get("historical_window_fraction", 0.2))
    window = max(1, min(n_obs, int(max(5, n_obs * window_frac))))
    rows.append({"name": "historical_recent_window", "type": "historical_window", "start": max(0, n_obs - window), "end": n_obs})
    rows.append({"name": "historical_first_window", "type": "historical_window", "start": 0, "end": window})

    if bool(cfg.get("enable_historical_replay_regimes", True)):
        abs_rets = np.abs(np.asarray(returns, dtype=float))
        vol_window = max(2, min(n_obs, int(cfg.get("historical_replay_window_bars", 20))))
        rolling = np.convolve(abs_rets, np.ones(vol_window, dtype=float) / vol_window, mode="valid") if abs_rets.size >= vol_window else abs_rets
        if rolling.size:
            peak_idx = int(np.argmax(rolling))
            start = max(0, min(peak_idx, n_obs - vol_window))
            rows.append({"name": "historical_volatility_shock_window", "type": "historical_window", "start": start, "end": min(n_obs, start + vol_window)})

    rows.extend(
        [
            {
                "name": "synthetic_returns_crash",
                "type": "synthetic_shock",
                "transforms": {"returns_multiplier": 1.8, "returns_shift": -0.0025},
            },
            {
                "name": "synthetic_spread_widening",
                "type": "synthetic_shock",
                "transforms": {"spread_multiplier": 2.0, "returns_multiplier": 0.9},
            },
            {
                "name": "synthetic_liquidity_drought",
                "type": "synthetic_shock",
                "transforms": {"liquidity_multiplier": 0.5, "returns_multiplier": 0.85},
            },
            {
                "name": "synthetic_borrow_unavailable",
                "type": "synthetic_shock",
                "transforms": {"borrow_availability": 0.2, "returns_shift": -0.001},
            },
            {
                "name": "synthetic_correlation_breakdown",
                "type": "synthetic_shock",
                "transforms": {"correlation_breakdown": 1.0, "returns_multiplier": 0.8},
            },
            {
                "name": "synthetic_jump_cluster",
                "type": "synthetic_shock",
                "transforms": {
                    "jump_magnitude": float(cfg.get("synthetic_jump_magnitude", 0.02)),
                    "jump_interval": int(cfg.get("synthetic_jump_interval", 7)),
                    "vol_cluster_multiplier": float(cfg.get("synthetic_vol_cluster_multiplier", 1.6)),
                },
            },
            {
                "name": "synthetic_liquidity_spread_overlay",
                "type": "synthetic_shock",
                "transforms": {
                    "spread_multiplier": float(cfg.get("overlay_spread_multiplier", 2.5)),
                    "liquidity_multiplier": float(cfg.get("overlay_liquidity_multiplier", 0.4)),
                    "returns_multiplier": 0.85,
                },
            },
        ]
    )
    rows.extend(resolve_scenario_pack_templates(scenario_packs))
    rows.extend(
        _build_synthetic_regime_switch_paths(
            returns=np.asarray(returns, dtype=float).reshape(-1),
            prices=np.asarray(prices_from_returns_proxy(returns), dtype=float),
            controls=cfg,
        )
    )
    return rows


def prices_from_returns_proxy(returns: np.ndarray) -> np.ndarray:
    arr = np.asarray(returns, dtype=float).reshape(-1)
    if arr.size == 0:
        return np.zeros((0, 1), dtype=float)
    px = np.cumprod(1.0 + arr)
    return np.column_stack([px, px])


def _build_synthetic_regime_switch_paths(
    *,
    returns: np.ndarray,
    prices: np.ndarray,
    controls: dict[str, Any],
) -> list[dict[str, Any]]:
    n_obs = int(returns.size)
    if n_obs <= 0:
        return []

    seed = int(controls.get("synthetic_path_seed", 314159))
    n_paths = max(1, int(controls.get("synthetic_path_count", 3)))
    jump_mag = float(controls.get("synthetic_jump_magnitude", 0.02))
    rng = np.random.default_rng(seed)
    base_vol = float(np.std(returns, ddof=1)) if n_obs > 1 else float(np.std(returns))
    base_vol = max(base_vol, 1e-4)

    rows: list[dict[str, Any]] = []
    for idx in range(n_paths):
        regime_lengths = []
        total = 0
        while total < n_obs:
            seg = int(rng.integers(max(6, n_obs // 12), max(8, n_obs // 4) + 1))
            regime_lengths.append(seg)
            total += seg

        path = np.zeros(n_obs, dtype=float)
        cursor = 0
        prev = 0.0
        for regime_id, seg_len in enumerate(regime_lengths):
            if cursor >= n_obs:
                break
            end = min(n_obs, cursor + seg_len)
            regime_mu = float(rng.normal(loc=-0.0002 * (1 + regime_id), scale=0.0008))
            regime_vol = float(base_vol * rng.uniform(0.8, 2.5))
            dof = int(rng.integers(3, 7))
            for t in range(cursor, end):
                vol_cluster = regime_vol * (1.0 + 0.65 * min(abs(prev) / base_vol, 2.0))
                fat_tail = float(rng.standard_t(dof) * vol_cluster)
                jump = -jump_mag * rng.uniform(0.8, 1.4) if rng.random() < 0.08 else 0.0
                path[t] = regime_mu + fat_tail + jump
                prev = path[t]
            cursor = end

        corr_component = np.zeros_like(path)
        if np.asarray(prices).ndim == 2 and prices.shape[1] > 1:
            asset_rets = np.diff(np.asarray(prices, dtype=float), axis=0) / np.where(prices[:-1] == 0.0, 1.0, prices[:-1])
            if asset_rets.size:
                dispersion = np.std(asset_rets, axis=1)
                scale = float(np.mean(dispersion)) if np.mean(dispersion) > 0 else 1.0
                corr_component[1 : 1 + min(dispersion.size, n_obs - 1)] = (dispersion / scale - 1.0) * 0.0012
        path = path + corr_component

        rows.append(
            {
                "name": f"synthetic_path_regime_switch_{idx + 1}",
                "type": "synthetic_path",
                "path_returns": path.tolist(),
                "stress_characteristics": {
                    "regime_switches": max(0, len(regime_lengths) - 1),
                    "volatility_clustering": True,
                    "fat_tails": True,
                    "jumps": True,
                    "liquidity_drought": True,
                    "correlation_breakdown": True,
                },
            }
        )
    return rows


def _compute_scenario_metrics(returns: np.ndarray) -> dict[str, float]:
    arr = np.asarray(returns, dtype=float).reshape(-1)
    if arr.size == 0:
        return {"total_return": 0.0, "max_drawdown": 0.0, "sharpe": 0.0}
    equity = np.cumprod(1.0 + arr)
    peak = np.maximum.accumulate(equity)
    drawdown = equity / np.where(peak == 0.0, 1.0, peak) - 1.0
    mean = float(np.mean(arr))
    vol = float(np.std(arr, ddof=1)) if arr.size > 1 else 0.0
    sharpe = (mean / vol) * np.sqrt(252.0) if vol > 1e-12 else 0.0
    return {
        "total_return": float(equity[-1] - 1.0),
        "max_drawdown": float(np.min(drawdown)) if drawdown.size else 0.0,
        "sharpe": float(sharpe),
    }


def _apply_scenario_transforms(*, base_returns: np.ndarray, prices: np.ndarray, transforms: dict[str, Any]) -> np.ndarray:
    shocked = np.asarray(base_returns, dtype=float).copy()
    returns_multiplier = float(transforms.get("returns_multiplier", 1.0))
    returns_shift = float(transforms.get("returns_shift", 0.0))
    shocked = shocked * returns_multiplier + returns_shift

    spread_mult = float(transforms.get("spread_multiplier", 1.0))
    liquidity_mult = float(transforms.get("liquidity_multiplier", 1.0))
    borrow_avail = float(transforms.get("borrow_availability", 1.0))
    corr_break = float(transforms.get("correlation_breakdown", 0.0))

    friction_penalty = 0.0001 * max(0.0, spread_mult - 1.0)
    liquidity_penalty = 0.0001 * max(0.0, 1.0 - liquidity_mult)
    borrow_penalty = 0.0002 * max(0.0, 1.0 - borrow_avail)
    shocked = shocked - (friction_penalty + liquidity_penalty + borrow_penalty)

    if corr_break > 0.0 and np.asarray(prices).ndim == 2 and prices.shape[1] > 1:
        asset_rets = np.diff(np.asarray(prices, dtype=float), axis=0) / np.where(prices[:-1] == 0.0, 1.0, prices[:-1])
        dispersion = np.std(asset_rets, axis=1)
        if dispersion.size:
            scale = np.mean(dispersion) if np.mean(dispersion) > 0 else 1.0
            dispersion_effect = np.zeros_like(shocked)
            dispersion_effect[1 : 1 + min(dispersion.size, shocked.size - 1)] = (dispersion / scale - 1.0) * 0.0005 * corr_break
            shocked = shocked + dispersion_effect

    jump_mag = float(transforms.get("jump_magnitude", 0.0))
    jump_interval = max(1, int(transforms.get("jump_interval", 0) or 1))
    if jump_mag > 0.0 and shocked.size:
        jump_mask = (np.arange(shocked.size) % jump_interval) == 0
        shocked[jump_mask] = shocked[jump_mask] - jump_mag

    cluster_mult = float(transforms.get("vol_cluster_multiplier", 1.0))
    if cluster_mult > 1.0 and shocked.size > 2:
        rolling_abs = np.abs(shocked)
        if rolling_abs.size > 1:
            local = np.zeros_like(rolling_abs)
            local[1:] = 0.5 * (rolling_abs[:-1] + rolling_abs[1:])
            scale = np.mean(local[1:]) if np.mean(local[1:]) > 0 else 1.0
            shocked = shocked * (1.0 + (cluster_mult - 1.0) * np.clip(local / scale - 1.0, 0.0, 2.0))

    return shocked


def _run_stress_scenario_wrappers(*, returns: np.ndarray, prices: np.ndarray, scenario_definitions: list[dict[str, Any]]) -> list[dict[str, Any]]:
    base_returns = np.asarray(returns, dtype=float).reshape(-1)
    rows: list[dict[str, Any]] = []
    for scenario in scenario_definitions:
        name = str(scenario.get("name", "unnamed"))
        kind = str(scenario.get("type", "synthetic_shock"))
        if kind == "historical_window":
            start = int(scenario.get("start", 0))
            end = int(scenario.get("end", base_returns.size))
            shocked = base_returns[max(0, start) : max(0, end)]
        elif kind == "synthetic_path":
            path_payload = np.asarray(scenario.get("path_returns", []), dtype=float).reshape(-1)
            shocked = path_payload if path_payload.size else base_returns
        else:
            transforms = scenario.get("transforms", {})
            shocked = _apply_scenario_transforms(
                base_returns=base_returns,
                prices=prices,
                transforms=transforms if isinstance(transforms, dict) else {},
            )
        metrics = _compute_scenario_metrics(shocked)
        rows.append(
            {
                "name": name,
                "type": kind,
                "observation_count": int(shocked.size),
                "pnl_total": float(np.sum(shocked)),
                "metrics": metrics,
                "stress_characteristics": scenario.get("stress_characteristics", {}),
            }
        )
    return rows


def _stress_gate_summary(scenario_payload: dict[str, Any], controls: dict[str, Any] | None = None) -> dict[str, Any]:
    cfg = dict(controls or {})
    checks = scenario_payload.get("scenario_guardrails", []) if isinstance(scenario_payload, dict) else []
    attributions = scenario_payload.get("scenario_attribution", []) if isinstance(scenario_payload, dict) else []
    if not isinstance(checks, list):
        checks = []
    if not isinstance(attributions, list):
        attributions = []
    failed = [row for row in checks if isinstance(row, dict) and not bool(row.get("passed", False))]
    failed_names = [str(row.get("scenario", "unnamed")) for row in failed]
    total = len(checks)
    drawdown_shocks = [abs(float(row.get("delta_max_drawdown", 0.0))) for row in attributions if isinstance(row, dict)]
    sharpe_shocks = [max(0.0, -float(row.get("delta_sharpe", 0.0))) for row in attributions if isinstance(row, dict)]
    return_shocks = [max(0.0, -float(row.get("delta_total_return", 0.0))) for row in attributions if isinstance(row, dict)]
    fragility_index = float(
        0.5 * (float(np.mean(drawdown_shocks)) if drawdown_shocks else 0.0)
        + 0.3 * (float(np.mean(sharpe_shocks)) if sharpe_shocks else 0.0)
        + 0.2 * (float(np.mean(return_shocks)) if return_shocks else 0.0)
    )
    survivability = max(0.0, min(1.0, (1.0 - fragility_index) * (float((total - len(failed)) / total) if total else 1.0)))
    threshold = float(cfg.get("stress_survivability_min", 0.55))
    model_gate_passed = (len(failed) == 0) and (survivability >= threshold)
    return {
        "stress_passed": len(failed) == 0,
        "stress_total_scenarios": int(total),
        "stress_failed_scenarios": int(len(failed)),
        "stress_pass_rate": float((total - len(failed)) / total) if total else 1.0,
        "stress_failed_scenario_names": failed_names,
        "stress_fragility_index": fragility_index,
        "stress_survivability_score": survivability,
        "stress_survivability_min": threshold,
        "stress_model_gate_passed": bool(model_gate_passed),
    }


def _compute_stream_metrics(returns: np.ndarray) -> dict[str, float]:
    arr = np.asarray(returns, dtype=float).reshape(-1)
    if arr.size == 0:
        return {"total_return": 0.0, "mean_return": 0.0, "volatility": 0.0, "sharpe": 0.0, "max_drawdown": 0.0}
    equity = np.cumprod(1.0 + arr)
    total = float(equity[-1] - 1.0)
    mean = float(np.mean(arr))
    vol = float(np.std(arr, ddof=1)) if arr.size > 1 else 0.0
    sharpe = float(mean / vol * np.sqrt(252.0)) if vol > 0.0 else 0.0
    running_peak = np.maximum.accumulate(equity)
    safe_peak = np.where(running_peak == 0.0, 1.0, running_peak)
    drawdown = equity / safe_peak - 1.0
    return {
        "total_return": total,
        "mean_return": mean,
        "volatility": vol,
        "sharpe": sharpe,
        "max_drawdown": float(np.min(drawdown)) if drawdown.size else 0.0,
    }


def _build_regime_ensemble_comparison(
    *,
    prices: np.ndarray,
    missing_mask: np.ndarray,
    regime_labels: np.ndarray,
) -> dict[str, Any]:
    px = np.asarray(prices, dtype=float)
    rets = np.zeros_like(px, dtype=float)
    rets[1:] = px[1:] / np.where(px[:-1] == 0.0, 1.0, px[:-1]) - 1.0
    market_ret = np.nanmean(np.where(np.isfinite(rets), rets, 0.0), axis=1)

    strat_defs = [
        ("ts_momentum", {"lookback_days": 60, "skip_days": 5}),
        ("ma_trend", {"ma_window": 30}),
        ("mean_reversion", {"lookback_days": 20, "zscore_threshold": 1.0}),
    ]
    strategy_names = tuple(name for name, _ in strat_defs)
    sleeves = np.zeros((px.shape[0], len(strat_defs)), dtype=float)

    for col, (name, params) in enumerate(strat_defs):
        entry_cfg = parse_entry_signal_config(name, params, default_lookback_days=60, default_skip_days=5)
        exit_cfg = parse_exit_signal_config("none", {}, default_lookback_days=60, default_skip_days=5)
        targets = build_targets(
            close_prices=px,
            missing_mask=missing_mask,
            entry_config=entry_cfg,
            exit_config=exit_cfg,
        )
        sleeves[:, col] = np.nanmean(targets, axis=1) * market_ret

    always_weights = np.full(len(strat_defs), 1.0 / len(strat_defs), dtype=float)
    always_on = sleeves @ always_weights

    gated_map: dict[str, tuple[float, ...]] = {}
    weighted_map: dict[str, tuple[float, ...]] = {}
    for label in sorted(set(str(x) for x in regime_labels.tolist())):
        is_up = "trend_up" in label and "macro_risk_on" in label
        is_risk_off = "macro_risk_off" in label or "vol_high" in label
        if is_up:
            gated_map[label] = (1.0, 0.0, 0.0)
            weighted_map[label] = (0.70, 0.20, 0.10)
        elif is_risk_off:
            gated_map[label] = (0.0, 0.0, 1.0)
            weighted_map[label] = (0.15, 0.20, 0.65)
        else:
            gated_map[label] = (0.0, 1.0, 0.0)
            weighted_map[label] = (0.25, 0.55, 0.20)

    gated_cfg = RegimeMetaPolicyConfig(
        strategy_names=strategy_names,
        regime_weight_map=gated_map,
        default_weights=tuple(always_weights.tolist()),
        turnover_limit=1.0,
        max_weight=1.0,
        min_weight=0.0,
    )
    gated_schedule, gated_diag = build_regime_weight_schedule(regime_labels=regime_labels, config=gated_cfg)
    regime_gated = np.sum(sleeves * gated_schedule, axis=1)

    weighted_cfg = RegimeMetaPolicyConfig(
        strategy_names=strategy_names,
        regime_weight_map=weighted_map,
        default_weights=tuple(always_weights.tolist()),
        turnover_limit=0.25,
        max_weight=0.70,
        min_weight=0.05,
    )
    weighted_schedule, weighted_diag = build_regime_weight_schedule(regime_labels=regime_labels, config=weighted_cfg)
    regime_weighted = np.sum(sleeves * weighted_schedule, axis=1)

    benchmark_equal_weight_momentum = sleeves[:, 0]
    benchmark_volatility_parity = sleeves @ np.array([0.4, 0.2, 0.4], dtype=float)

    series = {
        "always_on_baseline": always_on,
        "regime_gated": regime_gated,
        "regime_weighted": regime_weighted,
        BENCHMARK_BUY_HOLD: market_ret,
        BENCHMARK_EQUAL_WEIGHT_MOMENTUM: benchmark_equal_weight_momentum,
        BENCHMARK_VOLATILITY_PARITY: benchmark_volatility_parity,
    }
    comparison_rows = []
    baseline_stream = np.asarray(series["always_on_baseline"], dtype=float)
    baseline_metrics = _compute_stream_metrics(baseline_stream)
    baseline_sharpe = float(baseline_metrics.get("sharpe", 0.0))
    for name, stream in series.items():
        row = {"policy": name}
        row.update(_compute_stream_metrics(stream))
        rel = _compute_relative_alpha_ir(np.asarray(stream, dtype=float), baseline_stream)
        row["relative_alpha_vs_always_on"] = float(rel["alpha"])
        row["relative_ir_vs_always_on"] = float(rel["information_ratio"])
        row["delta_sharpe_vs_always_on"] = float(row.get("sharpe", 0.0) - baseline_sharpe)
        comparison_rows.append(row)

    regime_attribution = {
        name: attribute_pnl_by_regime(pnl=stream, regime_labels=regime_labels)
        for name, stream in series.items()
    }
    baseline_lookup = {
        str(row.get("regime")): float(row.get("pnl_mean", 0.0))
        for row in regime_attribution.get("always_on_baseline", [])
        if isinstance(row, dict)
    }
    uplift_by_bucket: list[dict[str, float | str | int]] = []
    for row in regime_attribution.get("regime_weighted", []):
        if not isinstance(row, dict):
            continue
        regime = str(row.get("regime", ""))
        base = baseline_lookup.get(regime, 0.0)
        uplift_by_bucket.append(
            {
                "regime": regime,
                "bars": int(row.get("bars", 0)),
                "baseline_pnl_mean": base,
                "routed_pnl_mean": float(row.get("pnl_mean", 0.0)),
                "uplift_pnl_mean": float(row.get("pnl_mean", 0.0) - base),
            }
        )

    return {
        "strategy_names": list(strategy_names),
        "comparison": comparison_rows,
        "regime_attribution": regime_attribution,
        "uplift_by_regime_bucket": uplift_by_bucket,
        "meta_policy_diagnostics": {
            "gated": {
                "mean_turnover": float(np.mean(gated_diag.get("meta_turnover", np.array([0.0])))),
                "max_concentration": float(np.max(gated_diag.get("meta_concentration", np.array([0.0])))),
            },
            "weighted": {
                "mean_turnover": float(np.mean(weighted_diag.get("meta_turnover", np.array([0.0])))),
                "max_concentration": float(np.max(weighted_diag.get("meta_concentration", np.array([0.0])))),
            },
        },
    }


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
    scenario_payload: dict[str, Any] | None = None,
    regime_labels: np.ndarray | None = None,
    regime_probabilities: np.ndarray | None = None,
    regime_states: np.ndarray | None = None,
    regime_pnl_attribution: list[dict[str, float | str | int]] | None = None,
    regime_ensemble_report: dict[str, Any] | None = None,
    governance: dict[str, Any] | None = None,
    corporate_action_splits: np.ndarray | None = None,
    corporate_action_dividends: np.ndarray | None = None,
    attribution_payload: dict[str, Any] | None = None,
    fill_rows: list[dict[str, Any]] | None = None,
    slippage_calibration_selection: dict[str, Any] | None = None,
    selected_test_suite: str = "custom",
    suite_composition: dict[str, Any] | None = None,
) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    run_dir = BACKTEST_OUTPUT_DIR / f"tsmom_backtest_{timestamp}"
    run_dir.mkdir(parents=True, exist_ok=True)

    governance_payload = _build_governance_metadata(governance)
    computed_checks = _evaluate_governance_gate_checks(
        metrics=metrics,
        fold_rows=None,
        governance=governance_payload,
    )
    governance_payload["gate_checks"].update(computed_checks)
    required_checks = governance_payload.get("promotion_required_checks", [])
    governance_payload["missing_required_checks"] = [
        name for name in required_checks if not governance_payload["gate_checks"].get(name, False)
    ]
    governance_payload["is_promotion_ready"] = not governance_payload["missing_required_checks"]
    governance_payload["audit_trail"].append({
        "timestamp": datetime.now().isoformat(),
        "event": "gate_checks_evaluated",
        "gate_checks": dict(governance_payload["gate_checks"]),
        "missing_required_checks": list(governance_payload["missing_required_checks"]),
    })

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
    _write_risk_monitoring_artifacts(
        run_dir=run_dir,
        timestamps=time_strings,
        symbol_order=symbol_order,
        trades=trades,
        returns=returns,
        diagnostics=risk_diagnostics,
        regime_labels=regime_labels,
        regime_probabilities=regime_probabilities,
        regime_states=regime_states,
    )
    if regime_labels is not None:
        _write_regime_labels_csv_json(
            run_dir=run_dir,
            timestamps=time_strings,
            regime_labels=np.asarray(regime_labels, dtype=object),
            regime_probabilities=np.asarray(regime_probabilities, dtype=float) if regime_probabilities is not None else None,
            regime_states=np.asarray(regime_states, dtype=object) if regime_states is not None else None,
        )
    if regime_pnl_attribution is not None:
        _write_regime_pnl_attribution(
            run_dir=run_dir,
            rows=regime_pnl_attribution,
        )
    if regime_ensemble_report is not None:
        _write_regime_ensemble_comparison(run_dir=run_dir, report=regime_ensemble_report)

    _write_corporate_actions_applied_csv(
        run_dir=run_dir,
        timestamps=time_strings,
        symbol_order=symbol_order,
        split_factors=corporate_action_splits,
        dividends=corporate_action_dividends,
    )

    metrics_rows = [{"metric": key, "value": float(value)} for key, value in metrics.items()]
    metrics_csv = run_dir / "metrics.csv"
    with metrics_csv.open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["metric", "value"])
        writer.writeheader()
        writer.writerows(metrics_rows)
    (run_dir / "metrics.json").write_text(json.dumps(metrics_rows, indent=2))
    (run_dir / "metric_schema_version.txt").write_text(f"{CANONICAL_METRIC_SCHEMA_VERSION}\n")
    metric_table_names = ["metrics"]
    if attribution_payload:
        write_attribution_artifacts(run_dir=run_dir, payload=attribution_payload)
        metric_table_names.extend(["attribution_timeseries", "attribution_summary"])
    metric_tables = _write_metric_table_manifest(run_dir=run_dir, run_type="backtest", table_names=metric_table_names)

    if robustness_report is not None:
        _write_robustness_report(run_dir=run_dir, robustness_report=robustness_report)
        _write_capacity_frontier_artifacts(run_dir=run_dir, robustness_report=robustness_report)
    if scenario_payload is not None:
        (run_dir / "stress_scenarios.json").write_text(json.dumps(scenario_payload, indent=2))

    _write_slippage_decomposition_artifacts(
        run_dir=run_dir,
        fill_rows=list(fill_rows or []),
        regime_labels=regime_labels,
        slippage_calibration_selection=dict(slippage_calibration_selection or {}),
    )

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

    manifest = _build_run_manifest(
        run_type="backtest",
        parameters=parameters or {},
        data_snapshot=data_snapshot or {},
        random_seed=random_seed,
        governance=governance_payload,
        metric_tables=metric_tables,
        extra_fingerprint_payload={
            "selected_test_suite": str(selected_test_suite or "custom"),
            "suite_composition": dict(suite_composition or {}),
        },
        result_summary={
            "metrics": {k: float(v) for k, v in metrics.items() if isinstance(v, (int, float))},
            **_collect_artifact_inventory(run_dir),
        },
    )
    (run_dir / "manifest.json").write_text(json.dumps(manifest, indent=2))
    (run_dir / "artifact_metadata.json").write_text(json.dumps({
        "schema_version": "1.0",
        "run_type": "backtest",
        "random_seeds": {"run_seed": random_seed, "python_random_seed": random_seed, "numpy_random_seed": random_seed},
        "data_fingerprint": dict((data_snapshot or {}).get("data_fingerprint", {})) if isinstance(data_snapshot, dict) else {},
        "selected_test_suite": str(selected_test_suite or "custom"),
        "suite_composition": dict(suite_composition or {}),
    }, indent=2))
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
            "run_id": manifest["run_id"],
            "config_hash": manifest["config_hash"],
            "config_checksum": manifest["config_checksum"],
            "data_snapshot_checksum": manifest["data_snapshot_checksum"],
            "manifest_checksum": manifest["manifest_checksum"],
            "parameters": parameters or {},
            "data_snapshot_identifiers": data_snapshot or {},
            "metrics": metrics,
            "significance": {"robustness": robustness_report} if robustness_report else {},
            "governance": governance_payload,
            "model_artifacts": _collect_artifact_inventory(run_dir).get("model_artifacts", []),
            "plot_artifacts": _collect_artifact_inventory(run_dir).get("plot_artifacts", []),
            "metric_artifacts": _collect_artifact_inventory(run_dir).get("metric_artifacts", []),
            "reproducibility_metadata": manifest.get("reproducibility_metadata", {}),
        }
    )

    return run_dir



def _write_corporate_actions_applied_csv(
    *,
    run_dir: Path,
    timestamps: list[str],
    symbol_order: list[str],
    split_factors: np.ndarray | None,
    dividends: np.ndarray | None,
) -> None:
    csv_path = run_dir / "corporate_actions_applied.csv"
    rows: list[dict[str, object]] = []
    if split_factors is not None and dividends is not None:
        split_vals = np.asarray(split_factors, dtype=float)
        div_vals = np.asarray(dividends, dtype=float)
        if split_vals.shape == div_vals.shape and split_vals.shape[0] == len(timestamps):
            for row_idx, ts in enumerate(timestamps):
                for col_idx, symbol in enumerate(symbol_order):
                    split = float(split_vals[row_idx, col_idx])
                    div = float(div_vals[row_idx, col_idx])
                    if abs(split - 1.0) > 1e-12 or abs(div) > 1e-12:
                        rows.append(
                            {
                                "timestamp": ts,
                                "symbol": symbol,
                                "split_factor": split,
                                "dividend": div,
                            }
                        )
    with csv_path.open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["timestamp", "symbol", "split_factor", "dividend"])
        writer.writeheader()
        writer.writerows(rows)



def _extract_run_dir_from_output(output_text: str) -> Path | None:
    marker = "Saved outputs to:"
    if marker not in output_text:
        return None
    tail = output_text.split(marker)[-1].strip().splitlines()[0].strip()
    if not tail:
        return None
    run_dir = Path(tail)
    return run_dir if run_dir.exists() else None


def _write_slippage_decomposition_artifacts(
    *,
    run_dir: Path,
    fill_rows: list[dict[str, Any]],
    regime_labels: np.ndarray | None,
    slippage_calibration_selection: dict[str, Any] | None = None,
) -> None:
    if not fill_rows:
        payload = {
            "expected_vs_observed_fill_slippage_drift_bps": 0.0,
            "by_regime": [],
            "by_liquidity_bucket": [],
            "calibration": dict(slippage_calibration_selection or {}),
        }
        (run_dir / "slippage_decomposition.json").write_text(json.dumps(payload, indent=2))
        return

    labels = np.asarray(regime_labels, dtype=object) if regime_labels is not None else np.asarray([], dtype=object)
    by_regime: dict[str, dict[str, float]] = {}
    by_liquidity: dict[str, dict[str, float]] = {}
    total_expected = 0.0
    total_observed = 0.0

    for row in fill_rows:
        requested = float(row.get("requested_size", 0.0) or 0.0)
        filled = float(row.get("filled_size", 0.0) or 0.0)
        residual = float(row.get("residual_size", 0.0) or 0.0)
        participation = max(float(row.get("participation_rate", 0.0) or 0.0), 0.0)
        expected_bps = participation * 10.0
        fill_ratio = abs(filled) / max(abs(requested), 1e-9) if requested != 0.0 else 0.0
        observed_bps = expected_bps * (1.0 + max(0.0, 1.0 - fill_ratio) + min(abs(residual) / max(abs(requested), 1e-9), 1.0))
        drift = observed_bps - expected_bps
        total_expected += expected_bps
        total_observed += observed_bps

        bar_index = int(row.get("bar_index", 0) or 0)
        regime = str(labels[bar_index]) if 0 <= bar_index < labels.size else "unlabeled"
        if participation < 0.10:
            bucket = "low"
        elif participation < 0.30:
            bucket = "medium"
        else:
            bucket = "high"

        slot = by_regime.setdefault(regime, {"count": 0.0, "expected_bps": 0.0, "observed_bps": 0.0, "drift_bps": 0.0})
        slot["count"] += 1
        slot["expected_bps"] += expected_bps
        slot["observed_bps"] += observed_bps
        slot["drift_bps"] += drift

        lslot = by_liquidity.setdefault(bucket, {"count": 0.0, "expected_bps": 0.0, "observed_bps": 0.0, "drift_bps": 0.0})
        lslot["count"] += 1
        lslot["expected_bps"] += expected_bps
        lslot["observed_bps"] += observed_bps
        lslot["drift_bps"] += drift

    payload = {
        "expected_vs_observed_fill_slippage_drift_bps": float(total_observed - total_expected),
        "by_regime": [{"regime": key, **vals} for key, vals in sorted(by_regime.items())],
        "by_liquidity_bucket": [{"liquidity_bucket": key, **vals} for key, vals in sorted(by_liquidity.items())],
        "calibration": dict(slippage_calibration_selection or {}),
    }
    (run_dir / "slippage_decomposition.json").write_text(json.dumps(payload, indent=2))


def _write_suite_artifact_bundle(
    *,
    run_dir: Path,
    suite_key: str,
    suite_composition: dict[str, Any],
    governance_metadata: dict[str, Any],
    stress_controls: dict[str, Any],
) -> None:
    metrics = json.loads((run_dir / "metrics.json").read_text()) if (run_dir / "metrics.json").exists() else []
    stress = json.loads((run_dir / "stress_scenarios.json").read_text()) if (run_dir / "stress_scenarios.json").exists() else {}
    manifest = json.loads((run_dir / "manifest.json").read_text()) if (run_dir / "manifest.json").exists() else {}
    governance = manifest.get("governance", {}) if isinstance(manifest, dict) else {}
    slippage_diag = json.loads((run_dir / "slippage_decomposition.json").read_text()) if (run_dir / "slippage_decomposition.json").exists() else {}
    bundle = {
        "suite": str(suite_key or "custom"),
        "composition": dict(suite_composition or {}),
        "execution_stack": (manifest.get("parameters", {}) if isinstance(manifest, dict) else {}).get("execution_model", ""),
        "slippage_calibration_source": (manifest.get("parameters", {}) if isinstance(manifest, dict) else {}).get("execution_model_calibration", {}).get("source", "config"),
        "fee_borrow_assumptions": {
            "costs_bps": (manifest.get("parameters", {}) if isinstance(manifest, dict) else {}).get("costs_bps", 0.0),
            "carry_model": (manifest.get("parameters", {}) if isinstance(manifest, dict) else {}).get("carry_model", ""),
        },
        "walk_forward_cpcv": {
            "cv_scheme": (manifest.get("parameters", {}) if isinstance(manifest, dict) else {}).get("wf_cv_scheme", "walk_forward"),
        },
        "scenario_packs": (manifest.get("parameters", {}) if isinstance(manifest, dict) else {}).get("scenario_packs", []),
        "governance_gates": governance.get("gate_checks", {}),
        "governance_thresholds": governance.get("gate_thresholds", {}),
        "calibration_diagnostics": slippage_diag.get("calibration", {}),
        "stress_outcomes": stress,
        "metrics": metrics,
        "ui_inputs": {
            "governance_metadata": dict(governance_metadata or {}),
            "stress_controls": dict(stress_controls or {}),
        },
    }
    (run_dir / f"suite_bundle_{suite_key or 'custom'}.json").write_text(json.dumps(bundle, indent=2))

def _write_capacity_frontier_artifacts(*, run_dir: Path, robustness_report: dict[str, Any]) -> None:
    capacity = robustness_report.get("capacity_diagnostics", {}) if isinstance(robustness_report, dict) else {}
    frontier = capacity.get("capacity_frontier", []) if isinstance(capacity, dict) else []
    if not isinstance(frontier, list) or not frontier:
        return
    normalized = [row for row in frontier if isinstance(row, dict)]
    if not normalized:
        return
    (run_dir / "capacity_frontier.json").write_text(json.dumps(normalized, indent=2))
    fieldnames = sorted({key for row in normalized for key in row.keys()})
    with (run_dir / "capacity_frontier.csv").open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row in normalized:
            writer.writerow(row)

    max_abs_alpha = max(abs(float(row.get("expected_alpha_net_cost_bps", 0.0))) for row in normalized)
    scale = max(max_abs_alpha, 1e-9)
    frontier_summary: list[dict[str, float]] = []
    for row in normalized:
        alpha = float(row.get("expected_alpha_net_cost_bps", 0.0))
        sharpe = float(row.get("projected_post_cost_sharpe", 0.0))
        robustness_score = 0.6 * (alpha / scale) + 0.4 * sharpe
        frontier_summary.append(
            {
                "aum_scale": float(row.get("aum_scale", 0.0)),
                "expected_alpha_net_cost_bps": alpha,
                "projected_post_cost_sharpe": sharpe,
                "participation_rate": float(row.get("participation_rate", 0.0)),
                "robustness_score": robustness_score,
            }
        )
    frontier_summary.sort(key=lambda item: float(item["aum_scale"]))
    (run_dir / "robustness_frontier.json").write_text(
        json.dumps({"schema_version": "1.0", "frontier": frontier_summary}, indent=2)
    )
    with (run_dir / "robustness_frontier.csv").open("w", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "aum_scale",
                "expected_alpha_net_cost_bps",
                "projected_post_cost_sharpe",
                "participation_rate",
                "robustness_score",
            ],
        )
        writer.writeheader()
        writer.writerows(frontier_summary)

    chart_series = capacity.get("capacity_chart_series", {}) if isinstance(capacity, dict) else {}
    if isinstance(chart_series, dict) and chart_series:
        (run_dir / "capacity_frontier_series.json").write_text(json.dumps(chart_series, indent=2))
        series_rows: list[dict[str, object]] = []
        series_keys = [str(key) for key, value in chart_series.items() if isinstance(value, list)]
        length = max((len(chart_series[key]) for key in series_keys), default=0)
        for idx in range(length):
            row: dict[str, object] = {"index": idx}
            for key in series_keys:
                values = chart_series.get(key, [])
                row[key] = values[idx] if idx < len(values) else None
            series_rows.append(row)
        with (run_dir / "capacity_frontier_series.csv").open("w", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=["index", *series_keys])
            writer.writeheader()
            writer.writerows(series_rows)


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


def _collect_feature_hashes(parameters: dict[str, Any], data_snapshot: dict[str, Any]) -> dict[str, str]:
    feature_hashes: dict[str, str] = {}
    for key, value in sorted((parameters or {}).items(), key=lambda item: str(item[0])):
        name = str(key)
        if "feature" in name.lower() or name.lower().startswith("xsmom_"):
            feature_hashes[name] = _stable_fingerprint({"name": name, "value": value})
    if isinstance(data_snapshot, dict):
        data_fingerprint = data_snapshot.get("data_fingerprint")
        if isinstance(data_fingerprint, dict):
            for key, value in sorted(data_fingerprint.items(), key=lambda item: str(item[0])):
                feature_hashes[f"data_fingerprint.{key}"] = _stable_fingerprint({"name": key, "value": value})
    return feature_hashes


def _collect_artifact_inventory(run_dir: Path) -> dict[str, list[str]]:
    model_artifacts: list[str] = []
    plot_artifacts: list[str] = []
    metric_artifacts: list[str] = []
    for path in sorted(run_dir.glob("*")):
        if not path.is_file():
            continue
        suffix = path.suffix.lower()
        if suffix in {".png", ".svg", ".jpg", ".jpeg"}:
            plot_artifacts.append(path.name)
        if suffix in {".json", ".csv"} and "metric" in path.name.lower():
            metric_artifacts.append(path.name)
        if "model" in path.name.lower() or "leaderboard" in path.name.lower() or "fold_" in path.name.lower():
            model_artifacts.append(path.name)
    return {
        "model_artifacts": model_artifacts,
        "plot_artifacts": plot_artifacts,
        "metric_artifacts": metric_artifacts,
    }


def _stable_fingerprint(payload: dict[str, Any]) -> str:
    encoded = json.dumps(payload, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()




def _dependency_version(name: str) -> str:
    try:
        return importlib.metadata.version(name)
    except importlib.metadata.PackageNotFoundError:
        return "unknown"


def _collect_dependency_versions() -> dict[str, str]:
    deps = ["numpy", "pandas", "scipy"]
    return {name: _dependency_version(name) for name in deps}


def _compute_config_hash(parameters: dict[str, Any]) -> str:
    return _stable_fingerprint(parameters or {})


def _normalize_governance_for_fingerprint(governance: dict[str, Any]) -> dict[str, Any]:
    if not isinstance(governance, dict):
        return {}
    payload = dict(governance)
    payload.pop("audit_trail", None)
    return payload


def _compute_manifest_checksum(manifest: dict[str, Any]) -> str:
    return _stable_fingerprint(manifest)


def _build_metric_table_manifest(*, run_type: str, table_names: list[str]) -> dict[str, Any]:
    tables: list[dict[str, Any]] = []
    for name in table_names:
        schema = METRIC_TABLE_SCHEMA_VERSIONS.get(name, CANONICAL_METRIC_SCHEMA_VERSION)
        tables.append(
            {
                "table": name,
                "schema_version": schema,
                "compatibility": {
                    "minimum_reader_schema": schema,
                    "maximum_reader_schema": schema,
                },
            }
        )
    return {
        "run_type": run_type,
        "manifest_schema_version": RUN_MANIFEST_SCHEMA_VERSION,
        "tables": tables,
    }


def _write_metric_table_manifest(*, run_dir: Path, run_type: str, table_names: list[str]) -> dict[str, Any]:
    payload = _build_metric_table_manifest(run_type=run_type, table_names=table_names)
    (run_dir / "metric_tables_manifest.json").write_text(json.dumps(payload, indent=2))
    return payload


def _assert_metric_table_compatibility(run_dir: Path) -> dict[str, Any]:
    path = run_dir / "metric_tables_manifest.json"
    if not path.exists():
        return {}
    payload = json.loads(path.read_text())
    tables = payload.get("tables", []) if isinstance(payload, dict) else []
    for table in tables:
        if not isinstance(table, dict):
            continue
        schema = str(table.get("schema_version", ""))
        compat = table.get("compatibility", {}) if isinstance(table.get("compatibility"), dict) else {}
        min_schema = str(compat.get("minimum_reader_schema", schema))
        max_schema = str(compat.get("maximum_reader_schema", schema))
        if schema < min_schema or schema > max_schema:
            raise ValueError(f"Incompatible metric schema for table {table.get('table')}: {schema} not in [{min_schema}, {max_schema}]")
    return payload if isinstance(payload, dict) else {}


def _build_run_manifest(*, run_type: str, parameters: dict[str, Any], data_snapshot: dict[str, Any], random_seed: int | None, governance: dict[str, Any], lineage: dict[str, Any] | None = None, result_summary: dict[str, Any] | None = None, extra_fingerprint_payload: dict[str, Any] | None = None, metric_tables: dict[str, Any] | None = None) -> dict[str, Any]:
    code_commit = _resolve_git_commit()
    dependency_versions = _collect_dependency_versions()
    config_hash = _compute_config_hash(parameters)
    dataset_fingerprint_details = {
        "dataset_fingerprint": data_snapshot.get("dataset_fingerprint"),
        "data_fingerprint": dict(data_snapshot.get("data_fingerprint", {})) if isinstance(data_snapshot.get("data_fingerprint"), dict) else {},
        "coverage_by_symbol": dict(data_snapshot.get("coverage_by_symbol", {})) if isinstance(data_snapshot.get("coverage_by_symbol"), dict) else {},
        "missingness_by_symbol": dict(data_snapshot.get("missingness_by_symbol", {})) if isinstance(data_snapshot.get("missingness_by_symbol"), dict) else {},
        "range_start": data_snapshot.get("range_start"),
        "range_end": data_snapshot.get("range_end"),
        "symbols": list(data_snapshot.get("symbols", [])) if isinstance(data_snapshot.get("symbols"), list) else [],
        "timeframe": data_snapshot.get("timeframe"),
    }
    random_seeds = {
        "run_seed": random_seed,
        "python_random_seed": random_seed,
        "numpy_random_seed": random_seed,
        "subroutine_seeds": {
            "backtest_engine": random_seed,
            "portfolio_construction": random_seed,
            "regime_model": random_seed,
            "walk_forward_cv": random_seed,
            "optimizer_sampler": random_seed,
            "sweep_worker_base_seed": random_seed,
            "sweep_combo_seed_formula": "worker_seed if provided else run_seed + combo_index",
        },
    }
    reproducibility_metadata = {
        "feature_hashes": _collect_feature_hashes(parameters, data_snapshot),
        "environment_fingerprint": _stable_fingerprint(_collect_environment_metadata()),
    }
    fingerprint_payload = {
        "metric_schema_version": CANONICAL_METRIC_SCHEMA_VERSION,
        "manifest_schema_version": RUN_MANIFEST_SCHEMA_VERSION,
        "code_commit_hash": code_commit,
        "config_hash": config_hash,
        "parameters": parameters,
        "random_seeds": random_seeds,
        "data_snapshot_ids": data_snapshot,
        "dataset_fingerprint_details": dataset_fingerprint_details,
        "governance": _normalize_governance_for_fingerprint(governance),
        "reproducibility_metadata": reproducibility_metadata,
    }
    if extra_fingerprint_payload:
        fingerprint_payload.update(extra_fingerprint_payload)

    manifest = {
        "run_type": run_type,
        "manifest_schema_version": RUN_MANIFEST_SCHEMA_VERSION,
        "created_at": datetime.now().isoformat(),
        "code_commit_hash": code_commit,
        "code_version": code_commit,
        "metric_schema_version": CANONICAL_METRIC_SCHEMA_VERSION,
        "parameters": parameters,
        "config_hash": config_hash,
        "data_snapshot_ids": data_snapshot,
        "data_snapshot_identifiers": data_snapshot,
        "dataset_fingerprint_details": dataset_fingerprint_details,
        "random_seed": random_seed,
        "random_seeds": random_seeds,
        "dependency_versions": dependency_versions,
        "environment": _collect_environment_metadata(),
        "reproducibility_metadata": reproducibility_metadata,
        "governance": _normalize_governance_for_fingerprint(governance),
        "metric_tables": metric_tables or {},
        "reproducibility_fingerprint": _stable_fingerprint(fingerprint_payload),
    }
    manifest["run_id"] = _stable_fingerprint({
        "run_type": run_type,
        "created_at": manifest["created_at"],
        "config_hash": config_hash,
        "data_snapshot_ids": data_snapshot,
        "code_commit_hash": code_commit,
    })
    manifest["config_checksum"] = config_hash
    manifest["data_snapshot_checksum"] = _stable_fingerprint(data_snapshot)
    if lineage is not None:
        manifest["lineage"] = lineage
    if result_summary is not None:
        manifest["result_summary"] = result_summary
    manifest["manifest_checksum"] = _compute_manifest_checksum(manifest)
    return manifest


def _load_run_manifest(manifest_path: Path, *, strict: bool = True) -> dict[str, Any]:
    manifest = json.loads(manifest_path.read_text())
    if not isinstance(manifest, dict):
        raise ValueError("Manifest must be a JSON object")
    required = [
        "manifest_schema_version",
        "code_commit_hash",
        "config_hash",
        "data_snapshot_ids",
        "random_seeds",
        "dependency_versions",
        "parameters",
    ]
    for key in required:
        if key not in manifest:
            raise ValueError(f"Manifest missing required field: {key}")
    if strict:
        if str(manifest.get("manifest_schema_version")) != RUN_MANIFEST_SCHEMA_VERSION:
            raise ValueError("Manifest schema version is incompatible")
        if str(manifest.get("code_commit_hash")) != _resolve_git_commit():
            raise ValueError("Current git commit does not match manifest code commit hash")
        config_hash = _compute_config_hash(dict(manifest.get("parameters", {})))
        if config_hash != str(manifest.get("config_hash")):
            raise ValueError("Manifest config hash does not match serialized parameters")
        if str(manifest.get("config_checksum", "")) not in {"", config_hash}:
            raise ValueError("Manifest config checksum does not match serialized parameters")
        if "data_snapshot_checksum" in manifest:
            snapshot_checksum = _stable_fingerprint(dict(manifest.get("data_snapshot_ids", {})))
            if snapshot_checksum != str(manifest.get("data_snapshot_checksum", "")):
                raise ValueError("Manifest data snapshot checksum mismatch")
        if "manifest_checksum" in manifest:
            expected_checksum = str(manifest.get("manifest_checksum", ""))
            materialized = dict(manifest)
            materialized.pop("manifest_checksum", None)
            if expected_checksum != _compute_manifest_checksum(materialized):
                raise ValueError("Manifest checksum mismatch")
        dep_versions = manifest.get("dependency_versions", {})
        if isinstance(dep_versions, dict):
            current = _collect_dependency_versions()
            for dep, dep_version in dep_versions.items():
                if str(dep_version) != str(current.get(dep, "unknown")):
                    raise ValueError(f"Dependency version mismatch for {dep}: expected {dep_version}, got {current.get(dep)}")
    return manifest


def replay_manifest_run(*, manifest_path: Path, cache_root: Path | None = None, strict: bool = True) -> str:
    manifest = _load_run_manifest(manifest_path, strict=strict)
    run_type = str(manifest.get("run_type", ""))
    if run_type != "backtest":
        raise ValueError("Replay command currently supports backtest manifests only")
    params = dict(manifest.get("parameters", {}))
    tickers = [str(v) for v in params.get("tickers", [])]
    if not tickers:
        raise ValueError("Manifest parameters are missing tickers")
    output = run_time_series_momentum_backtest(
        tickers=tickers,
        start_date=date.fromisoformat(str(params["start_date"])),
        end_date=date.fromisoformat(str(params["end_date"])),
        cache_root=Path(cache_root or params.get("cache_root") or BACKTEST_CACHE_DIR),
        lookback_days=int(params.get("lookback_days", 90)),
        skip_days=int(params.get("skip_days", 5)),
        costs_bps=float(params.get("costs_bps", 5.0)),
        execution_model=str(params.get("execution_model", "bps")),
        execution_model_params=dict(params.get("execution_model_params", {})),
        carry_model=str(params.get("carry_model", "short_borrow")),
        carry_model_params=dict(params.get("carry_model_params", {})),
        entry_signal=str(params.get("entry_signal", "ts_momentum")),
        entry_signal_params=dict(params.get("entry_signal_params", {})),
        exit_signal=str(params.get("exit_signal", "none")),
        exit_signal_params=dict(params.get("exit_signal_params", {})),
        signal_rebalance_interval=int(params.get("signal_rebalance_interval", 1)),
        starting_capital=float(params.get("starting_capital", 100_000.0)),
        bet_sizing_mode=str(params.get("bet_sizing_mode", "half_kelly")),
        custom_bet_pct=float(params.get("custom_bet_pct", 10.0)),
        strategy=str(params.get("strategy", "momentum")),
        xsmom_top_quantile=float(params.get("xsmom_top_quantile", 0.2)),
        xsmom_bottom_quantile=float(params.get("xsmom_bottom_quantile", 0.2)),
        xsmom_long_only=bool(params.get("xsmom_long_only", False)),
        xsmom_vol_lookback_days=int(params.get("xsmom_vol_lookback_days", 20)),
        timeframe=str(params.get("timeframe", "1m")),
        portfolio_method=str(params.get("portfolio_method", "equal_weight")),
        portfolio_vol_lookback_bars=int(params.get("portfolio_vol_lookback_bars", 20)),
        portfolio_target_volatility=float(params.get("portfolio_target_volatility", 0.10)),
        portfolio_max_symbol_weight=float(params.get("portfolio_max_symbol_weight", 0.25)),
        portfolio_max_sector_weight=float(params.get("portfolio_max_sector_weight", 0.60)),
        portfolio_rebalance_frequency_bars=int(params.get("portfolio_rebalance_frequency_bars", 1)),
        portfolio_clustering_linkage=str(params.get("portfolio_clustering_linkage", "single")),
        portfolio_covariance_shrinkage=float(params.get("portfolio_covariance_shrinkage", 0.15)),
        portfolio_max_gross_exposure=float(params.get("portfolio_max_gross_exposure", 1.0)),
        portfolio_min_net_exposure=float(params.get("portfolio_min_net_exposure", -1.0)),
        portfolio_max_net_exposure=float(params.get("portfolio_max_net_exposure", 1.0)),
        portfolio_max_net_gamma=float(params.get("portfolio_max_net_gamma")) if params.get("portfolio_max_net_gamma") not in (None, "") else None,
        portfolio_max_abs_vega_bucket=float(params.get("portfolio_max_abs_vega_bucket")) if params.get("portfolio_max_abs_vega_bucket") not in (None, "") else None,
        portfolio_max_abs_delta_per_underlying=float(params.get("portfolio_max_abs_delta_per_underlying")) if params.get("portfolio_max_abs_delta_per_underlying") not in (None, "") else None,
    )
    run_dir = _extract_saved_output_dir(output)
    _assert_metric_table_compatibility(run_dir)
    replay_manifest = json.loads((run_dir / "manifest.json").read_text())
    expected_fingerprint = str(manifest.get("reproducibility_fingerprint", ""))
    if strict and expected_fingerprint and str(replay_manifest.get("reproducibility_fingerprint", "")) != expected_fingerprint:
        raise ValueError("Replay reproducibility fingerprint mismatch")
    return output

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
        "data_fingerprint": dict(getattr(arrays.metadata, "data_fingerprint", {}) or {}),
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
        "data_fingerprint": dict(parameters.get("data_fingerprint", {})) if isinstance(parameters.get("data_fingerprint"), dict) else {},
    }
    payload["dataset_fingerprint"] = _stable_fingerprint(payload)
    return payload


def _build_report_run_details(run_dir: Path) -> dict[str, Any]:
    manifest_path = run_dir / "manifest.json"
    if not manifest_path.exists():
        return {}
    try:
        manifest = json.loads(manifest_path.read_text())
    except json.JSONDecodeError:
        return {}
    if not isinstance(manifest, dict):
        return {}
    return {
        "code_commit_hash": manifest.get("code_commit_hash", manifest.get("code_version")),
        "code_version": manifest.get("code_version"),
        "config_hash": manifest.get("config_hash"),
        "dataset_fingerprint_details": manifest.get("dataset_fingerprint_details", {}),
        "random_seeds": manifest.get("random_seeds", {}),
    }



def _build_governance_metadata(raw: dict[str, Any] | None) -> dict[str, Any]:
    source = dict(raw or {})
    promotion_state = str(source.get("promotion_state", "research")).strip().lower()
    if promotion_state not in PROMOTION_STATES:
        promotion_state = "research"

    required_checks = list(PROMOTION_REQUIRED_CHECKS.get(promotion_state, []))
    checks = source.get("checks") if isinstance(source.get("checks"), dict) else {}

    def _float(name: str, default: float) -> float:
        value = source.get(name, default)
        try:
            return float(value)
        except (TypeError, ValueError):
            return float(default)

    def _int(name: str, default: int) -> int:
        value = source.get(name, default)
        try:
            return int(float(value))
        except (TypeError, ValueError):
            return int(default)

    gate_thresholds = {
        "min_oos_periods": _int("min_oos_periods", 3),
        "min_stability_score": _float("min_stability_score", 0.55),
        "max_turnover_total": _float("max_turnover_total", 4.0),
        "min_capacity_score": _float("min_capacity_score", 0.5),
        "max_signal_agreement_drift": _float("max_signal_agreement_drift", 0.10),
        "max_fill_slippage_drift_bps": _float("max_fill_slippage_drift_bps", 5.0),
        "max_pnl_attribution_divergence": _float("max_pnl_attribution_divergence", 0.15),
        "min_friction_adjusted_edge": _float("min_friction_adjusted_edge", 0.0),
        "min_causal_effect_tstat": _float("min_causal_effect_tstat", 1.96),
        "max_pretrend_pvalue": _float("max_pretrend_pvalue", 0.10),
        "min_placebo_pvalue": _float("min_placebo_pvalue", 0.05),
        "max_relative_attenuation": _float("max_relative_attenuation", 0.50),
        "max_feature_mean_shift": _float("max_feature_mean_shift", 0.15),
        "max_feature_std_ratio_shift": _float("max_feature_std_ratio_shift", 0.35),
        "max_feature_psi": _float("max_feature_psi", 0.20),
        "max_label_psi": _float("max_label_psi", 0.20),
        "max_label_kld": _float("max_label_kld", 0.12),
        "max_label_error_rate_drift": _float("max_label_error_rate_drift", 0.03),
        "max_residual_mean_shift": _float("max_residual_mean_shift", 0.10),
        "max_residual_std_ratio_shift": _float("max_residual_std_ratio_shift", 0.25),
        "max_residual_autocorr": _float("max_residual_autocorr", 0.20),
        "max_false_alarm_rate": _float("max_false_alarm_rate", 0.25),
        "min_deflated_sharpe_ratio": _float("min_deflated_sharpe_ratio", 0.10),
        "max_reality_check_pvalue": _float("max_reality_check_pvalue", 0.10),
        "max_parameter_stability_penalty": _float("max_parameter_stability_penalty", 0.40),
        "max_train_validation_drift": _float("max_train_validation_drift", 0.30),
        "max_validation_test_drift": _float("max_validation_test_drift", 0.30),
        "max_train_test_drift": _float("max_train_test_drift", 0.45),
        "retrain_on": source.get("retrain_on", ["high"]),
        "recalibrate_on": source.get("recalibrate_on", ["medium"]),
        "alert_routing": source.get(
            "alert_routing",
            {
                "low": ["dashboard"],
                "medium": ["dashboard", "slack:#ml-monitoring"],
                "high": ["dashboard", "slack:#incident-ml", "pagerduty:ml-oncall"],
            },
        ),
        "feature_thresholds": source.get("feature_thresholds", {}),
    }

    drift_monitoring = evaluate_drift_monitoring(
        expected=source.get("expected_outcomes") if isinstance(source.get("expected_outcomes"), dict) else {},
        observed=source.get("observed_outcomes") if isinstance(source.get("observed_outcomes"), dict) else {},
        thresholds=gate_thresholds,
    )
    causal_robustness = _evaluate_causal_robustness(
        causal_validation=source.get("causal_validation") if isinstance(source.get("causal_validation"), dict) else {},
        thresholds=gate_thresholds,
    )

    experiment_id = str(source.get("experiment_id", "")).strip()

    gate_checks = {
        "dataset_lock": bool(source.get("dataset_snapshot_lock", "").strip()),
        "oos_periods": bool(checks.get("oos_periods", False)),
        "stability_threshold": bool(checks.get("stability_threshold", False)),
        "turnover_capacity": bool(checks.get("turnover_capacity", False)),
        "approval": bool(checks.get("approval", False)),
        "signal_diagnostics": bool(checks.get("signal_diagnostics", False)),
        "drift_monitoring": bool(drift_monitoring.get("within_tolerance", False)),
        "friction_adjusted_edge": bool(checks.get("friction_adjusted_edge", False)),
        "causal_robustness": bool(causal_robustness.get("pass", False)),
        "deflated_sharpe_reality_check": bool(checks.get("deflated_sharpe_reality_check", False)),
        "parameter_stability_penalty": bool(checks.get("parameter_stability_penalty", False)),
        "train_validation_test_drift": bool(checks.get("train_validation_test_drift", False)),
        "experiment_id": bool(experiment_id),
    }

    missing_required = [name for name in required_checks if not gate_checks.get(name, False)]

    comments = source.get("comments") if isinstance(source.get("comments"), list) else []
    review_actions = source.get("review_actions") if isinstance(source.get("review_actions"), list) else []
    decision_log = source.get("decision_log") if isinstance(source.get("decision_log"), list) else []

    normalized_comments = [
        {
            "owner": str(row.get("owner", "research_lab_ui")),
            "note": str(row.get("note", "")),
            "timestamp": str(row.get("timestamp", datetime.now().isoformat())),
        }
        for row in comments
        if isinstance(row, dict) and str(row.get("note", "")).strip()
    ]
    normalized_review_actions = [
        {
            "owner": str(row.get("owner", "research_lab_ui")),
            "action": str(row.get("action", "")),
            "status": str(row.get("status", "recorded")),
            "timestamp": str(row.get("timestamp", datetime.now().isoformat())),
        }
        for row in review_actions
        if isinstance(row, dict) and str(row.get("action", "")).strip()
    ]
    normalized_decision_log = [
        {
            "owner": str(row.get("owner", "research_lab_ui")),
            "decision": str(row.get("decision", "pending")),
            "reason": str(row.get("reason", "")),
            "timestamp": str(row.get("timestamp", datetime.now().isoformat())),
        }
        for row in decision_log
        if isinstance(row, dict)
    ]

    governance = {
        "hypothesis_id": str(source.get("hypothesis_id", "")).strip(),
        "experiment_id": experiment_id,
        "owner": str(source.get("owner", "")).strip(),
        "dataset_snapshot_lock": str(source.get("dataset_snapshot_lock", "")).strip(),
        "acceptance_criteria": str(source.get("acceptance_criteria", "")).strip(),
        "approval_status": str(source.get("approval_status", "pending")).strip() or "pending",
        "promotion_state": promotion_state,
        "promotion_required_checks": required_checks,
        "gate_thresholds": gate_thresholds,
        "gate_checks": gate_checks,
        "drift_monitoring": drift_monitoring,
        "causal_robustness": causal_robustness,
        "missing_required_checks": missing_required,
        "is_promotion_ready": not missing_required,
        "comments": normalized_comments,
        "review_actions": normalized_review_actions,
        "decision_log": normalized_decision_log,
        "audit_trail": [
            {
                "timestamp": datetime.now().isoformat(),
                "event": "governance_metadata_emitted",
                "promotion_state": promotion_state,
                "approval_status": str(source.get("approval_status", "pending")).strip() or "pending",
                "required_checks": required_checks,
                "missing_required_checks": missing_required,
            }
        ],
    }
    return governance


def _evaluate_governance_gate_checks(
    *,
    metrics: dict[str, float],
    fold_rows: list[dict[str, Any]] | None,
    governance: dict[str, Any],
) -> dict[str, bool]:
    thresholds = governance.get("gate_thresholds", {}) if isinstance(governance.get("gate_thresholds"), dict) else {}
    min_oos = int(thresholds.get("min_oos_periods", 3))
    min_stability = float(thresholds.get("min_stability_score", 0.55))
    max_turnover = float(thresholds.get("max_turnover_total", 4.0))
    min_capacity = float(thresholds.get("min_capacity_score", 0.5))

    oos_periods = len(fold_rows) if isinstance(fold_rows, list) else 0
    sharpe = float(metrics.get("sharpe", 0.0))
    rolling_sharpe_mean = float(metrics.get("rolling_sharpe_mean", sharpe))
    stability_score = max(0.0, min(1.0, 0.5 + 0.25 * sharpe + 0.25 * rolling_sharpe_mean))
    turnover_total = float(metrics.get("turnover_total", 0.0))
    capacity_score = max(0.0, 1.0 - (turnover_total / max(max_turnover, 1e-9)))

    signal_diag_ready = bool(metrics.get("signal_diagnostics_ready", False))
    drift_monitoring = governance.get("drift_monitoring", {}) if isinstance(governance.get("drift_monitoring"), dict) else {}
    drift_within_tolerance = bool(drift_monitoring.get("within_tolerance", False))
    causal_robustness = governance.get("causal_robustness", {}) if isinstance(governance.get("causal_robustness"), dict) else {}
    diagnostics = _compute_governance_diagnostics(metrics=metrics, fold_rows=fold_rows, thresholds=thresholds)
    governance["governance_diagnostics"] = diagnostics
    experiment_id = str(governance.get("experiment_id", "")).strip()

    checks = {
        "dataset_lock": bool(governance.get("dataset_snapshot_lock")),
        "signal_diagnostics": signal_diag_ready,
        "oos_periods": oos_periods >= max(1, min_oos),
        "stability_threshold": stability_score >= min_stability,
        "turnover_capacity": turnover_total <= max_turnover and capacity_score >= min_capacity,
        "approval": str(governance.get("approval_status", "pending")).lower() in {"approved", "waived"},
        "drift_monitoring": drift_within_tolerance,
        "friction_adjusted_edge": float(metrics.get("friction_adjusted_edge", float("-inf"))) >= float(thresholds.get("min_friction_adjusted_edge", 0.0)),
        "causal_robustness": bool(causal_robustness.get("pass", False)),
        "deflated_sharpe_reality_check": bool(diagnostics.get("deflated_sharpe_reality_check", {}).get("pass", False)),
        "parameter_stability_penalty": bool(diagnostics.get("parameter_stability", {}).get("pass", False)),
        "train_validation_test_drift": bool(diagnostics.get("train_validation_test_drift", {}).get("pass", False)),
        "experiment_id": bool(experiment_id),
    }
    return checks


def _compute_governance_diagnostics(
    *,
    metrics: dict[str, float],
    fold_rows: list[dict[str, Any]] | None,
    thresholds: dict[str, Any],
) -> dict[str, Any]:
    deflated = float(metrics.get("deflated_sharpe_ratio", float("nan")))
    if not np.isfinite(deflated):
        deflated = float(metrics.get("probabilistic_sharpe_ratio", float("nan")))
    white_pvalue = float(metrics.get("white_reality_check_pvalue", metrics.get("corrected_pvalue", 1.0)))
    spa_pvalue = float(metrics.get("spa_pvalue", white_pvalue))
    combined_reality_pvalue = max(white_pvalue, spa_pvalue)
    min_deflated = float(thresholds.get("min_deflated_sharpe_ratio", 0.10))
    max_reality = float(thresholds.get("max_reality_check_pvalue", 0.10))
    dsr_check = {
        "deflated_sharpe_ratio": deflated if np.isfinite(deflated) else None,
        "white_reality_check_pvalue": white_pvalue,
        "spa_pvalue": spa_pvalue,
        "combined_reality_check_pvalue": combined_reality_pvalue,
        "min_deflated_sharpe_ratio": min_deflated,
        "max_reality_check_pvalue": max_reality,
        "pass": bool(np.isfinite(deflated) and deflated >= min_deflated and combined_reality_pvalue <= max_reality),
    }

    fold_list = fold_rows if isinstance(fold_rows, list) else []
    selected = [json.dumps(row.get("selected_params", {}), sort_keys=True) for row in fold_list if isinstance(row, dict)]
    unique_selected = len(set(selected))
    fold_count = len(fold_list)
    churn_ratio = float(unique_selected / max(1, fold_count))
    validation_scores = np.asarray([float(row.get("validation_score", 0.0)) for row in fold_list if isinstance(row, dict)], dtype=float)
    validation_std = float(np.std(validation_scores, ddof=0)) if validation_scores.size else 0.0
    validation_mean = float(np.mean(np.abs(validation_scores))) if validation_scores.size else 0.0
    normalized_variability = validation_std / max(1e-9, 0.25 + validation_mean)
    stability_penalty = max(0.0, min(1.0, 0.6 * churn_ratio + 0.4 * min(1.0, normalized_variability)))
    max_penalty = float(thresholds.get("max_parameter_stability_penalty", 0.40))
    stability_diag = {
        "fold_count": fold_count,
        "unique_selected_params": unique_selected,
        "churn_ratio": churn_ratio,
        "validation_score_std": validation_std,
        "normalized_validation_variability": normalized_variability,
        "parameter_stability_penalty": stability_penalty,
        "max_parameter_stability_penalty": max_penalty,
        "pass": bool(fold_count > 0 and stability_penalty <= max_penalty),
    }

    train_scores: list[float] = []
    validation_scores_ladder: list[float] = []
    test_scores: list[float] = []
    for row in fold_list:
        if not isinstance(row, dict):
            continue
        diagnostics_rows = row.get("diagnostics", [])
        if not isinstance(diagnostics_rows, list):
            continue
        best_diag = max(
            [item for item in diagnostics_rows if isinstance(item, dict)],
            key=lambda item: float(item.get("validation_score", float("-inf"))),
            default=None,
        )
        if not isinstance(best_diag, dict):
            continue
        train_metric = float(best_diag.get("train_metrics", {}).get("sharpe", 0.0))
        val_metric = float(best_diag.get("validation_metrics", {}).get("sharpe", 0.0))
        test_metric = float((row.get("oos_metrics", {}) if isinstance(row.get("oos_metrics"), dict) else {}).get("sharpe", 0.0))
        train_scores.append(train_metric)
        validation_scores_ladder.append(val_metric)
        test_scores.append(test_metric)

    train_arr = np.asarray(train_scores, dtype=float)
    val_arr = np.asarray(validation_scores_ladder, dtype=float)
    test_arr = np.asarray(test_scores, dtype=float)
    tv_drift = float(np.mean(np.abs(train_arr - val_arr))) if train_arr.size and val_arr.size else float("inf")
    vt_drift = float(np.mean(np.abs(val_arr - test_arr))) if val_arr.size and test_arr.size else float("inf")
    tt_drift = float(np.mean(np.abs(train_arr - test_arr))) if train_arr.size and test_arr.size else float("inf")
    max_tv = float(thresholds.get("max_train_validation_drift", 0.30))
    max_vt = float(thresholds.get("max_validation_test_drift", 0.30))
    max_tt = float(thresholds.get("max_train_test_drift", 0.45))
    drift_diag = {
        "folds_evaluated": int(min(train_arr.size, val_arr.size, test_arr.size)),
        "train_validation_abs_drift": tv_drift,
        "validation_test_abs_drift": vt_drift,
        "train_test_abs_drift": tt_drift,
        "max_train_validation_drift": max_tv,
        "max_validation_test_drift": max_vt,
        "max_train_test_drift": max_tt,
        "pass": bool(np.isfinite(tv_drift) and np.isfinite(vt_drift) and np.isfinite(tt_drift) and tv_drift <= max_tv and vt_drift <= max_vt and tt_drift <= max_tt),
    }

    return {
        "deflated_sharpe_reality_check": dsr_check,
        "parameter_stability": stability_diag,
        "train_validation_test_drift": drift_diag,
    }


def _evaluate_causal_robustness(*, causal_validation: dict[str, Any], thresholds: dict[str, float]) -> dict[str, Any]:
    methods = causal_validation.get("methods") if isinstance(causal_validation.get("methods"), list) else []
    method_results: list[dict[str, Any]] = []
    for entry in methods:
        if not isinstance(entry, dict):
            continue
        method = str(entry.get("method", "unknown")).strip().lower()
        effect_tstat = float(entry.get("effect_tstat", 0.0))
        pretrend_pvalue = float(entry.get("pretrend_pvalue", 0.0))
        placebo_pvalue = float(entry.get("placebo_pvalue", 0.0))
        relative_attenuation = float(entry.get("relative_attenuation", 1.0))
        passed = (
            effect_tstat >= float(thresholds.get("min_causal_effect_tstat", 1.96))
            and pretrend_pvalue >= float(thresholds.get("max_pretrend_pvalue", 0.10))
            and placebo_pvalue >= float(thresholds.get("min_placebo_pvalue", 0.05))
            and relative_attenuation <= float(thresholds.get("max_relative_attenuation", 0.50))
        )
        method_results.append(
            {
                "method": method,
                "effect_tstat": effect_tstat,
                "pretrend_pvalue": pretrend_pvalue,
                "placebo_pvalue": placebo_pvalue,
                "relative_attenuation": relative_attenuation,
                "pass": bool(passed),
            }
        )

    required_methods = {"difference_in_differences", "synthetic_control", "propensity_score_matching"}
    available = {str(row.get("method", "")).lower() for row in method_results}
    missing_methods = sorted(required_methods - available)
    overall_pass = bool(method_results) and not missing_methods and all(bool(row.get("pass", False)) for row in method_results)
    return {
        "required_methods": sorted(required_methods),
        "missing_methods": missing_methods,
        "method_results": method_results,
        "pass": overall_pass,
    }

def _append_experiment_index(entry: dict[str, Any]) -> None:
    append_experiment_entry(BACKTEST_OUTPUT_DIR, entry)




def _timestamp_to_datetime(value: object) -> datetime:
    if isinstance(value, datetime):
        return value

    if isinstance(value, np.datetime64):
        nanos = value.astype("datetime64[ns]").astype(np.int64)
        return datetime.utcfromtimestamp(float(nanos) / 1_000_000_000.0)

    if isinstance(value, (np.integer, int, np.floating, float)):
        numeric = float(value)
        abs_numeric = abs(numeric)
        if abs_numeric >= 1e17:
            seconds = numeric / 1_000_000_000.0
        elif abs_numeric >= 1e14:
            seconds = numeric / 1_000_000.0
        elif abs_numeric >= 1e11:
            seconds = numeric / 1_000.0
        else:
            seconds = numeric
        return datetime.utcfromtimestamp(seconds)

    if hasattr(value, "to_pydatetime"):
        try:
            dt = value.to_pydatetime()
            if isinstance(dt, datetime):
                return dt
        except Exception:
            pass

    return datetime.fromisoformat(str(value))


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
    margin_utilization = np.asarray(diagnostics.get("margin_utilization", np.zeros(len(timestamps))), dtype=float)
    cash = np.asarray(diagnostics.get("cash", np.zeros(len(timestamps))), dtype=float)
    margin_requirement = np.asarray(diagnostics.get("margin_requirement", np.zeros(len(timestamps))), dtype=float)
    excess_liquidity = np.asarray(diagnostics.get("excess_liquidity", np.zeros(len(timestamps))), dtype=float)
    buying_power = np.asarray(diagnostics.get("buying_power", np.zeros(len(timestamps))), dtype=float)
    forced_liquidation = np.asarray(diagnostics.get("forced_liquidation", np.zeros(len(timestamps))), dtype=float)
    deleveraging_scale = np.asarray(diagnostics.get("deleveraging_scale", np.ones(len(timestamps))), dtype=float)
    net_gamma = np.asarray(diagnostics.get("net_gamma_exposure", np.zeros(len(timestamps))), dtype=float)
    max_vega_bucket = np.asarray(diagnostics.get("max_vega_bucket_exposure", np.zeros(len(timestamps))), dtype=float)
    max_underlying_delta = np.asarray(diagnostics.get("max_underlying_delta_exposure", np.zeros(len(timestamps))), dtype=float)
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
            "margin_utilization": float(margin_utilization[idx]) if idx < margin_utilization.size else 0.0,
            "cash": float(cash[idx]) if idx < cash.size else 0.0,
            "margin_requirement": float(margin_requirement[idx]) if idx < margin_requirement.size else 0.0,
            "excess_liquidity": float(excess_liquidity[idx]) if idx < excess_liquidity.size else 0.0,
            "buying_power": float(buying_power[idx]) if idx < buying_power.size else 0.0,
            "forced_liquidation": float(forced_liquidation[idx]) if idx < forced_liquidation.size else 0.0,
            "deleveraging_scale": float(deleveraging_scale[idx]) if idx < deleveraging_scale.size else 1.0,
            "net_gamma_exposure": float(net_gamma[idx]) if idx < net_gamma.size else 0.0,
            "max_vega_bucket_exposure": float(max_vega_bucket[idx]) if idx < max_vega_bucket.size else 0.0,
            "max_underlying_delta_exposure": float(max_underlying_delta[idx]) if idx < max_underlying_delta.size else 0.0,
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
                "margin_utilization",
                "cash",
                "margin_requirement",
                "excess_liquidity",
                "buying_power",
                "forced_liquidation",
                "deleveraging_scale",
                "net_gamma_exposure",
                "max_vega_bucket_exposure",
                "max_underlying_delta_exposure",
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


def _write_risk_monitoring_artifacts(
    *,
    run_dir: Path,
    timestamps: list[str],
    symbol_order: list[str],
    trades: np.ndarray,
    returns: np.ndarray,
    diagnostics: dict[str, np.ndarray],
    regime_labels: np.ndarray | None = None,
    regime_probabilities: np.ndarray | None = None,
    regime_states: np.ndarray | None = None,
) -> None:
    periods = len(timestamps)
    gross = np.asarray(diagnostics.get("gross_exposure", np.zeros(periods)), dtype=float)
    concentration = np.asarray(diagnostics.get("concentration", np.zeros(periods)), dtype=float)
    turnover = np.asarray(diagnostics.get("turnover", np.zeros(periods)), dtype=float)
    leverage = np.asarray(diagnostics.get("leverage_usage", np.zeros(periods)), dtype=float)
    excess_liquidity = np.asarray(diagnostics.get("excess_liquidity", np.zeros(periods)), dtype=float)

    returns_arr = np.asarray(returns, dtype=float)
    tail_cut = float(np.quantile(returns_arr, 0.05)) if returns_arr.size else 0.0
    var_95 = abs(min(0.0, tail_cut))
    cvar_samples = returns_arr[returns_arr <= tail_cut]
    cvar_95 = abs(float(np.mean(cvar_samples))) if cvar_samples.size else var_95

    running_equity = np.cumprod(1.0 + returns_arr) if returns_arr.size else np.array([], dtype=float)
    safe_peak = np.maximum.accumulate(np.where(running_equity == 0.0, 1.0, running_equity)) if running_equity.size else np.array([], dtype=float)
    drawdown = running_equity / np.where(safe_peak == 0.0, 1.0, safe_peak) - 1.0 if running_equity.size else np.array([], dtype=float)

    row_corr = np.corrcoef(np.nan_to_num(np.asarray(trades, dtype=float), nan=0.0), rowvar=True) if periods > 1 else np.eye(periods)
    corr_spikes: np.ndarray
    if row_corr.ndim == 2 and row_corr.shape[0] == periods:
        abs_corr = np.abs(row_corr)
        np.fill_diagonal(abs_corr, 0.0)
        corr_spikes = np.max(abs_corr, axis=1)
    else:
        corr_spikes = np.zeros(periods, dtype=float)

    regime_arr = np.asarray(regime_labels, dtype=object) if regime_labels is not None else np.array([], dtype=object)
    probs = np.asarray(regime_probabilities, dtype=float) if regime_probabilities is not None else np.zeros((periods, 0), dtype=float)
    states = np.asarray(regime_states, dtype=object) if regime_states is not None else np.array([], dtype=object)
    valid_probs = probs.ndim == 2 and probs.shape[0] == periods and states.size == probs.shape[1]

    threshold = {
        "gross_exposure": 1.2,
        "concentration": 0.35,
        "var_95": 0.03,
        "cvar_95": 0.05,
        "drawdown": -0.10,
        "correlation_spike": 0.85,
        "liquidity_risk": 0.75,
        "model_confidence": 0.55,
    }

    dashboard_rows: list[dict[str, object]] = []
    intervention_rows: list[dict[str, object]] = []
    open_positions = np.cumsum(np.nan_to_num(np.asarray(trades, dtype=float), nan=0.0), axis=0) if periods else np.zeros((0, len(symbol_order)))
    for idx, ts in enumerate(timestamps):
        regime_state = str(regime_arr[idx]) if idx < regime_arr.size else "unknown"
        confidence = float(np.max(probs[idx])) if valid_probs else 0.5
        liq_proxy = min(2.0, float(turnover[idx]) * 4.0 + float(leverage[idx])) if idx < turnover.size and idx < leverage.size else 0.0
        if idx < excess_liquidity.size and excess_liquidity[idx] < 0.0:
            liq_proxy = max(liq_proxy, 1.0)
        row = {
            "timestamp": ts,
            "gross_exposure": float(gross[idx]) if idx < gross.size else 0.0,
            "concentration": float(concentration[idx]) if idx < concentration.size else 0.0,
            "var_95": var_95,
            "cvar_95": cvar_95,
            "drawdown": float(drawdown[idx]) if idx < drawdown.size else 0.0,
            "correlation_spike": float(corr_spikes[idx]) if idx < corr_spikes.size else 0.0,
            "liquidity_risk": liq_proxy,
            "model_confidence": confidence,
            "regime_state": regime_state,
        }
        dashboard_rows.append(row)

        position_row = open_positions[idx] if idx < open_positions.shape[0] else np.zeros(len(symbol_order), dtype=float)
        top_symbol = ""
        top_notional = 0.0
        if position_row.size:
            best = int(np.argmax(np.abs(position_row)))
            top_symbol = symbol_order[best]
            top_notional = float(position_row[best])

        checks = {
            "gross_exposure": row["gross_exposure"] > threshold["gross_exposure"],
            "concentration": row["concentration"] > threshold["concentration"],
            "drawdown": row["drawdown"] < threshold["drawdown"],
            "correlation_spike": row["correlation_spike"] > threshold["correlation_spike"],
            "liquidity_risk": row["liquidity_risk"] > threshold["liquidity_risk"],
            "model_confidence": row["model_confidence"] < threshold["model_confidence"],
            "tail_risk": (row["var_95"] > threshold["var_95"]) or (row["cvar_95"] > threshold["cvar_95"]),
        }
        triggered = [name for name, active in checks.items() if active]
        if triggered:
            action = "kill_switch" if any(name in {"drawdown", "liquidity_risk", "tail_risk"} for name in triggered) else "alert"
            intervention_rows.append(
                {
                    "timestamp": ts,
                    "action": action,
                    "triggers": ",".join(triggered),
                    "reason": f"threshold breach: {', '.join(triggered)}",
                    "regime_state": regime_state,
                    "model_confidence": confidence,
                    "top_position": top_symbol,
                    "top_position_size": top_notional,
                }
            )

    (run_dir / "risk_dashboard.json").write_text(json.dumps(dashboard_rows, indent=2))
    with (run_dir / "risk_dashboard.csv").open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(dashboard_rows[0].keys()) if dashboard_rows else ["timestamp"])
        writer.writeheader()
        writer.writerows(dashboard_rows)

    (run_dir / "risk_interventions.json").write_text(json.dumps(intervention_rows, indent=2))
    with (run_dir / "risk_interventions.csv").open("w", newline="") as handle:
        cols = list(intervention_rows[0].keys()) if intervention_rows else ["timestamp", "action", "triggers", "reason"]
        writer = csv.DictWriter(handle, fieldnames=cols)
        writer.writeheader()
        writer.writerows(intervention_rows)


def _write_regime_labels_csv_json(
    *,
    run_dir: Path,
    timestamps: list[str],
    regime_labels: np.ndarray,
    regime_probabilities: np.ndarray | None = None,
    regime_states: np.ndarray | None = None,
) -> None:
    rows = []
    states = np.asarray(regime_states, dtype=object) if regime_states is not None else np.array([], dtype=object)
    probs = np.asarray(regime_probabilities, dtype=float) if regime_probabilities is not None else np.zeros((len(timestamps), 0), dtype=float)
    valid_probs = probs.ndim == 2 and probs.shape[0] == len(timestamps) and states.size == probs.shape[1]

    for idx, ts in enumerate(timestamps):
        label = ""
        if idx < regime_labels.size:
            label = str(regime_labels[idx])
        row: dict[str, object] = {"timestamp": ts, "regime": label}
        if valid_probs:
            for col, state in enumerate(states.tolist()):
                row[f"prob_{state}"] = float(probs[idx, col])
        rows.append(row)

    csv_path = run_dir / "regimes.csv"
    with csv_path.open("w", newline="") as handle:
        fieldnames = ["timestamp", "regime"]
        if valid_probs:
            fieldnames.extend([f"prob_{state}" for state in states.tolist()])
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    (run_dir / "regimes.json").write_text(json.dumps(rows, indent=2))


def _write_regime_pnl_attribution(*, run_dir: Path, rows: list[dict[str, float | str | int]]) -> None:
    csv_path = run_dir / "regime_pnl_attribution.csv"
    with csv_path.open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["regime", "bars", "pnl_total", "pnl_mean"])
        writer.writeheader()
        writer.writerows(rows)
    (run_dir / "regime_pnl_attribution.json").write_text(json.dumps(rows, indent=2))


def _friction_adjusted_edge_checks(
    *,
    friction_on_metrics: dict[str, float],
    friction_off_metrics: dict[str, float],
    governance_payload: dict[str, Any],
) -> dict[str, Any]:
    thresholds = governance_payload.get("gate_thresholds", {}) if isinstance(governance_payload.get("gate_thresholds"), dict) else {}
    min_edge = float(thresholds.get("min_friction_adjusted_edge", 0.0))
    on_sharpe = float(friction_on_metrics.get("sharpe", 0.0))
    off_sharpe = float(friction_off_metrics.get("sharpe", 0.0))
    edge = on_sharpe
    edge_retention = 0.0 if abs(off_sharpe) < 1e-12 else edge / off_sharpe
    passed = edge >= min_edge
    return {
        "friction_on_sharpe": on_sharpe,
        "friction_off_sharpe": off_sharpe,
        "friction_adjusted_edge": edge,
        "friction_edge_retention": float(edge_retention),
        "min_friction_adjusted_edge": min_edge,
        "pass": bool(passed),
    }



def _build_slippage_model(
    *,
    model_name: str,
    costs_bps: float,
    params: dict[str, object] | None,
    as_of_date: str | None = None,
) -> tuple[object, SlippageCalibrationSelection]:
    cfg = dict(params or {})
    name = str(model_name).strip().lower()
    default_selection = SlippageCalibrationSelection(
        params=dict(cfg),
        source="config",
        effective_date=None,
        warning_flags=[],
    )
    if name == "bps":
        return BpsSlippage(float(cfg.get("bps", costs_bps))), default_selection
    if name == "spread":
        return SpreadSlippage(float(cfg.get("spread_bps", costs_bps))), default_selection
    if name == "participation":
        snapshots_path = cfg.get("snapshot_path")
        if snapshots_path is not None:
            payload = load_slippage_calibration_snapshots(str(snapshots_path))
            selected = select_slippage_calibration_snapshot(
                payload,
                as_of_date=as_of_date,
                default_params={
                    "base_bps": float(cfg.get("base_bps", 0.0)),
                    "impact_coefficient_bps": float(cfg.get("impact_coefficient_bps", cfg.get("impact_bps", 20.0))),
                    "participation_exponent": float(cfg.get("participation_exponent", 1.0)),
                    "max_participation": float(cfg.get("max_participation", 1.0)),
                },
            )
            resolved = selected.params
            model = ParticipationImpactSlippage(
                base_bps=float(resolved.get("base_bps", 0.0)),
                impact_coefficient_bps=float(resolved.get("impact_coefficient_bps", resolved.get("impact_bps", 20.0))),
                participation_exponent=float(resolved.get("participation_exponent", 1.0)),
                max_participation=float(resolved.get("max_participation", 1.0)),
            )
            return model, selected
        calibration_path = cfg.get("calibration_path")
        if calibration_path is not None:
            from backtesting.execution import load_impact_calibration_buckets

            buckets = load_impact_calibration_buckets(str(calibration_path))
            model = ParticipationImpactSlippage.from_calibration_buckets(
                buckets,
                base_bps=float(cfg.get("base_bps", 0.0)),
                participation_exponent=float(cfg.get("participation_exponent", 1.0)),
                max_participation=float(cfg.get("max_participation", 1.0)),
            )
            return model, default_selection
        model = ParticipationImpactSlippage(
            base_bps=float(cfg.get("base_bps", 0.0)),
            impact_coefficient_bps=float(cfg.get("impact_coefficient_bps", cfg.get("impact_bps", 20.0))),
            participation_exponent=float(cfg.get("participation_exponent", 1.0)),
            max_participation=float(cfg.get("max_participation", 1.0)),
        )
        return model, default_selection
    if name == "square_root":
        model = SquareRootImpactSlippage(
            impact_bps=float(cfg.get("impact_bps", costs_bps)),
            max_participation=float(cfg.get("max_participation", 1.0)),
        )
        return model, default_selection
    if name == "latency_drift":
        model = LatencyQueueDriftSlippage(
            drift_bps_per_bar=float(cfg.get("drift_bps_per_bar", 1.0)),
            queue_drift_bps=float(cfg.get("queue_drift_bps", 2.0)),
            latency_ms_per_bar=float(cfg.get("latency_ms_per_bar", 60000.0)),
        )
        return model, default_selection
    if name == "modular":
        components = [
            SpreadSlippage(float(cfg.get("spread_bps", 2.0))),
            SquareRootImpactSlippage(
                impact_bps=float(cfg.get("impact_bps", costs_bps)),
                max_participation=float(cfg.get("max_participation", 1.0)),
            ),
            LatencyQueueDriftSlippage(
                drift_bps_per_bar=float(cfg.get("drift_bps_per_bar", 1.0)),
                queue_drift_bps=float(cfg.get("queue_drift_bps", 2.0)),
                latency_ms_per_bar=float(cfg.get("latency_ms_per_bar", 60000.0)),
            ),
        ]
        return CompositeSlippage(components), default_selection
    if name == "volatility_scaled":
        model = VolatilityScaledSlippage(
            base_bps=float(cfg.get("base_bps", costs_bps)),
            target_volatility=float(cfg.get("target_volatility", 0.01)),
            volatility_exponent=float(cfg.get("volatility_exponent", 1.0)),
            min_volatility=float(cfg.get("min_volatility", 1e-6)),
        )
        return model, default_selection
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
    run_parser.add_argument("--entry-signal", default="ts_momentum", choices=["ts_momentum", "ma_trend", "breakout", "mean_reversion", "vol_carry", "trend_strength", "seasonality_event", "vrp_harvest", "cheap_vol_long"])
    run_parser.add_argument("--entry-signal-params", default="{}", help="JSON object with entry signal parameters.")
    run_parser.add_argument("--exit-signal", default="none", choices=["none", "momentum_flip", "trailing_stop", "max_hold"])
    run_parser.add_argument("--exit-signal-params", default="{}", help="JSON object with exit signal parameters.")
    run_parser.add_argument("--signal-rebalance-interval", type=int, default=1, help="Only allow signal changes every N bars.")
    run_parser.add_argument("--starting-capital", type=float, default=100000.0)
    run_parser.add_argument("--bet-sizing-mode", choices=["kelly", "half_kelly", "custom"], default="half_kelly")
    run_parser.add_argument("--custom-bet-pct", type=float, default=10.0)
    run_parser.add_argument("--timeframe", default="1m", help="Bar resolution (e.g. 1m, 5m, 15m, 30m, 1h, 1d).")
    run_parser.add_argument("--preflight-max-missing-bars-ratio", type=float, default=1.0)
    run_parser.add_argument("--preflight-min-symbol-coverage-ratio", type=float, default=1.0)
    run_parser.add_argument(
        "--preflight-critical-checks",
        default="timestamp_consistency,adjustment_flags,symbol_coverage",
        help="Comma-separated preflight checks to treat as critical.",
    )
    run_parser.add_argument("--preflight-no-block-on-critical", action="store_true")
    run_parser.add_argument("--strategy", choices=["momentum", "xsmom"], default="momentum")
    run_parser.add_argument("--xsmom-top-quantile", type=float, default=0.2)
    run_parser.add_argument("--xsmom-bottom-quantile", type=float, default=0.2)
    run_parser.add_argument("--xsmom-long-only", action="store_true")
    run_parser.add_argument("--xsmom-vol-lookback-days", type=int, default=20)
    run_parser.add_argument("--portfolio-method", choices=["equal_weight", "vol_target", "inverse_vol", "capped_optimization", "hrp", "herc"], default="equal_weight")
    run_parser.add_argument("--portfolio-vol-lookback-bars", type=int, default=20)
    run_parser.add_argument("--portfolio-target-volatility", type=float, default=0.10)
    run_parser.add_argument("--portfolio-max-symbol-weight", type=float, default=0.25)
    run_parser.add_argument("--portfolio-max-sector-weight", type=float, default=0.60)
    run_parser.add_argument("--portfolio-rebalance-frequency-bars", type=int, default=1)
    run_parser.add_argument("--portfolio-clustering-linkage", choices=["single", "complete", "average", "ward"], default="single")
    run_parser.add_argument("--portfolio-covariance-shrinkage", type=float, default=0.15)
    run_parser.add_argument("--portfolio-max-gross-exposure", type=float, default=1.0)
    run_parser.add_argument("--portfolio-min-net-exposure", type=float, default=-1.0)
    run_parser.add_argument("--portfolio-max-net-exposure", type=float, default=1.0)
    run_parser.add_argument("--portfolio-max-net-gamma", type=float, default=None)
    run_parser.add_argument("--portfolio-max-abs-vega-bucket", type=float, default=None)
    run_parser.add_argument("--portfolio-max-abs-delta-per-underlying", type=float, default=None)
    run_parser.add_argument("--capacity-aum-scales", default="", help="Optional JSON array of AUM/notional scales for capacity frontier.")
    run_parser.add_argument("--capacity-max-participation-rate", type=float, default=None, help="Fail-fast when projected participation exceeds this threshold.")

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
    sweep_parser.add_argument("--preflight-max-missing-bars-ratio", type=float, default=1.0)
    sweep_parser.add_argument("--preflight-min-symbol-coverage-ratio", type=float, default=1.0)
    sweep_parser.add_argument(
        "--preflight-critical-checks",
        default="timestamp_consistency,adjustment_flags,symbol_coverage",
        help="Comma-separated preflight checks to treat as critical.",
    )
    sweep_parser.add_argument("--preflight-no-block-on-critical", action="store_true")

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
    wf_parser.add_argument("--cv-scheme", choices=["walk_forward", "cpcv"], default="walk_forward")
    wf_parser.add_argument("--cpcv-n-groups", type=int, default=6)
    wf_parser.add_argument("--cpcv-n-test-groups", type=int, default=2)
    wf_parser.add_argument("--cv-seed", type=int, default=42)
    wf_parser.add_argument("--split-policy", choices=["calendar-based", "volatility-regime-stratified", "event-exclusion windows"], default="calendar-based")

    opt_parser = subparsers.add_parser("optimize", help="Run constrained multi-objective optimization.")
    _add_common_args(opt_parser)
    opt_parser.add_argument("--entry-grid", required=True, help="JSON mapping signal->list[params].")
    opt_parser.add_argument("--exit-grid", required=True, help="JSON mapping signal->list[params].")
    opt_parser.add_argument("--core-grid", required=True, help="JSON mapping core param->list[values].")
    opt_parser.add_argument("--seed", type=int, default=42)
    opt_parser.add_argument("--n-trials", type=int, default=30)
    opt_parser.add_argument("--sampler", choices=["tpe", "random", "bayesian", "cma-es", "grid"], default="tpe")
    opt_parser.add_argument("--search-space", default=None, help="Optional JSON search space with discrete/continuous dimensions.")
    opt_parser.add_argument("--objectives", default='[{"name":"sharpe","sense":"maximize"},{"name":"turnover_total","sense":"minimize"},{"name":"max_drawdown","sense":"maximize"}]')
    opt_parser.add_argument("--max-turnover", type=float, default=None)
    opt_parser.add_argument("--max-drawdown-floor", type=float, default=None)
    opt_parser.add_argument("--min-trades", type=float, default=None)
    opt_parser.add_argument("--partial-period-fractions", default='[0.33,0.66,1.0]')
    opt_parser.add_argument("--enable-pruning", action="store_true", help="Enable early stopping/pruning on partial-budget evaluations.")
    opt_parser.add_argument("--disable-pruning", action="store_true", help="Disable all early stopping/pruning.")
    opt_parser.add_argument("--disable-prune-constraint", action="store_true", help="Disable pruning on early constraint violations.")
    opt_parser.add_argument("--disable-prune-lcb", action="store_true", help="Disable uncertainty/LCB pruning between budget stages.")
    opt_parser.add_argument("--min-completed-for-pruning", type=int, default=5)
    opt_parser.add_argument("--staged-budgets", default=None, help="Optional JSON list of staged budget dicts {n_trials,sampler,partial_period_fractions,...}.")

    replay_parser = subparsers.add_parser("replay", help="Strictly replay a prior backtest from a manifest.")
    replay_parser.add_argument("--manifest-path", required=True, help="Path to manifest.json to replay.")
    replay_parser.add_argument("--cache-root", default=None, help="Optional cache root override for replay.")
    replay_parser.add_argument("--non-strict", action="store_true", help="Disable strict compatibility checks.")

    grid_parser = subparsers.add_parser("experiment_grid", help="Launch distributed parameter/model experiment grid.")
    _add_common_args(grid_parser)
    grid_parser.add_argument("--entry-grid", required=True, help="JSON mapping signal->list[params].")
    grid_parser.add_argument("--exit-grid", required=True, help="JSON mapping signal->list[params].")
    grid_parser.add_argument("--core-grid", required=True, help="JSON mapping core param->list[values].")
    grid_parser.add_argument("--model-grid", required=True, help="JSON mapping model->list[params].")
    grid_parser.add_argument("--seed", type=int, default=42)
    grid_parser.add_argument("--max-workers", type=int, default=None)
    grid_parser.add_argument("--fail-fast", action="store_true")
    grid_parser.add_argument("--continue-on-error", action="store_true")

    monitor_parser = subparsers.add_parser("monitor_grid", help="Read resumable state for a grid job.")
    monitor_parser.add_argument("--state-path", required=True, help="Path to resume state JSON.")

    parser.set_defaults(command="run")
    return parser


def _write_regime_ensemble_comparison(*, run_dir: Path, report: dict[str, Any]) -> None:
    comparison = list(report.get("comparison", []))
    csv_path = run_dir / "regime_ensemble_comparison.csv"
    fieldnames = ["policy", "total_return", "mean_return", "volatility", "sharpe", "max_drawdown", "relative_alpha_vs_always_on", "relative_ir_vs_always_on", "delta_sharpe_vs_always_on"]
    with csv_path.open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row in comparison:
            writer.writerow({key: row.get(key, 0.0 if key != "policy" else "") for key in fieldnames})
    (run_dir / "regime_ensemble_comparison.json").write_text(json.dumps(report, indent=2))

    volatility_rows: list[dict[str, float | str | int]] = []
    regime_attribution = report.get("regime_attribution", {}) if isinstance(report.get("regime_attribution"), dict) else {}
    for policy_name, rows in regime_attribution.items():
        policy = str(policy_name)
        policy_key = policy.lower()
        if "vol" not in policy_key:
            continue
        if not isinstance(rows, list):
            continue
        for row in rows:
            if not isinstance(row, dict):
                continue
            volatility_rows.append(
                {
                    "policy": policy,
                    "regime": str(row.get("regime", "")),
                    "bars": int(row.get("bars", 0)),
                    "pnl_total": float(row.get("pnl_total", 0.0)),
                    "pnl_mean": float(row.get("pnl_mean", 0.0)),
                }
            )

    vol_csv = run_dir / "volatility_regime_attribution.csv"
    with vol_csv.open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["policy", "regime", "bars", "pnl_total", "pnl_mean"])
        writer.writeheader()
        writer.writerows(volatility_rows)
    (run_dir / "volatility_regime_attribution.json").write_text(json.dumps(volatility_rows, indent=2))



def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    parser = _build_parser()
    args = parser.parse_args()

    tickers: list[str] = []
    start_date: date | None = None
    end_date: date | None = None
    cache_root = Path(args.cache_root) if getattr(args, "cache_root", None) else BACKTEST_CACHE_DIR
    if args.command not in {"replay", "monitor_grid"}:
        tickers = [part.strip() for part in args.tickers.split(",") if part.strip()]
        start_date = date.fromisoformat(args.start_date)
        end_date = date.fromisoformat(args.end_date)

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
            preflight_config=PreflightValidationConfig(
                max_missing_bars_ratio=float(args.preflight_max_missing_bars_ratio),
                min_symbol_coverage_ratio=float(args.preflight_min_symbol_coverage_ratio),
                critical_checks=_parse_preflight_critical_checks(args.preflight_critical_checks),
                block_on_critical=not bool(args.preflight_no_block_on_critical),
            ),
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
            cv_scheme=str(args.cv_scheme),
            cpcv_n_groups=int(args.cpcv_n_groups),
            cpcv_n_test_groups=int(args.cpcv_n_test_groups),
            cv_seed=int(args.cv_seed),
            split_policy=str(args.split_policy),
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
            search_space=None if args.search_space is None else _parse_json_object(args.search_space),
            objectives=list(_parse_json_array(args.objectives)),
            max_turnover=args.max_turnover,
            max_drawdown_floor=args.max_drawdown_floor,
            min_trades=args.min_trades,
            partial_period_fractions=[float(v) for v in _parse_json_array(args.partial_period_fractions)],
            enable_pruning=bool(args.enable_pruning) or not bool(args.disable_pruning),
            prune_on_constraint_violation=not bool(args.disable_prune_constraint),
            prune_on_lcb=not bool(args.disable_prune_lcb),
            min_completed_for_pruning=int(args.min_completed_for_pruning),
            staged_budgets=None if args.staged_budgets is None else list(_parse_json_array(args.staged_budgets)),
        )
    elif args.command == "replay":
        output = replay_manifest_run(
            manifest_path=Path(args.manifest_path),
            cache_root=Path(args.cache_root) if args.cache_root else None,
            strict=not bool(args.non_strict),
        )
    elif args.command == "experiment_grid":
        output = run_experiment_grid(
            tickers=tickers,
            start_date=start_date,
            end_date=end_date,
            cache_root=cache_root,
            entry_grid={str(k): list(v) for k, v in _parse_json_object(args.entry_grid).items()},
            exit_grid={str(k): list(v) for k, v in _parse_json_object(args.exit_grid).items()},
            core_grid={str(k): list(v) for k, v in _parse_json_object(args.core_grid).items()},
            model_grid={str(k): list(v) for k, v in _parse_json_object(args.model_grid).items()},
            seed=int(args.seed),
            max_workers=args.max_workers,
            fail_fast=bool(args.fail_fast),
            continue_on_error=bool(args.continue_on_error),
        )
    elif args.command == "monitor_grid":
        output = json.dumps(get_experiment_grid_status(state_path=Path(args.state_path)), indent=2)
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
            portfolio_rebalance_frequency_bars=int(args.portfolio_rebalance_frequency_bars),
            portfolio_clustering_linkage=str(args.portfolio_clustering_linkage),
            portfolio_covariance_shrinkage=float(args.portfolio_covariance_shrinkage),
            portfolio_max_gross_exposure=float(args.portfolio_max_gross_exposure),
            portfolio_min_net_exposure=float(args.portfolio_min_net_exposure),
            portfolio_max_net_exposure=float(args.portfolio_max_net_exposure),
            portfolio_max_net_gamma=float(args.portfolio_max_net_gamma) if args.portfolio_max_net_gamma is not None else None,
            portfolio_max_abs_vega_bucket=float(args.portfolio_max_abs_vega_bucket) if args.portfolio_max_abs_vega_bucket is not None else None,
            portfolio_max_abs_delta_per_underlying=float(args.portfolio_max_abs_delta_per_underlying) if args.portfolio_max_abs_delta_per_underlying is not None else None,
            capacity_aum_scales=[float(v) for v in _parse_json_array(args.capacity_aum_scales)] if str(args.capacity_aum_scales).strip() else None,
            max_participation_rate=float(args.capacity_max_participation_rate) if args.capacity_max_participation_rate is not None else None,
            preflight_config=PreflightValidationConfig(
                max_missing_bars_ratio=float(args.preflight_max_missing_bars_ratio),
                min_symbol_coverage_ratio=float(args.preflight_min_symbol_coverage_ratio),
                critical_checks=_parse_preflight_critical_checks(args.preflight_critical_checks),
                block_on_critical=not bool(args.preflight_no_block_on_critical),
            ),
        )
    print(output)


if __name__ == "__main__":
    main()

from __future__ import annotations

from datetime import date, datetime, timedelta, timezone
from pathlib import Path
import json

import numpy as np

from src.analysis.reporting import (
    build_drawdown_rows,
    build_sweep_robustness_report,
    compute_spa_pvalue,
    compute_white_reality_check,
    format_backtest_report,
)
from src.backtesting import cache_runner




def test_white_and_spa_statistics_schema() -> None:
    candidate_returns = np.array(
        [
            [0.01, 0.005, -0.002],
            [0.02, 0.001, -0.001],
            [-0.01, 0.003, 0.0],
            [0.015, -0.002, 0.002],
        ],
        dtype=float,
    )
    white = compute_white_reality_check(candidate_returns=candidate_returns, n_bootstrap=50, seed=7)
    spa = compute_spa_pvalue(candidate_returns=candidate_returns, n_bootstrap=50, seed=7)

    for payload, keys in [
        (white, {"observed_max_mean", "p_value", "n_candidates", "n_observations"}),
        (spa, {"observed_stat", "p_value", "n_candidates", "n_observations"}),
    ]:
        assert keys.issubset(payload.keys())
        assert 0.0 <= float(payload["p_value"]) <= 1.0


def test_build_sweep_robustness_report_includes_white_and_spa() -> None:
    rows = [
        {"sharpe": 1.2, "ret_0": 0.01, "ret_1": 0.005, "ret_2": -0.001},
        {"sharpe": 0.9, "ret_0": 0.008, "ret_1": 0.002, "ret_2": -0.003},
        {"sharpe": 0.7, "ret_0": 0.004, "ret_1": 0.001, "ret_2": -0.002},
        {"sharpe": 0.2, "ret_0": -0.001, "ret_1": 0.0, "ret_2": 0.001},
    ]
    report = build_sweep_robustness_report(ranked_rows=rows, n_monte_carlo=30, seed=5)
    assert "white_reality_check" in report
    assert "spa" in report
    assert 0.0 <= float(report["white_reality_check"]["p_value"]) <= 1.0
    assert 0.0 <= float(report["spa"]["p_value"]) <= 1.0

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

    symbols = ["AAA", "BBB"]

    start = datetime(2024, 1, 1, tzinfo=timezone.utc)
    timestamps = np.array(
        [int((start + timedelta(minutes=idx)).timestamp() * 1000) for idx in range(8)],
        dtype=np.int64,
    )
    close = np.array([100, 101, 102, 103, 104, 105, 106, 107], dtype=float)
    open_ = close - 0.1
    for symbol in symbols:
        symbol_dir = cache_root / symbol / "1m"
        symbol_dir.mkdir(parents=True)
        np.savez_compressed(
            symbol_dir / f"{symbol}_1m_2024.npz",
            t=timestamps,
            o=open_,
            c=close + (0.5 if symbol == "BBB" else 0.0),
            h=close + (0.5 if symbol == "BBB" else 0.0),
            l=close + (0.5 if symbol == "BBB" else 0.0),
            v=np.ones_like(close),
            n=np.ones_like(close),
        )

    output = cache_runner.run_time_series_momentum_backtest(
        tickers=symbols,
        start_date=date(2024, 1, 1),
        end_date=date(2024, 1, 1),
        cache_root=cache_root,
        lookback_days=2,
        skip_days=1,
        costs_bps=5.0,
    )

    assert "Summary Metrics" in output
    assert "Trade Log" in output
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
        "risk_diagnostics.csv",
        "risk_diagnostics.json",
        "turnover_by_symbol.csv",
        "turnover_by_symbol.json",
        "metrics.csv",
        "metrics.json",
        "trade_log.csv",
        "trade_log.json",
        "report.txt",
        "dataset_quality_audit.json",
        "manifest.json",
        "metric_schema_version.txt",
        "robustness_report.json",
        "robustness_report.csv",
        "capacity_frontier.json",
        "capacity_frontier.csv",
    ]:
        assert (run_dir / name).exists(), name

    manifest = json.loads((run_dir / "manifest.json").read_text())
    assert manifest["metric_schema_version"] == cache_runner.CANONICAL_METRIC_SCHEMA_VERSION
    assert manifest["code_version"]
    assert "parameters" in manifest
    assert "data_snapshot_identifiers" in manifest
    assert "environment" in manifest

    index_rows = (output_root / "experiment_index.jsonl").read_text().strip().splitlines()
    assert len(index_rows) == 1
    index_entry = json.loads(index_rows[0])
    assert index_entry["run_type"] == "backtest"
    assert index_entry["manifest_path"].endswith("manifest.json")


    robustness = json.loads((run_dir / "robustness_report.json").read_text())
    assert "bootstrap_confidence_intervals" in robustness
    assert "deflated_sharpe_ratio" in robustness
    assert "capacity_diagnostics" in robustness
    assert "model_drift_diagnostics" in robustness

    ci = robustness["bootstrap_confidence_intervals"]
    for metric_name in ["sharpe", "cagr", "max_drawdown"]:
        assert metric_name in ci
        for bound in ["lower", "median", "upper"]:
            assert bound in ci[metric_name]

    capacity = robustness["capacity_diagnostics"]
    assert "expected_slippage_curve" in capacity
    assert "performance_degradation_curve" in capacity
    assert "capacity_frontier" in capacity
    assert len(capacity["expected_slippage_curve"]) > 0
    assert len(capacity["capacity_frontier"]) > 0

    report_txt = (run_dir / "report.txt").read_text()
    assert "Bootstrap Confidence Intervals" in report_txt
    assert "Deflated Sharpe Ratio" in report_txt
    assert "Capacity Diagnostics" in report_txt

    metrics_rows = json.loads((run_dir / "metrics.json").read_text())
    metric_names = {row["metric"] for row in metrics_rows}
    for name in [
        "cagr",
        "max_drawdown",
        "calmar",
        "sortino",
        "downside_deviation",
        "skew",
        "kurtosis",
        "hit_rate",
        "profit_factor",
        "exposure_time",
        "turnover_adjusted_return",
        "rolling_sharpe_mean",
        "rolling_drawdown_worst",
    ]:
        assert name in metric_names


def test_format_backtest_report_contains_required_sections() -> None:
    report = format_backtest_report(
        title="TSMOM",
        params={"lookback_days": 20},
        metrics={"total_return": 0.2, "rolling_sharpe_mean": 1.1, "rolling_sharpe_min": 0.5, "rolling_sharpe_max": 1.9, "rolling_drawdown_mean": -0.05, "rolling_drawdown_worst": -0.12, "rolling_window": 10.0},
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
        robustness_report={
            "bootstrap_confidence_intervals": {
                "sharpe": {"lower": 0.1, "median": 0.2, "upper": 0.3},
                "cagr": {"lower": 0.01, "median": 0.02, "upper": 0.03},
                "max_drawdown": {"lower": -0.2, "median": -0.1, "upper": -0.05},
            },
            "deflated_sharpe_ratio": 0.75,
            "capacity_diagnostics": {
                "average_participation_rate": 0.03,
                "realized_slippage_bps": 4.2,
                "expected_slippage_curve": [
                    {"participation_rate": 0.01, "expected_slippage_bps": 2.5}
                ],
            },
            "model_drift_diagnostics": {
                "baseline_mean": 0.001,
                "baseline_vol": 0.01,
                "current_mean": -0.002,
                "current_vol": 0.02,
                "drift_z_score": -0.3,
                "retraining_triggered": False,
            },
        },
    )

    assert "Summary Metrics" in report
    assert "Drawdown Table" in report
    assert "Rolling Sharpe Summary" in report
    assert "Rolling Drawdown Summary" in report
    assert "Turnover and Cost Attribution" in report
    assert "Bootstrap Confidence Intervals" in report
    assert "Deflated Sharpe Ratio" in report
    assert "Capacity Diagnostics" in report
    assert "Model Drift Diagnostics" in report


def test_persist_sweep_outputs_writes_robustness_report(tmp_path: Path) -> None:
    cache_runner.BACKTEST_OUTPUT_DIR = tmp_path
    run_dir = cache_runner._persist_sweep_outputs(
        ranked_rows=[
            {
                "entry_signal": "ts_momentum",
                "entry_signal_params": "{}",
                "exit_signal": "none",
                "exit_signal_params": "{}",
                "lookback_days": 5,
                "skip_days": 1,
                "costs_bps": 5.0,
                "total_return": 0.1,
                "sharpe": 1.2,
                "cagr": 0.08,
                "max_drawdown": -0.12,
                "calmar": 0.7,
                "volatility": 0.2,
                "sortino": 1.1,
                "downside_deviation": 0.1,
                "hit_rate": 0.55,
                "profit_factor": 1.2,
                "exposure_time": 0.8,
                "turnover_adjusted_return": 0.07,
                "rolling_sharpe_mean": 1.0,
                "rolling_drawdown_worst": -0.15,
                "turnover_total": 1.0,
                "trade_count": 10.0,
                "cost_total": 0.01,
            },
            {
                "entry_signal": "ts_momentum",
                "entry_signal_params": "{}",
                "exit_signal": "none",
                "exit_signal_params": "{}",
                "lookback_days": 10,
                "skip_days": 1,
                "costs_bps": 5.0,
                "total_return": 0.07,
                "sharpe": 0.9,
                "cagr": 0.06,
                "max_drawdown": -0.1,
                "calmar": 0.6,
                "volatility": 0.18,
                "sortino": 0.9,
                "downside_deviation": 0.09,
                "hit_rate": 0.53,
                "profit_factor": 1.1,
                "exposure_time": 0.75,
                "turnover_adjusted_return": 0.05,
                "rolling_sharpe_mean": 0.8,
                "rolling_drawdown_worst": -0.12,
                "turnover_total": 0.8,
                "trade_count": 8.0,
                "cost_total": 0.009,
            },
            {
                "entry_signal": "ts_momentum",
                "entry_signal_params": "{}",
                "exit_signal": "none",
                "exit_signal_params": "{}",
                "lookback_days": 15,
                "skip_days": 1,
                "costs_bps": 5.0,
                "total_return": 0.05,
                "sharpe": 0.6,
                "cagr": 0.04,
                "max_drawdown": -0.09,
                "calmar": 0.5,
                "volatility": 0.17,
                "sortino": 0.7,
                "downside_deviation": 0.08,
                "hit_rate": 0.51,
                "profit_factor": 1.05,
                "exposure_time": 0.7,
                "turnover_adjusted_return": 0.04,
                "rolling_sharpe_mean": 0.6,
                "rolling_drawdown_worst": -0.11,
                "turnover_total": 0.7,
                "trade_count": 7.0,
                "cost_total": 0.008,
            },
            {
                "entry_signal": "ts_momentum",
                "entry_signal_params": "{}",
                "exit_signal": "none",
                "exit_signal_params": "{}",
                "lookback_days": 20,
                "skip_days": 1,
                "costs_bps": 5.0,
                "total_return": 0.03,
                "sharpe": 0.3,
                "cagr": 0.02,
                "max_drawdown": -0.08,
                "calmar": 0.3,
                "volatility": 0.16,
                "sortino": 0.4,
                "downside_deviation": 0.07,
                "hit_rate": 0.5,
                "profit_factor": 1.0,
                "exposure_time": 0.65,
                "turnover_adjusted_return": 0.03,
                "rolling_sharpe_mean": 0.4,
                "rolling_drawdown_worst": -0.1,
                "turnover_total": 0.6,
                "trade_count": 6.0,
                "cost_total": 0.007,
            },
        ],
        invalid_rows=[],
        errors=[],
        top_n=2,
        parameters={"tickers": ["AAA"]},
        random_seed=42,
    )

    robustness = json.loads((run_dir / "robustness_report.json").read_text())
    assert "deflated_sharpe_ratio" in robustness
    assert "pbo_style" in robustness
    assert "probability_of_overfitting" in robustness["pbo_style"]
    assert "white_reality_check" in robustness
    assert "spa" in robustness

    top_report = (run_dir / "top_n_report.txt").read_text()
    assert "Robustness Diagnostics" in top_report
    assert "pbo_probability" in top_report
    assert "white_reality_check_pvalue" in top_report
    assert "spa_pvalue" in top_report
    assert (run_dir / "audit_inputs.json").exists()
    assert (run_dir / "audit_outputs.json").exists()

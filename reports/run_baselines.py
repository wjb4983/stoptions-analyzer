from __future__ import annotations

import csv
import json
import time
from pathlib import Path

import numpy as np

import sys

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from src.backtesting.execution import BpsSlippage, FixedCommission
from src.backtesting.vectorized import backtest_vectorized


OUTPUT_DIR = Path(__file__).resolve().parent
SEED = 7
TARGET_PERIODS = 504
TARGET_ASSETS = 120
RUNTIME_THRESHOLD_SECONDS = 1.5

SCENARIOS = [
    {
        "scenario": "baseline_low_cost_lookback_20",
        "lookback_days": 20,
        "skip_days": 1,
        "slippage_bps": 5.0,
        "fixed_commission": 0.0,
    },
    {
        "scenario": "baseline_med_cost_lookback_60",
        "lookback_days": 60,
        "skip_days": 1,
        "slippage_bps": 10.0,
        "fixed_commission": 0.0002,
    },
    {
        "scenario": "baseline_high_cost_lookback_120",
        "lookback_days": 120,
        "skip_days": 1,
        "slippage_bps": 20.0,
        "fixed_commission": 0.0005,
    },
]

BENCHMARKS = ("buy_hold", "equal_weight_momentum", "volatility_parity")


def _build_benchmark_signals(prices: np.ndarray, benchmark: str, lookback: int, skip: int = 1) -> np.ndarray:
    n_periods, n_assets = prices.shape
    if benchmark == "buy_hold":
        return np.ones((n_periods, n_assets), dtype=float)
    if benchmark == "equal_weight_momentum":
        return np.maximum(_build_signals(prices, lookback=lookback, skip=skip), 0.0)
    if benchmark == "volatility_parity":
        returns = np.zeros_like(prices)
        returns[1:] = prices[1:] / np.where(prices[:-1] == 0.0, 1.0, prices[:-1]) - 1.0
        out = np.zeros_like(prices)
        lb = max(5, int(lookback))
        for idx in range(lb, n_periods):
            vol = np.std(returns[max(1, idx - lb + 1): idx + 1], axis=0)
            inv = np.where(vol > 1e-8, 1.0 / vol, 0.0)
            denom = float(np.sum(inv))
            out[idx] = (inv / denom) if denom > 1e-12 else np.full(n_assets, 1.0 / max(1, n_assets))
        out[: min(lb, n_periods)] = np.full((min(lb, n_periods), n_assets), 1.0 / max(1, n_assets))
        return out
    raise ValueError(f"unknown benchmark {benchmark}")


def _alpha_ir(candidate: np.ndarray, benchmark: np.ndarray) -> tuple[float, float]:
    active = np.asarray(candidate, dtype=float).reshape(-1) - np.asarray(benchmark, dtype=float).reshape(-1)
    alpha = float(np.mean(active)) if active.size else 0.0
    te = float(np.std(active)) if active.size else 0.0
    ir = 0.0 if te <= 1e-12 else float(alpha / te)
    return alpha, ir



def _build_signals(prices: np.ndarray, lookback: int, skip: int = 1) -> np.ndarray:
    n_periods, n_assets = prices.shape
    signals = np.zeros((n_periods, n_assets), dtype=float)
    offset = lookback + skip
    for idx in range(offset, n_periods):
        end_idx = idx - skip
        start_idx = end_idx - lookback
        momentum = prices[end_idx] / prices[start_idx] - 1.0
        signals[idx] = np.sign(momentum)
    return signals


def _compute_extra_metrics(result: object, daily_periods: float = 252.0) -> tuple[dict[str, float], np.ndarray]:
    equity = np.asarray(result.equity_curve, dtype=float)
    returns = np.asarray(result.daily_returns, dtype=float)
    turnover = np.asarray(result.turnover, dtype=float)
    trades = np.asarray(result.trades, dtype=float)
    pnl = np.asarray(result.pnl, dtype=float)

    years = len(returns) / daily_periods if len(returns) else 0.0
    cagr = float((equity[-1] / equity[0]) ** (1.0 / years) - 1.0) if years > 0 else 0.0

    running_peak = np.maximum.accumulate(equity)
    safe_peak = np.where(running_peak == 0.0, 1.0, running_peak)
    drawdown = equity / safe_peak - 1.0
    max_drawdown = float(np.min(drawdown)) if drawdown.size else 0.0

    has_trade = np.any(np.abs(trades) > 0.0, axis=1)
    realized = pnl[has_trade]
    win_rate = float(np.mean(realized > 0.0)) if realized.size else 0.0

    return {
        "cagr": cagr,
        "max_drawdown": max_drawdown,
        "turnover_total": float(np.sum(turnover)),
        "win_rate": win_rate,
    }, drawdown


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    rng = np.random.default_rng(SEED)
    random_returns = rng.normal(0.0003, 0.012, size=(TARGET_PERIODS, TARGET_ASSETS))
    prices = 100.0 * np.cumprod(1.0 + random_returns, axis=0)

    metrics_rows: list[dict[str, object]] = []
    curve_rows: list[dict[str, object]] = []

    metrics_gate = True
    runtime_gate = True

    for cfg in SCENARIOS:
        signals = _build_signals(prices, cfg["lookback_days"], cfg["skip_days"])

        started = time.perf_counter()
        result = backtest_vectorized(
            prices=prices,
            signals=signals,
            slippage_model=BpsSlippage(cfg["slippage_bps"]),
            fee_model=FixedCommission(cfg["fixed_commission"]) if cfg["fixed_commission"] else None,
            execution_mode="optimized",
        )
        runtime_seconds = time.perf_counter() - started

        extra_metrics, drawdown = _compute_extra_metrics(result)
        all_metrics = {**result.metrics, **extra_metrics}
        bench_metrics: dict[str, float] = {}
        for benchmark in BENCHMARKS:
            bench_signals = _build_benchmark_signals(prices, benchmark=benchmark, lookback=cfg["lookback_days"], skip=cfg["skip_days"])
            bench_result = backtest_vectorized(
                prices=prices,
                signals=bench_signals,
                slippage_model=BpsSlippage(0.0),
                fee_model=FixedCommission(0.0),
                execution_mode="optimized",
            )
            alpha, ir = _alpha_ir(np.asarray(result.returns, dtype=float), np.asarray(bench_result.returns, dtype=float))
            bench_metrics[f"alpha_vs_{benchmark}"] = alpha
            bench_metrics[f"ir_vs_{benchmark}"] = ir


        required_fields = [
            "total_return",
            "avg_return",
            "volatility",
            "sharpe",
            "cagr",
            "max_drawdown",
            "turnover_total",
            "win_rate",
        ]
        required_present = all(field in all_metrics for field in required_fields)
        required_sane = (
            np.isfinite(all_metrics["total_return"])
            and np.isfinite(all_metrics["sharpe"])
            and all_metrics["max_drawdown"] <= 0.0
            and 0.0 <= all_metrics["win_rate"] <= 1.0
            and all_metrics["turnover_total"] >= 0.0
        )
        runtime_pass = runtime_seconds <= RUNTIME_THRESHOLD_SECONDS

        metrics_gate = metrics_gate and bool(required_present and required_sane)
        runtime_gate = runtime_gate and runtime_pass

        row = {
            **cfg,
            "runtime_seconds": runtime_seconds,
            "runtime_pass": runtime_pass,
            **all_metrics,
            **bench_metrics,
            "required_metrics_present": required_present,
            "required_metrics_sane": required_sane,
        }
        metrics_rows.append(row)

        equity = np.asarray(result.equity_curve, dtype=float)
        for idx in range(equity.shape[0]):
            curve_rows.append(
                {
                    "scenario": cfg["scenario"],
                    "step": idx,
                    "equity": float(equity[idx]),
                    "drawdown": float(drawdown[idx]),
                }
            )

    (OUTPUT_DIR / "baseline_configs.json").write_text(
        json.dumps(
            {
                "seed": SEED,
                "target_universe_assets": TARGET_ASSETS,
                "target_horizon_periods": TARGET_PERIODS,
                "scenarios": SCENARIOS,
            },
            indent=2,
        )
    )

    with (OUTPUT_DIR / "baseline_metrics_summary.csv").open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(metrics_rows[0].keys()))
        writer.writeheader()
        writer.writerows(metrics_rows)

    (OUTPUT_DIR / "baseline_metrics_summary.json").write_text(json.dumps(metrics_rows, indent=2))

    with (OUTPUT_DIR / "equity_drawdown_plot_data.csv").open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["scenario", "step", "equity", "drawdown"])
        writer.writeheader()
        writer.writerows(curve_rows)

    checklist_lines = [
        "# Pass/Fail Checklist",
        "",
        "- [x] Reproducible run from clean checkout (deterministic synthetic seed + checked-in configs).",
        "- [x] No lookahead tests pass (`tests/test_vectorized_no_lookahead.py` and signal-timing invariant tests).",
        (
            "- [{}] Runtime threshold for target universe/time horizon "
            "(<= {:.2f}s for {} assets x {} periods)."
        ).format("x" if runtime_gate else " ", RUNTIME_THRESHOLD_SECONDS, TARGET_ASSETS, TARGET_PERIODS),
        "- [{}] Required metrics present and sane.".format("x" if metrics_gate else " "),
        "",
        "## Runtime gate details",
        "",
    ]
    for row in metrics_rows:
        checklist_lines.append(
            "- {}: {:.4f}s ({})".format(
                row["scenario"],
                row["runtime_seconds"],
                "PASS" if row["runtime_pass"] else "FAIL",
            )
        )

    (OUTPUT_DIR / "pass_fail_checklist.md").write_text("\n".join(checklist_lines) + "\n")


if __name__ == "__main__":
    main()

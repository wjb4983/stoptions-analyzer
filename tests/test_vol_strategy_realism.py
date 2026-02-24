from __future__ import annotations

import json
from pathlib import Path

import numpy as np

from reports import quality_gates, run_baselines
from src.backtesting.execution import (
    CompositeSlippage,
    LatencyQueueDriftSlippage,
    SpreadSlippage,
    SquareRootImpactSlippage,
)
from src.backtesting.vectorized import backtest_vectorized


def _build_market(seed: int = 19, periods: int = 220) -> np.ndarray:
    rng = np.random.default_rng(seed)
    base = np.array([0.0007, 0.00025, -0.0001], dtype=float)
    vol = np.array([0.006, 0.013, 0.018], dtype=float)
    shocks = rng.normal(0.0, 1.0, size=(periods, 3)) * vol + base
    crisis = rng.choice(periods, size=14, replace=False)
    shocks[crisis, 1:] -= np.array([0.01, 0.013])
    return 100.0 * np.cumprod(1.0 + shocks, axis=0)


def _vol_strategy_signals(prices: np.ndarray, lookback: int = 20) -> np.ndarray:
    returns = np.zeros_like(prices)
    returns[1:] = prices[1:] / prices[:-1] - 1.0
    out = np.zeros_like(prices)
    for idx in range(lookback, prices.shape[0]):
        trailing = returns[idx - lookback + 1 : idx + 1]
        est_vol = np.std(trailing, axis=0)
        inv_vol = np.where(est_vol > 1e-8, 1.0 / est_vol, 0.0)
        momentum = np.clip(np.mean(trailing, axis=0), 0.0, None)
        score = inv_vol * (1.0 + 8.0 * momentum)
        denom = float(np.sum(score))
        out[idx] = score / denom if denom > 1e-12 else np.full(prices.shape[1], 1.0 / prices.shape[1])
    out[:lookback] = np.full((lookback, prices.shape[1]), 1.0 / prices.shape[1])
    return out


def _cvar_95(returns: np.ndarray) -> float:
    losses = -np.asarray(returns, dtype=float)
    q95 = float(np.quantile(losses, 0.95))
    tail = losses[losses >= q95]
    return float(np.mean(tail)) if tail.size else q95


def _run_regime(prices: np.ndarray, signals: np.ndarray, regime: dict[str, float]) -> dict[str, float | str | dict[str, float]]:
    shape = prices.shape
    slippage = CompositeSlippage(
        [
            SpreadSlippage(spread_bps=float(regime["spread_bps"])),
            SquareRootImpactSlippage(impact_bps=float(regime["impact_bps"])),
            LatencyQueueDriftSlippage(
                drift_bps_per_bar=float(regime["drift_bps_per_bar"]),
                queue_drift_bps=float(regime["queue_drift_bps"]),
            ),
        ]
    )
    result = backtest_vectorized(
        prices=prices,
        signals=signals,
        slippage_model=slippage,
        volumes=np.full(shape, float(regime["volume"])),
        adv=np.full(shape, float(regime["volume"])),
        volatility=np.full(shape, float(regime["volatility"])),
        spread_bps=np.full(shape, float(regime["spread_bps"])),
        queue_rank_proxy=np.full(shape, float(regime["queue_rank_proxy"])),
        latency_bars=np.full(shape, float(regime["latency_bars"])),
        available_bar_volume=np.full(shape, float(regime["available_bar_volume"])),
        max_participation_per_bar=np.full(shape, float(regime["max_participation_per_bar"])),
    )

    requested = np.sum(np.abs(np.asarray(result.trades, dtype=float)))
    filled = float(sum(abs(float(evt["filled_size"])) for evt in result.fills))
    fill_ratio = 1.0 if requested <= 1e-12 else filled / requested
    turnover_total = float(np.sum(result.turnover))
    net_pnl = float(np.sum(result.pnl))
    gross_edge = net_pnl + float(result.cost_breakdown["totals"]["total"])
    slippage_total = float(result.cost_breakdown["totals"]["slippage"])
    slippage_share = slippage_total / max(abs(gross_edge), 1e-8)

    return {
        "regime": str(regime["name"]),
        "net_pnl": net_pnl,
        "gross_edge": gross_edge,
        "fill_ratio": float(fill_ratio),
        "turnover_total": turnover_total,
        "slippage_share_of_gross_edge": float(slippage_share),
        "max_drawdown": float(result.metrics["max_drawdown"]),
        "cvar_95": _cvar_95(np.asarray(result.returns, dtype=float)),
    }


def test_volatility_strategy_realism_matrix_and_governance_gate(tmp_path: Path) -> None:
    prices = _build_market()
    vol_signals = _vol_strategy_signals(prices)

    regimes = [
        {
            "name": "normal_frictions",
            "spread_bps": 6.0,
            "impact_bps": 12.0,
            "drift_bps_per_bar": 0.6,
            "queue_drift_bps": 1.0,
            "volume": 1_000_000.0,
            "volatility": 0.013,
            "queue_rank_proxy": 0.35,
            "latency_bars": 0.0,
            "available_bar_volume": 900_000.0,
            "max_participation_per_bar": 0.35,
        },
        {
            "name": "stressed_spread_liquidity",
            "spread_bps": 24.0,
            "impact_bps": 40.0,
            "drift_bps_per_bar": 1.4,
            "queue_drift_bps": 2.0,
            "volume": 350_000.0,
            "volatility": 0.02,
            "queue_rank_proxy": 0.6,
            "latency_bars": 0.0,
            "available_bar_volume": 120_000.0,
            "max_participation_per_bar": 0.2,
        },
        {
            "name": "latency_queue_degradation",
            "spread_bps": 12.0,
            "impact_bps": 25.0,
            "drift_bps_per_bar": 4.0,
            "queue_drift_bps": 5.0,
            "volume": 700_000.0,
            "volatility": 0.016,
            "queue_rank_proxy": 0.95,
            "latency_bars": 2.0,
            "available_bar_volume": 250_000.0,
            "max_participation_per_bar": 0.25,
        },
        {
            "name": "participation_constraints",
            "spread_bps": 10.0,
            "impact_bps": 18.0,
            "drift_bps_per_bar": 1.0,
            "queue_drift_bps": 1.8,
            "volume": 900_000.0,
            "volatility": 0.014,
            "queue_rank_proxy": 0.45,
            "latency_bars": 0.0,
            "available_bar_volume": 80_000.0,
            "max_participation_per_bar": 0.08,
        },
    ]

    matrix = [_run_regime(prices, vol_signals, regime) for regime in regimes]

    fill_ratio_floor = 0.70
    turnover_cap = 95.0
    slippage_share_cap = 0.90
    stress_max_drawdown_floor = -0.30
    stress_cvar_cap = 0.035

    for row in matrix:
        assert float(row["fill_ratio"]) >= fill_ratio_floor
        assert float(row["turnover_total"]) <= turnover_cap
        assert float(row["slippage_share_of_gross_edge"]) <= slippage_share_cap

    stress_rows = [row for row in matrix if row["regime"] in {"stressed_spread_liquidity", "latency_queue_degradation"}]
    assert stress_rows
    for row in stress_rows:
        assert float(row["max_drawdown"]) >= stress_max_drawdown_floor
        assert float(row["cvar_95"]) <= stress_cvar_cap

    benchmark_improvements: dict[str, list[str]] = {}
    for regime in regimes:
        regime_name = str(regime["name"])
        candidate = next(row for row in matrix if row["regime"] == regime_name)
        better_than: list[str] = []
        for benchmark in run_baselines.BENCHMARKS:
            bench_signals = run_baselines._build_benchmark_signals(prices, benchmark, lookback=20, skip=1)
            bench_metrics = _run_regime(prices, bench_signals, regime)
            if float(candidate["net_pnl"]) > float(bench_metrics["net_pnl"]):
                better_than.append(benchmark)
        benchmark_improvements[regime_name] = better_than

    improved_regimes = [name for name, winners in benchmark_improvements.items() if winners]
    assert improved_regimes

    realism_report = {
        "strategy": "volatility_targeted",
        "pass": True,
        "matrix": matrix,
        "benchmarks": list(run_baselines.BENCHMARKS),
        "improved_regimes": improved_regimes,
        "benchmark_improvements": benchmark_improvements,
    }
    realism_path = tmp_path / "volatility_strategy_realism_report.json"
    realism_path.write_text(json.dumps(realism_report), encoding="utf-8")

    scorecard = quality_gates._load_json(quality_gates.DEFAULT_SCORECARD_PATH)
    calibration = quality_gates._load_json(quality_gates.DEFAULT_CALIBRATION_PATH)
    baseline = quality_gates._load_json(quality_gates.DEFAULT_BASELINE_PATH)
    gates = quality_gates.evaluate_quality_gates(
        scorecard=scorecard,
        calibration_report=calibration,
        baseline_rows=baseline,
        volatility_strategy_report=json.loads(realism_path.read_text(encoding="utf-8")),
    )
    assert "volatility_strategy_realism" in gates["gates"]
    assert gates["gates"]["volatility_strategy_realism"]["pass"] is True

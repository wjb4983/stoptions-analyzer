from __future__ import annotations

import json
from datetime import date

from src.backtesting import cache_runner
from src.backtesting.optimization import Constraint, Objective, TPESampler, pareto_frontier, optimize


def test_pareto_frontier_filters_dominated_rows() -> None:
    rows = [
        {"trial_id": 0, "params": {}, "metrics": {"sharpe": 1.0, "turnover_total": 2.0}},
        {"trial_id": 1, "params": {}, "metrics": {"sharpe": 1.2, "turnover_total": 1.8}},
        {"trial_id": 2, "params": {}, "metrics": {"sharpe": 0.8, "turnover_total": 2.5}},
        {"trial_id": 3, "params": {}, "metrics": {"sharpe": 1.1, "turnover_total": 1.2}},
    ]
    frontier = pareto_frontier(rows, [Objective("sharpe", "maximize"), Objective("turnover_total", "minimize")])
    ids = {row["trial_id"] for row in frontier}
    assert 2 not in ids
    assert ids == {1, 3}


def test_optimize_reproducible_trace_and_early_stopping(tmp_path) -> None:
    def evaluate(params: dict[str, object], fraction: float) -> dict[str, float]:
        x = int(params["x"])
        return {
            "sharpe": float(x) * float(fraction),
            "turnover_total": float(5 - x),
            "max_drawdown": -0.05 * float(6 - x),
            "trade_count": float(x),
        }

    kwargs = dict(
        space={"x": [1, 2, 3, 4, 5]},
        evaluate=evaluate,
        objectives=[Objective("sharpe", "maximize"), Objective("turnover_total", "minimize")],
        constraints=[
            Constraint(metric="turnover_total", max_value=3.0),
            Constraint(metric="trade_count", min_value=2.0),
        ],
        sampler=TPESampler(random_fraction=0.0),
        n_trials=12,
        seed=123,
        partial_period_fractions=[0.3, 0.6, 1.0],
    )

    out1 = optimize(output_dir=tmp_path / "r1", **kwargs)
    out2 = optimize(output_dir=tmp_path / "r2", **kwargs)

    trace1 = (tmp_path / "r1" / "trials.jsonl").read_text()
    trace2 = (tmp_path / "r2" / "trials.jsonl").read_text()
    assert trace1 == trace2
    assert out1["pareto_count"] >= 1
    assert any(json.loads(line)["stopped_early"] for line in trace1.splitlines())


def test_run_strategy_optimization_persists_artifacts(monkeypatch, tmp_path) -> None:
    cache_runner.BACKTEST_OUTPUT_DIR = tmp_path / "outputs"

    def fake_combo(payload: dict[str, object]) -> dict[str, object]:
        idx = int(payload["combo_index"])
        frac = float(payload.get("partial_period_fraction", 1.0))
        return {
            "sharpe": float(1.0 + idx) * frac,
            "turnover_total": float(idx + 1),
            "max_drawdown": -0.1 + idx * 0.01,
            "trade_count": float(10 + idx),
            "total_return": 0.01,
        }

    monkeypatch.setattr(cache_runner, "_execute_sweep_combo", fake_combo)

    output = cache_runner.run_strategy_optimization(
        tickers=["AAA"],
        start_date=date(2024, 1, 1),
        end_date=date(2024, 1, 5),
        cache_root=tmp_path / "cache",
        entry_grid={"ts_momentum": [{}]},
        exit_grid={"none": [{}]},
        core_grid={"lookback_days": [20], "skip_days": [5], "costs_bps": [1.0, 2.0]},
        seed=7,
        n_trials=8,
        sampler_name="tpe",
        max_turnover=3.0,
        min_trades=1.0,
        partial_period_fractions=[0.5, 1.0],
    )

    run_dirs = sorted((tmp_path / "outputs").glob("tsmom_optimize_*"))
    assert len(run_dirs) == 1
    run_dir = run_dirs[0]
    assert (run_dir / "trials.jsonl").exists()
    assert (run_dir / "pareto_frontier.json").exists()
    assert "Optimization complete" in output

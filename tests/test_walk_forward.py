from __future__ import annotations

import csv
from datetime import date

from src.backtesting import cache_runner
from src.backtesting.walk_forward import (
    WalkForwardResult,
    build_walk_forward_folds,
    persist_walk_forward_outputs,
    run_walk_forward_optimization,
)


def test_build_walk_forward_folds_boundaries_and_no_leakage() -> None:
    folds = build_walk_forward_folds(
        total_bars=30,
        train_bars=10,
        validation_bars=5,
        test_bars=5,
        step_bars=5,
    )

    assert len(folds) == 3
    first = folds[0]
    assert (first.train_start, first.train_end) == (0, 10)
    assert (first.validation_start, first.validation_end) == (10, 15)
    assert (first.test_start, first.test_end) == (15, 20)

    for fold in folds:
        assert fold.train_end <= fold.validation_start
        assert fold.validation_end <= fold.test_start
        assert fold.train_end <= fold.test_start


def test_walk_forward_aggregation_deterministic_tie_break() -> None:
    folds = build_walk_forward_folds(
        total_bars=20,
        train_bars=6,
        validation_bars=4,
        test_bars=4,
        step_bars=4,
    )
    candidates = [{"name": "z"}, {"name": "a"}]

    def fake_eval(candidate: dict[str, object], start: int, end: int) -> dict[str, object]:
        length = end - start
        return {
            "metrics": {"sharpe": 1.0, "total_return": float(length)},
            "equity": [{"timestamp": f"t{i}", "equity": float(i + 1)} for i in range(length)],
        }

    result = run_walk_forward_optimization(
        folds=folds,
        parameter_candidates=candidates,
        evaluate_segment=fake_eval,
        score_metric="sharpe",
    )

    assert all(fold["selected_params"]["name"] == "a" for fold in result.folds)
    assert result.aggregate_metrics["sharpe_mean"] == 1.0
    assert result.stability["unique_selected_params"] == 1


def test_persist_walk_forward_outputs(tmp_path) -> None:
    result = WalkForwardResult(
        folds=[
            {
                "fold_id": 0,
                "indices": {"train": [0, 5], "validation": [5, 8], "test": [8, 10]},
                "selected_params": {"lookback_days": 20},
                "validation_score": 1.2,
                "oos_metrics": {"sharpe": 0.9},
                "oos_equity": [
                    {"timestamp": "2024-01-01T00:00:00", "equity": 1.0},
                    {"timestamp": "2024-01-01T00:01:00", "equity": 1.01},
                ],
                "diagnostics": [{"params": {"lookback_days": 20}, "validation_score": 1.2}],
            }
        ],
        aggregate_metrics={"sharpe_mean": 0.9},
        stability={"unique_selected_params": 1},
    )
    run_dir = tmp_path / "wf"
    persist_walk_forward_outputs(run_dir=run_dir, result=result)

    assert (run_dir / "aggregate_metrics.json").exists()
    assert (run_dir / "stability.json").exists()
    fold_dir = run_dir / "folds" / "fold_000"
    assert (fold_dir / "selected_params.json").exists()
    rows = list(csv.DictReader((fold_dir / "oos_equity.csv").read_text().splitlines()))
    assert len(rows) == 2


def test_run_walk_forward_backtest_persists_fold_artifacts(monkeypatch, tmp_path) -> None:
    cache_runner.BACKTEST_OUTPUT_DIR = tmp_path / "outputs"

    class _Arrays:
        def __init__(self) -> None:
            import numpy as np
            from datetime import datetime, timedelta

            self.close_prices = np.ones((24, 1), dtype=float)
            self.missing_mask = np.zeros((24, 1), dtype=bool)
            t0 = datetime(2024, 1, 1)
            self.date_index = [t0 + timedelta(minutes=i) for i in range(24)]

    monkeypatch.setattr(cache_runner, "load_backtest_engine_arrays", lambda **kwargs: _Arrays())

    class _Result:
        def __init__(self, n: int) -> None:
            import numpy as np

            self.metrics = {"sharpe": 1.0, "total_return": 0.01 * n}
            self.equity_curve = np.linspace(1.0, 1.0 + (0.01 * n), num=n)

    monkeypatch.setattr(cache_runner, "build_targets", lambda **kwargs: kwargs["close_prices"] * 0.0)
    monkeypatch.setattr(cache_runner, "backtest_vectorized", lambda **kwargs: _Result(kwargs["prices"].shape[0]))

    output = cache_runner.run_walk_forward_backtest(
        tickers=["AAA"],
        start_date=date(2024, 1, 1),
        end_date=date(2024, 1, 2),
        cache_root=tmp_path / "cache",
        entry_grid={"ts_momentum": [{"lookback_days": 5, "skip_days": 1}]},
        exit_grid={"none": [{}]},
        core_grid={"lookback_days": [5], "skip_days": [1], "costs_bps": [1.0]},
        train_bars=8,
        validation_bars=4,
        test_bars=4,
        step_bars=4,
    )

    assert "Walk-forward complete" in output
    run_dirs = list((tmp_path / "outputs").glob("tsmom_walk_forward_*"))
    assert len(run_dirs) == 1
    fold_dirs = sorted((run_dirs[0] / "folds").glob("fold_*"))
    assert fold_dirs
    assert (fold_dirs[0] / "selected_params.json").exists()
    assert (fold_dirs[0] / "oos_metrics.json").exists()
    assert (fold_dirs[0] / "oos_equity.csv").exists()
    assert (fold_dirs[0] / "diagnostics.json").exists()

from __future__ import annotations

import csv
from datetime import date

import pytest

from src.backtesting import cache_runner
from src.backtesting.walk_forward import (
    WalkForwardResult,
    build_cpcv_walk_forward_folds,
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

    assert len(folds) == 2
    first = folds[0]
    assert (first.train_start, first.train_end) == (0, 10)
    assert (first.validation_start, first.validation_end) == (11, 16)
    assert (first.test_start, first.test_end) == (17, 22)

    for fold in folds:
        assert fold.train_end < fold.validation_start
        assert fold.validation_end < fold.test_start
        assert fold.train_end < fold.test_start
        assert all(fold.leakage_checks.values())


def test_build_walk_forward_folds_purge_and_embargo_boundaries() -> None:
    folds = build_walk_forward_folds(
        total_bars=32,
        train_bars=10,
        validation_bars=5,
        test_bars=4,
        step_bars=4,
        purge_window_bars=2,
        embargo_window_bars=3,
        label_horizon_bars=1,
    )

    assert len(folds) == 3
    first = folds[0]
    assert (first.train_start, first.train_end) == (0, 10)
    assert (first.validation_start, first.validation_end) == (12, 17)
    assert (first.test_start, first.test_end) == (20, 24)
    assert first.excluded_ranges == {
        "train_validation": [10, 12],
        "validation_test": [17, 20],
    }
    assert all(first.leakage_checks.values())


def test_build_walk_forward_folds_raises_on_overlap_constraints() -> None:
    with pytest.raises(ValueError, match="label_horizon_bars"):
        build_walk_forward_folds(
            total_bars=30,
            train_bars=10,
            validation_bars=5,
            test_bars=5,
            label_horizon_bars=0,
        )


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
    assert "excluded_ranges" in result.folds[0]
    assert "leakage_checks" in result.folds[0]


def test_walk_forward_nested_optimization_uses_inner_scores() -> None:
    outer_folds = build_walk_forward_folds(
        total_bars=40,
        train_bars=10,
        validation_bars=5,
        test_bars=5,
        step_bars=10,
        purge_window_bars=1,
        embargo_window_bars=1,
    )

    inner_folds = {
        outer_folds[0].fold_id: [
            build_walk_forward_folds(
                total_bars=20,
                train_bars=8,
                validation_bars=4,
                test_bars=1,
                step_bars=4,
                purge_window_bars=1,
                embargo_window_bars=0,
            )[0]
        ]
    }

    candidates = [{"name": "a"}, {"name": "b"}]

    def fake_eval(candidate: dict[str, object], start: int, end: int) -> dict[str, object]:
        score = 2.0 if candidate["name"] == "b" and end <= 13 else 1.0
        return {"metrics": {"sharpe": score}, "equity": []}

    result = run_walk_forward_optimization(
        folds=outer_folds[:1],
        parameter_candidates=candidates,
        evaluate_segment=fake_eval,
        nested_inner_folds=inner_folds,
    )

    assert result.folds[0]["selected_params"]["name"] == "b"
    assert "inner_diagnostics" in result.folds[0]["diagnostics"][0]


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

            self.close_prices = np.ones((40, 1), dtype=float)
            self.missing_mask = np.zeros((40, 1), dtype=bool)
            t0 = datetime(2024, 1, 1)
            self.date_index = [t0 + timedelta(minutes=i) for i in range(40)]

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
    assert (run_dirs[0] / "split_metadata.json").exists()
    assert (run_dirs[0] / "fold_boundaries.csv").exists()
    summary = (run_dirs[0] / "fold_summary.json").read_text()
    assert "excluded_ranges" in summary
    assert "leakage_checks" in summary
    report_path = run_dirs[0] / "report.txt"
    assert report_path.exists()
    assert "Walk-Forward Backtest Report" in report_path.read_text()
    assert "Report:" in output


def test_run_walk_forward_backtest_supports_fraction_windows(monkeypatch, tmp_path) -> None:
    cache_runner.BACKTEST_OUTPUT_DIR = tmp_path / "outputs"

    class _Arrays:
        def __init__(self) -> None:
            import numpy as np
            from datetime import datetime, timedelta

            self.close_prices = np.ones((90, 1), dtype=float)
            self.missing_mask = np.zeros((90, 1), dtype=bool)
            t0 = datetime(2024, 1, 1)
            self.date_index = [t0 + timedelta(minutes=i) for i in range(90)]

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
        train_fraction=0.6,
        validation_fraction=0.2,
        test_fraction=0.2,
        step_fraction=0.2,
        nested_optimization=True,
    )

    assert "Walk-forward complete" in output
    run_dirs = list((tmp_path / "outputs").glob("tsmom_walk_forward_*"))
    assert len(run_dirs) == 1


def test_run_walk_forward_backtest_handles_integer_timestamps(monkeypatch, tmp_path) -> None:
    cache_runner.BACKTEST_OUTPUT_DIR = tmp_path / "outputs"

    class _Arrays:
        def __init__(self) -> None:
            import numpy as np

            self.close_prices = np.ones((60, 1), dtype=float)
            self.missing_mask = np.zeros((60, 1), dtype=bool)
            base_ms = 1704067200000
            self.date_index = np.array([base_ms + (i * 60_000) for i in range(60)], dtype=np.int64)

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
        train_fraction=0.6,
        validation_fraction=0.2,
        test_fraction=0.2,
        step_fraction=0.2,
    )

    assert "Walk-forward complete" in output


def test_run_walk_forward_backtest_manifest_includes_lineage_parent(monkeypatch, tmp_path) -> None:
    cache_runner.BACKTEST_OUTPUT_DIR = tmp_path / "outputs"

    class _Arrays:
        def __init__(self) -> None:
            import numpy as np
            from datetime import datetime, timedelta

            self.close_prices = np.ones((40, 1), dtype=float)
            self.missing_mask = np.zeros((40, 1), dtype=bool)
            t0 = datetime(2024, 1, 1)
            self.date_index = [t0 + timedelta(minutes=i) for i in range(40)]

    monkeypatch.setattr(cache_runner, "load_backtest_engine_arrays", lambda **kwargs: _Arrays())

    class _Result:
        def __init__(self, n: int) -> None:
            import numpy as np

            self.metrics = {"sharpe": 1.0, "total_return": 0.01 * n}
            self.equity_curve = np.linspace(1.0, 1.0 + (0.01 * n), num=n)

    monkeypatch.setattr(cache_runner, "build_targets", lambda **kwargs: kwargs["close_prices"] * 0.0)
    monkeypatch.setattr(cache_runner, "backtest_vectorized", lambda **kwargs: _Result(kwargs["prices"].shape[0]))

    cache_runner.run_walk_forward_backtest(
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
        lineage_parent_manifest="/tmp/ancestor-manifest.json",
    )

    run_dir = sorted((tmp_path / "outputs").glob("tsmom_walk_forward_*"))[-1]
    manifest = __import__("json").loads((run_dir / "manifest.json").read_text())
    assert manifest["lineage"]["lineage_parent_manifest"] == "/tmp/ancestor-manifest.json"


def test_build_cpcv_walk_forward_folds_generates_combinatorial_partitions() -> None:
    folds = build_cpcv_walk_forward_folds(total_bars=60, n_groups=6, n_test_groups=2)
    assert len(folds) == 30
    assert all(fold.test_start < fold.test_end for fold in folds)


def test_walk_forward_stability_includes_fold_reuse() -> None:
    folds = build_walk_forward_folds(total_bars=30, train_bars=8, validation_bars=4, test_bars=4, step_bars=4)

    def fake_eval(candidate: dict[str, object], start: int, end: int) -> dict[str, object]:
        return {"metrics": {"sharpe": float(end - start)}, "equity": []}

    result = run_walk_forward_optimization(
        folds=folds,
        parameter_candidates=[{"name": "a"}],
        evaluate_segment=fake_eval,
    )
    reuse = result.stability.get("fold_reuse", {})
    assert "train_avg_reuse" in reuse
    assert "validation_avg_reuse" in reuse
    assert "test_avg_reuse" in reuse

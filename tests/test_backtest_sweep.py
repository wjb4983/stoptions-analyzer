from __future__ import annotations

from datetime import date

from src.backtesting import cache_runner


def test_generate_sweep_combinations_skips_invalid() -> None:
    entry_grid = {
        "ts_momentum": [{"lookback_days": 20, "skip_days": 5}],
        "ma_trend": [{"ma_window": 10}],
    }
    exit_grid = {
        "none": [{}],
        "max_hold": [{"max_hold_bars": 0}],
    }
    core_grid = {
        "lookback_days": [20],
        "skip_days": [5],
        "costs_bps": [5.0],
    }

    valid, invalid = cache_runner.generate_sweep_combinations(
        entry_grid=entry_grid,
        exit_grid=exit_grid,
        core_grid=core_grid,
    )

    assert valid
    assert invalid
    assert any(row["combo"]["exit_signal"] == "max_hold" for row in invalid)


def test_run_parameter_sweep_ranks_by_sharpe_and_is_reproducible(tmp_path) -> None:
    entry_grid = {"ts_momentum": [{"lookback_days": 20, "skip_days": 5}, {"lookback_days": 30, "skip_days": 3}]}
    exit_grid = {"none": [{}], "momentum_flip": [{"lookback_days": 20, "skip_days": 5}]}
    core_grid = {"lookback_days": [20], "skip_days": [5], "costs_bps": [1.0, 3.0]}

    cache_runner.BACKTEST_OUTPUT_DIR = tmp_path / "outputs"

    def fake_eval(payload: dict[str, object]) -> dict[str, object]:
        combo_index = int(payload["combo_index"])
        sharpe = 10.0 - combo_index
        return {
            "entry_signal": payload["entry_signal"],
            "entry_signal_params": "{}",
            "exit_signal": payload["exit_signal"],
            "exit_signal_params": "{}",
            "lookback_days": int(payload["lookback_days"]),
            "skip_days": int(payload["skip_days"]),
            "costs_bps": float(payload["costs_bps"]),
            "total_return": sharpe / 100.0,
            "sharpe": sharpe,
            "max_drawdown": -0.1,
            "volatility": 0.2,
            "turnover_total": 1.0,
            "cost_total": 0.01,
        }

    out1 = cache_runner.run_parameter_sweep(
        tickers=["AAA"],
        start_date=date(2024, 1, 1),
        end_date=date(2024, 1, 2),
        cache_root=tmp_path / "cache",
        entry_grid=entry_grid,
        exit_grid=exit_grid,
        core_grid=core_grid,
        seed=7,
        top_n=3,
        evaluator=fake_eval,
    )
    out2 = cache_runner.run_parameter_sweep(
        tickers=["AAA"],
        start_date=date(2024, 1, 1),
        end_date=date(2024, 1, 2),
        cache_root=tmp_path / "cache",
        entry_grid=entry_grid,
        exit_grid=exit_grid,
        core_grid=core_grid,
        seed=7,
        top_n=3,
        evaluator=fake_eval,
    )

    run_dirs = sorted((tmp_path / "outputs").glob("tsmom_sweep_*"))
    assert len(run_dirs) == 2

    leaderboard1 = (run_dirs[0] / "leaderboard.csv").read_text()
    leaderboard2 = (run_dirs[1] / "leaderboard.csv").read_text()
    assert leaderboard1 == leaderboard2
    assert "Sweep complete" in out1
    assert "Sweep complete" in out2

    import csv
    rows = list(csv.DictReader(leaderboard1.splitlines()))
    assert len(rows) > 1
    assert float(rows[0]["sharpe"]) >= float(rows[1]["sharpe"])

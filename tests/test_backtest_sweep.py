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
            "cagr": sharpe / 120.0,
            "max_drawdown": -0.1,
            "calmar": sharpe / 12.0,
            "volatility": 0.2,
            "sortino": sharpe / 2.0,
            "downside_deviation": 0.1,
            "hit_rate": 0.6,
            "profit_factor": 1.5,
            "exposure_time": 0.8,
            "turnover_adjusted_return": 0.03,
            "rolling_sharpe_mean": sharpe - 1.0,
            "rolling_drawdown_worst": -0.08,
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
    assert "cagr" in rows[0]
    assert "calmar" in rows[0]
    assert "turnover_adjusted_return" in rows[0]


def test_run_multi_signal_backtest_ranks_and_persists(monkeypatch, tmp_path) -> None:
    cache_runner.BACKTEST_OUTPUT_DIR = tmp_path / "outputs"

    def fake_single(**kwargs):
        entry = kwargs["entry_signal"]
        exit_ = kwargs["exit_signal"]
        stamp = f"{entry}_{exit_}"
        run_dir = tmp_path / stamp
        run_dir.mkdir(parents=True, exist_ok=True)
        sharpe = {
            ("ts_momentum", "none"): 1.5,
            ("ma_trend", "none"): 0.8,
        }.get((entry, exit_), 0.1)
        metrics = [
            {"metric": "total_return", "value": sharpe / 10.0},
            {"metric": "sharpe", "value": sharpe},
            {"metric": "cagr", "value": 0.08},
            {"metric": "max_drawdown", "value": -0.2},
            {"metric": "calmar", "value": 0.4},
            {"metric": "volatility", "value": 0.3},
            {"metric": "sortino", "value": 1.1},
            {"metric": "hit_rate", "value": 0.55},
            {"metric": "profit_factor", "value": 1.4},
            {"metric": "turnover_adjusted_return", "value": 0.04},
            {"metric": "rolling_sharpe_mean", "value": 0.8},
            {"metric": "rolling_drawdown_worst", "value": -0.12},
            {"metric": "turnover_total", "value": 1.2},
            {"metric": "cost_total", "value": 0.02},
        ]
        (run_dir / "metrics.json").write_text(__import__("json").dumps(metrics))
        return f"report\n\nSaved outputs to: {run_dir}"

    monkeypatch.setattr(cache_runner, "run_time_series_momentum_backtest", fake_single)

    output = cache_runner.run_multi_signal_backtest(
        tickers=["AAA"],
        start_date=date(2024, 1, 1),
        end_date=date(2024, 1, 2),
        cache_root=tmp_path / "cache",
        lookback_days=20,
        skip_days=5,
        costs_bps=5.0,
        entry_signals=["ts_momentum", "ma_trend"],
        exit_signals=["none"],
    )

    assert "Ranked combinations" in output
    assert "#1 entry=ts_momentum exit=none" in output

    run_dirs = list((tmp_path / "outputs").glob("tsmom_multi_signal_*"))
    assert len(run_dirs) == 1
    leaderboard = (run_dirs[0] / "leaderboard.csv").read_text()
    assert "entry_signal,exit_signal,total_return,sharpe,cagr,max_drawdown,calmar" in leaderboard


def test_run_multi_signal_backtest_applies_conservative_runtime_params(monkeypatch, tmp_path) -> None:
    cache_runner.BACKTEST_OUTPUT_DIR = tmp_path / "outputs"
    captured: list[dict[str, object]] = []

    def fake_single(**kwargs):
        captured.append(kwargs)
        run_dir = tmp_path / f"run_{len(captured)}"
        run_dir.mkdir(parents=True, exist_ok=True)
        metrics = [
            {"metric": "total_return", "value": 0.01},
            {"metric": "sharpe", "value": 1.0},
            {"metric": "cagr", "value": 0.05},
            {"metric": "max_drawdown", "value": -0.1},
            {"metric": "calmar", "value": 0.5},
            {"metric": "volatility", "value": 0.2},
            {"metric": "sortino", "value": 0.9},
            {"metric": "hit_rate", "value": 0.52},
            {"metric": "profit_factor", "value": 1.2},
            {"metric": "turnover_adjusted_return", "value": 0.02},
            {"metric": "rolling_sharpe_mean", "value": 0.7},
            {"metric": "rolling_drawdown_worst", "value": -0.08},
            {"metric": "turnover_total", "value": 1.0},
            {"metric": "cost_total", "value": 0.01},
        ]
        (run_dir / "metrics.json").write_text(__import__("json").dumps(metrics))
        return f"report\n\nSaved outputs to: {run_dir}"

    monkeypatch.setattr(cache_runner, "run_time_series_momentum_backtest", fake_single)

    cache_runner.run_multi_signal_backtest(
        tickers=["AAA"],
        start_date=date(2024, 1, 1),
        end_date=date(2024, 1, 2),
        cache_root=tmp_path / "cache",
        lookback_days=20,
        skip_days=5,
        costs_bps=5.0,
        entry_signals=["ts_momentum", "ma_trend"],
        exit_signals=["momentum_flip"],
    )

    assert captured
    for call in captured:
        assert int(call["signal_rebalance_interval"]) == 390
    ts_call = next(call for call in captured if call["entry_signal"] == "ts_momentum")
    assert ts_call["entry_signal_params"]["long_only"] is True
    assert float(ts_call["entry_signal_params"]["min_abs_return"]) == 0.01


def test_apply_discrete_bet_sizing_rounds_down_shares() -> None:
    prices = __import__("numpy").array([[101.0]], dtype=float)
    signals = __import__("numpy").array([[1.0]], dtype=float)
    sized = cache_runner._apply_discrete_bet_sizing(
        signals=signals,
        prices=prices,
        starting_capital=1000.0,
        bet_fraction=0.5,
    )
    # 50% budget => $500, at $101/share -> floor(4) shares => $404 exposure => 0.404 weight
    assert float(sized[0, 0]) == 404.0 / 1000.0

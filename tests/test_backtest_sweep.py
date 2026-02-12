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


def test_run_parameter_sweep_persists_manifest_and_reproducible_fingerprint(tmp_path) -> None:
    entry_grid = {"ts_momentum": [{"lookback_days": 20, "skip_days": 5}, {"lookback_days": 30, "skip_days": 3}]}
    exit_grid = {"none": [{}]}
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

    cache_runner.run_parameter_sweep(
        tickers=["AAA"],
        start_date=date(2024, 1, 1),
        end_date=date(2024, 1, 2),
        cache_root=tmp_path / "cache",
        entry_grid=entry_grid,
        exit_grid=exit_grid,
        core_grid=core_grid,
        seed=123,
        top_n=3,
        evaluator=fake_eval,
    )
    cache_runner.run_parameter_sweep(
        tickers=["AAA"],
        start_date=date(2024, 1, 1),
        end_date=date(2024, 1, 2),
        cache_root=tmp_path / "cache",
        entry_grid=entry_grid,
        exit_grid=exit_grid,
        core_grid=core_grid,
        seed=123,
        top_n=3,
        evaluator=fake_eval,
    )

    run_dirs = sorted((tmp_path / "outputs").glob("tsmom_sweep_*"))
    assert len(run_dirs) == 2

    manifest_one = __import__("json").loads((run_dirs[0] / "manifest.json").read_text())
    manifest_two = __import__("json").loads((run_dirs[1] / "manifest.json").read_text())

    for manifest in (manifest_one, manifest_two):
        assert manifest["metric_schema_version"] == cache_runner.CANONICAL_METRIC_SCHEMA_VERSION
        assert manifest["code_version"]
        assert manifest["parameters"]["core_grid"] == core_grid
        assert manifest["random_seed"] == 123
        assert "data_snapshot_identifiers" in manifest
        assert "environment" in manifest

    assert manifest_one["reproducibility_fingerprint"] == manifest_two["reproducibility_fingerprint"]

    index_rows = (tmp_path / "outputs" / "experiment_index.jsonl").read_text().strip().splitlines()
    assert len(index_rows) == 2


def test_run_parameter_sweep_resume_reuses_checkpointed_queue(tmp_path) -> None:
    entry_grid = {"ts_momentum": [{"lookback_days": 20, "skip_days": 5}]}
    exit_grid = {"none": [{}]}
    core_grid = {"lookback_days": [20, 30], "skip_days": [5], "costs_bps": [1.0, 2.0]}
    cache_runner.BACKTEST_OUTPUT_DIR = tmp_path / "outputs"

    state_path = tmp_path / "outputs" / "resume_state.json"
    job_id = cache_runner._stable_fingerprint(
        {
            "tickers": ["AAA"],
            "start_date": date(2024, 1, 1).isoformat(),
            "end_date": date(2024, 1, 2).isoformat(),
            "entry_grid": entry_grid,
            "exit_grid": exit_grid,
            "core_grid": core_grid,
            "seed": 11,
        }
    )
    cache_runner._persist_resume_state(
        state_path=state_path,
        job_id=job_id,
        seed=11,
        queued_indices=[1, 3],
        completed_rows=[
            {
                "entry_signal": "ts_momentum",
                "entry_signal_params": "{}",
                "exit_signal": "none",
                "exit_signal_params": "{}",
                "lookback_days": 20,
                "skip_days": 5,
                "costs_bps": 1.0,
                "total_return": 0.1,
                "sharpe": 1.0,
                "cagr": 0.01,
                "max_drawdown": -0.1,
                "calmar": 0.1,
                "volatility": 0.2,
                "sortino": 0.5,
                "downside_deviation": 0.1,
                "hit_rate": 0.6,
                "profit_factor": 1.2,
                "exposure_time": 0.8,
                "turnover_adjusted_return": 0.03,
                "rolling_sharpe_mean": 0.8,
                "rolling_drawdown_worst": -0.05,
                "turnover_total": 1.0,
                "trade_count": 5.0,
                "cost_total": 0.01,
            }
        ],
        invalid_rows=[],
        errors=[],
        retry_counts={},
    )

    seen_indices: list[int] = []

    def fake_eval(payload: dict[str, object]) -> dict[str, object]:
        idx = int(payload["combo_index"])
        seen_indices.append(idx)
        return {
            "entry_signal": payload["entry_signal"],
            "entry_signal_params": "{}",
            "exit_signal": payload["exit_signal"],
            "exit_signal_params": "{}",
            "lookback_days": int(payload["lookback_days"]),
            "skip_days": int(payload["skip_days"]),
            "costs_bps": float(payload["costs_bps"]),
            "total_return": 0.2,
            "sharpe": 2.0 + idx,
            "cagr": 0.02,
            "max_drawdown": -0.1,
            "calmar": 0.2,
            "volatility": 0.2,
            "sortino": 0.6,
            "downside_deviation": 0.1,
            "hit_rate": 0.6,
            "profit_factor": 1.3,
            "exposure_time": 0.8,
            "turnover_adjusted_return": 0.03,
            "rolling_sharpe_mean": 1.1,
            "rolling_drawdown_worst": -0.05,
            "turnover_total": 1.0,
            "trade_count": 5.0,
            "cost_total": 0.01,
        }

    cache_runner.run_parameter_sweep(
        tickers=["AAA"],
        start_date=date(2024, 1, 1),
        end_date=date(2024, 1, 2),
        cache_root=tmp_path / "cache",
        entry_grid=entry_grid,
        exit_grid=exit_grid,
        core_grid=core_grid,
        seed=11,
        evaluator=fake_eval,
        resume_state_path=state_path,
    )

    assert sorted(seen_indices) == [1, 3]
    run_dirs = sorted((tmp_path / "outputs").glob("tsmom_sweep_*"))
    manifest = __import__("json").loads((run_dirs[-1] / "manifest.json").read_text())
    assert manifest["lineage"]["resumed_from"] == str(state_path)


def test_run_parameter_sweep_retries_failures_deterministically(tmp_path) -> None:
    entry_grid = {"ts_momentum": [{"lookback_days": 20, "skip_days": 5}]}
    exit_grid = {"none": [{}]}
    core_grid = {"lookback_days": [20], "skip_days": [5], "costs_bps": [1.0, 2.0, 3.0]}
    cache_runner.BACKTEST_OUTPUT_DIR = tmp_path / "outputs"

    attempt_count: dict[int, int] = {}

    def flaky_eval(payload: dict[str, object]) -> dict[str, object]:
        idx = int(payload["combo_index"])
        attempt_count[idx] = attempt_count.get(idx, 0) + 1
        if idx == 1 and attempt_count[idx] == 1:
            raise RuntimeError("transient worker failure")
        worker_seed = int(payload["worker_seed"])
        return {
            "entry_signal": payload["entry_signal"],
            "entry_signal_params": "{}",
            "exit_signal": payload["exit_signal"],
            "exit_signal_params": "{}",
            "lookback_days": int(payload["lookback_days"]),
            "skip_days": int(payload["skip_days"]),
            "costs_bps": float(payload["costs_bps"]),
            "total_return": worker_seed / 10_000.0,
            "sharpe": float(worker_seed),
            "cagr": 0.01,
            "max_drawdown": -0.1,
            "calmar": 0.2,
            "volatility": 0.2,
            "sortino": 0.6,
            "downside_deviation": 0.1,
            "hit_rate": 0.6,
            "profit_factor": 1.3,
            "exposure_time": 0.8,
            "turnover_adjusted_return": 0.03,
            "rolling_sharpe_mean": 1.1,
            "rolling_drawdown_worst": -0.05,
            "turnover_total": 1.0,
            "trade_count": 5.0,
            "cost_total": 0.01,
        }

    common = dict(
        tickers=["AAA"],
        start_date=date(2024, 1, 1),
        end_date=date(2024, 1, 2),
        cache_root=tmp_path / "cache",
        entry_grid=entry_grid,
        exit_grid=exit_grid,
        core_grid=core_grid,
        seed=5,
        evaluator=flaky_eval,
        retry_policy=cache_runner.SweepRetryPolicy(max_attempts=3, stale_worker_timeout_seconds=5.0),
    )
    cache_runner.run_parameter_sweep(**common)
    leaderboard_one = sorted((tmp_path / "outputs").glob("tsmom_sweep_*"))[-1] / "leaderboard.csv"
    first_text = leaderboard_one.read_text()

    attempt_count.clear()
    cache_runner.run_parameter_sweep(**common)
    leaderboard_two = sorted((tmp_path / "outputs").glob("tsmom_sweep_*"))[-1] / "leaderboard.csv"
    second_text = leaderboard_two.read_text()

    assert first_text == second_text

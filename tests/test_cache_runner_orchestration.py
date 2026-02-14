from __future__ import annotations

from datetime import date
from pathlib import Path

import numpy as np
import pytest

from src.backtesting import cache_runner


def test_resolve_benchmark_selection_defaults_for_empty_inputs() -> None:
    assert cache_runner._resolve_benchmark_selection(None) == list(cache_runner.DEFAULT_BENCHMARK_SELECTION)
    assert cache_runner._resolve_benchmark_selection([]) == list(cache_runner.DEFAULT_BENCHMARK_SELECTION)


def test_resolve_benchmark_selection_normalizes_deduplicates_and_filters_invalid() -> None:
    selected = [
        " Buy Hold ",
        "equal-weight momentum",
        "EQUAL_WEIGHT_MOMENTUM",
        "volatility-parity",
        "unknown_benchmark",
    ]

    resolved = cache_runner._resolve_benchmark_selection(selected)

    assert resolved == [
        cache_runner.BENCHMARK_BUY_HOLD,
        cache_runner.BENCHMARK_EQUAL_WEIGHT_MOMENTUM,
        cache_runner.BENCHMARK_VOLATILITY_PARITY,
    ]


def test_promotion_required_checks_are_cumulative_by_state() -> None:
    prior_checks: set[str] = set()
    for state in cache_runner.PROMOTION_STATES:
        checks = cache_runner.PROMOTION_REQUIRED_CHECKS[state]
        assert checks
        assert prior_checks.issubset(set(checks))
        prior_checks = set(checks)

    assert "approval" in cache_runner.PROMOTION_REQUIRED_CHECKS["production"]
    assert "experiment_id" in cache_runner.PROMOTION_REQUIRED_CHECKS["shadow"]


def test_main_run_command_parses_and_propagates(monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]) -> None:
    captured: dict[str, object] = {}

    def fake_run_time_series_momentum_backtest(**kwargs):
        captured.update(kwargs)
        return "ok-run"

    monkeypatch.setattr(cache_runner, "run_time_series_momentum_backtest", fake_run_time_series_momentum_backtest)
    monkeypatch.setattr(
        "sys.argv",
        [
            "cache_runner.py",
            "run",
            "--tickers",
            "AAA, BBB",
            "--start-date",
            "2024-01-01",
            "--end-date",
            "2024-01-31",
            "--cache-root",
            "/tmp/cache-root",
            "--lookback-days",
            "30",
            "--skip-days",
            "2",
            "--costs-bps",
            "1.5",
            "--entry-signal-params",
            '{"lookback_days": 15}',
            "--exit-signal-params",
            '{"max_hold_bars": 5}',
            "--capacity-aum-scales",
            "[1000000,2000000]",
        ],
    )

    cache_runner.main()
    out = capsys.readouterr().out

    assert "ok-run" in out
    assert captured["tickers"] == ["AAA", "BBB"]
    assert captured["start_date"] == date(2024, 1, 1)
    assert captured["end_date"] == date(2024, 1, 31)
    assert captured["cache_root"] == Path("/tmp/cache-root")
    assert captured["lookback_days"] == 30
    assert captured["skip_days"] == 2
    assert captured["costs_bps"] == 1.5
    assert captured["entry_signal_params"] == {"lookback_days": 15}
    assert captured["exit_signal_params"] == {"max_hold_bars": 5}
    assert captured["capacity_aum_scales"] == [1000000.0, 2000000.0]


def test_main_sweep_command_parses_and_propagates(monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]) -> None:
    captured: dict[str, object] = {}

    def fake_run_parameter_sweep(**kwargs):
        captured.update(kwargs)
        return "ok-sweep"

    monkeypatch.setattr(cache_runner, "run_parameter_sweep", fake_run_parameter_sweep)
    monkeypatch.setattr(
        "sys.argv",
        [
            "cache_runner.py",
            "sweep",
            "--tickers",
            "AAA",
            "--start-date",
            "2024-01-01",
            "--end-date",
            "2024-01-10",
            "--entry-grid",
            '{"ts_momentum":[{"lookback_days":20,"skip_days":5}]}',
            "--exit-grid",
            '{"none":[{}]}',
            "--core-grid",
            '{"lookback_days":[20],"skip_days":[5],"costs_bps":[1.0]}',
            "--top-n",
            "3",
        ],
    )

    cache_runner.main()
    out = capsys.readouterr().out

    assert "ok-sweep" in out
    assert captured["tickers"] == ["AAA"]
    assert captured["top_n"] == 3
    assert captured["entry_grid"] == {"ts_momentum": [{"lookback_days": 20, "skip_days": 5}]}


def test_main_walk_forward_and_optimize_parse_and_propagate(monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]) -> None:
    walk_calls: list[dict[str, object]] = []
    opt_calls: list[dict[str, object]] = []

    monkeypatch.setattr(cache_runner, "run_walk_forward_backtest", lambda **kwargs: walk_calls.append(kwargs) or "ok-wf")
    monkeypatch.setattr(cache_runner, "run_strategy_optimization", lambda **kwargs: opt_calls.append(kwargs) or "ok-opt")

    monkeypatch.setattr(
        "sys.argv",
        [
            "cache_runner.py",
            "walk_forward",
            "--tickers",
            "AAA",
            "--start-date",
            "2024-01-01",
            "--end-date",
            "2024-03-01",
            "--entry-grid",
            '{"ts_momentum":[{"lookback_days":20,"skip_days":5}]}',
            "--exit-grid",
            '{"none":[{}]}',
            "--core-grid",
            '{"lookback_days":[20],"skip_days":[5],"costs_bps":[1.0]}',
            "--nested-optimization",
            "--cv-scheme",
            "cpcv",
        ],
    )
    cache_runner.main()

    monkeypatch.setattr(
        "sys.argv",
        [
            "cache_runner.py",
            "optimize",
            "--tickers",
            "AAA",
            "--start-date",
            "2024-01-01",
            "--end-date",
            "2024-03-01",
            "--entry-grid",
            '{"ts_momentum":[{"lookback_days":20,"skip_days":5}]}',
            "--exit-grid",
            '{"none":[{}]}',
            "--core-grid",
            '{"lookback_days":[20],"skip_days":[5],"costs_bps":[1.0]}',
            "--n-trials",
            "7",
            "--sampler",
            "grid",
            "--partial-period-fractions",
            "[0.5,1.0]",
        ],
    )
    cache_runner.main()
    out = capsys.readouterr().out

    assert "ok-wf" in out
    assert "ok-opt" in out
    assert walk_calls and opt_calls
    assert walk_calls[0]["nested_optimization"] is True
    assert walk_calls[0]["cv_scheme"] == "cpcv"
    assert opt_calls[0]["n_trials"] == 7
    assert opt_calls[0]["sampler_name"] == "grid"
    assert opt_calls[0]["partial_period_fractions"] == [0.5, 1.0]


def test_output_parsing_helpers_handle_missing_or_malformed_artifacts(tmp_path: Path) -> None:
    run_dir = tmp_path / "run"
    run_dir.mkdir()

    assert cache_runner._load_benchmark_rows_from_run_dir(run_dir) == []
    assert cache_runner._load_stress_payload_from_run_dir(run_dir) == {}

    (run_dir / "manifest.json").write_text("{not-json}")
    (run_dir / "stress_scenarios.json").write_text("[]")

    assert cache_runner._load_benchmark_rows_from_run_dir(run_dir) == []
    assert cache_runner._load_stress_payload_from_run_dir(run_dir) == {}

    with pytest.raises(ValueError, match="Could not locate saved output directory"):
        cache_runner._extract_saved_output_dir("report-without-marker")


def test_invalid_benchmark_and_promotion_inputs_raise_user_facing_errors(capsys: pytest.CaptureFixture[str]) -> None:
    prices = np.ones((5, 2), dtype=float)
    with pytest.raises(ValueError, match="Unsupported benchmark: not_real"):
        cache_runner._build_benchmark_signals(benchmark="not_real", prices=prices, lookback_days=10, skip_days=2)

    parser = cache_runner._build_parser()
    with pytest.raises(SystemExit) as exc:
        parser.parse_args([
            "run",
            "--tickers",
            "AAA",
            "--start-date",
            "2024-01-01",
            "--end-date",
            "2024-01-02",
            "--promotion-state",
            "shadow",
            "--benchmarks",
            "buy_hold",
        ])
    assert exc.value.code == 2
    err = capsys.readouterr().err
    assert "unrecognized arguments" in err
    assert "--promotion-state" in err
    assert "--benchmarks" in err

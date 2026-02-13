# Baseline benchmark artifacts

This folder stores MVP baseline benchmark outputs for the vectorized backtest engine.

## Reproduction command

From a clean checkout:

```bash
python reports/run_baselines.py
pytest tests/test_vectorized_no_lookahead.py tests/test_backtest_invariants_and_edge_cases.py -q
```

## Included artifacts

- `baseline_configs.json`: scenario inputs and synthetic universe/horizon definition.
- `baseline_metrics_summary.csv` / `.json`: metrics and runtime per scenario.
- `equity_drawdown_plot_data.csv`: time-series data for equity and drawdown curves.
- `pass_fail_checklist.md`: pass/fail checklist against MVP acceptance gates.

## Backtest sweep artifacts

When using `PYTHONPATH=src python -m backtesting.cache_runner sweep ...`, additional sweep reports are generated under `src/data/backtest_outputs/tsmom_sweep_*` with leaderboard and top-N combo summaries.

## Slippage calibration artifacts

Calibrate dated participation slippage snapshots from historical fills:

```bash
python reports/calibrate_slippage_snapshots.py --fills reports/historical_fills.json
```

Artifacts written in `reports/`:

- `slippage_calibration_snapshots.json`: dated parameter snapshots with `stable` flags.
- `calibration_report.json`: fit error and stability diagnostics for calibration quality.

## Benchmark scorecard bundle

Run the benchmark bundle with fixed datasets and expected ranges:

```bash
python reports/benchmark_bundle.py
```

This writes a single artifact per run:

- `benchmark_scorecard.json`: robust OOS, statistical significance, execution realism, stress resilience, and reproducibility checks with a promotion gate that fails when any critical dimension fails.

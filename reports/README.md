# Baseline benchmark artifacts

This folder stores MVP baseline benchmark outputs for the vectorized backtest engine.

## Reproduction command

From a clean checkout:

```bash
python reports/run_baselines.py
python reports/benchmark_bundle.py
python reports/quality_gates.py
python reports/generate_cards.py
python reports/validate_governance_artifacts.py
pytest -m core -q
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

## Merge/deploy quality gates and cards

Evaluate no-arbitrage diagnostics and required merge/deploy gates (data quality, leakage tests, validation integrity, calibration, friction-adjusted performance, no-arbitrage surface), then generate standardized cards:

```bash
python reports/no_arb_diagnostics.py --surface reports/surface_total_variance_snapshot.npz
python reports/quality_gates.py
python reports/generate_cards.py
```

Artifacts written in `reports/`:

- `no_arb_diagnostics_report.json`
- `quality_gates_report.json`
- `model_card.json`
- `strategy_card.json`

## Nightly benchmark comparison

Compare newly generated mainline baseline metrics against the prior baseline snapshot:

```bash
cp reports/baseline_metrics_summary.json reports/prior_baseline_metrics_summary.json
python reports/run_baselines.py
python reports/nightly_benchmark_report.py
```

Artifacts written in `reports/`:

- `nightly_benchmark_comparison.json`
- `nightly_benchmark_report.md`

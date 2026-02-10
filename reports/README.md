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

# Stoptions Analyzer

"For God so loved the world, that he gave his only begotten Son, that whosoever believeth in him should not perish, but have everlasting life." — John 3:16

## Tests

Install dependencies and run pytest from the repo root:

```bash
pip install -r requirements.txt
pytest
```


## Backtesting GUI workflow

In the Backtesting page, select one or more **Entry Signals** and **Exit Signals** via checkboxes.
The app runs every entry/exit pair using the same lookback/skip/cost/date parameters, starting capital, and bet-size mode (Kelly / Half Kelly / custom %), then prints a ranked leaderboard in the Run Output panel. For each combo, a portfolio-value-over-time chart (x-axis=day) is saved in that combo output folder.

## Backtest CLI

### Single-run entry/exit signal selection

Run one backtest combo with explicit entry/exit signal definitions:

```bash
PYTHONPATH=src python -m backtesting.cache_runner run \
  --tickers AAPL,MSFT \
  --start-date 2024-01-01 \
  --end-date 2024-03-01 \
  --entry-signal ts_momentum \
  --exit-signal none
```

Breakout entry + trailing stop exit with typed JSON params:

```bash
PYTHONPATH=src python -m backtesting.cache_runner run \
  --tickers AAPL \
  --start-date 2024-01-01 \
  --end-date 2024-03-01 \
  --entry-signal breakout \
  --entry-signal-params '{"breakout_window": 55}' \
  --exit-signal trailing_stop \
  --exit-signal-params '{"trailing_stop_pct": 0.08}'
```

### Parallel parameter sweep

Run entry/exit/core parameter grids in parallel and produce ranked artifacts:

```bash
PYTHONPATH=src python -m backtesting.cache_runner sweep \
  --tickers AAPL,MSFT \
  --start-date 2024-01-01 \
  --end-date 2024-03-01 \
  --entry-grid '{"ts_momentum": [{"lookback_days": 60, "skip_days": 5}, {"lookback_days": 90, "skip_days": 10}], "breakout": [{"breakout_window": 55}]}' \
  --exit-grid '{"none": [{}], "trailing_stop": [{"trailing_stop_pct": 0.08}]}' \
  --core-grid '{"lookback_days": [60, 90], "skip_days": [5], "costs_bps": [2.5, 5.0]}' \
  --seed 42 \
  --top-n 10
```


### Execution modeling notes (MVP vs implemented)

The backtesting MVP execution assumptions and current implementation details are documented in `docs/backtest_mvp.md`, including:

- partial-fill support with participation caps and residual carry,
- latency inputs (`latency_bars`, `latency_ms`),
- queue-rank effects (`queue_rank_proxy`),
- a model-risk/assumptions table linked to execution and event-driven adapter classes, and
- deterministic vs stochastic execution-setting examples.

Sweep outputs are written under `src/data/backtest_outputs/tsmom_sweep_*` and include:
- `leaderboard.csv` / `leaderboard.json`
- `per_combo_summary.csv` / `per_combo_summary.json`
- `top_n_report.txt`
- `skipped_invalid_combos.json`
- `errors.json`

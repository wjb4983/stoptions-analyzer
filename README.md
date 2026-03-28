# Stoptions Analyzer

"For God so loved the world, that he gave his only begotten Son, that whosoever believeth in him should not perish, but have everlasting life." — John 3:16


## Docker quickstart

Use Docker to run the project without installing Python locally.

### 1) Build the image

```bash
./scripts/docker.sh build  # builds Python 3.10 + tk-enabled image
```

### 2) Run tests in Docker

```bash
./scripts/docker.sh test
```

This runs the fast **smoke** marker set by default (recommended for quick validation).

### 3) Run the full suite (optional, longer)

```bash
./scripts/docker.sh test-all
```

### 4) Run any command in Docker

```bash
./scripts/docker.sh run pytest -m core -ra
```

### 5) Open an interactive shell

```bash
./scripts/docker.sh shell
```

One-line copy/paste flow:

```bash
git clone <your-repo-url> && cd stoptions-analyzer && ./scripts/docker.sh build && ./scripts/docker.sh test
```

> Notes:
> - The container sets `PYTHONPATH=/app/src` so CLI modules work out of the box.
> - The image includes common development tools (`git`, `bash`, `curl`, `openssh-client`) plus Tk runtime libraries (`tk`) for broader day-to-day usage and UI-related test imports.
> - Running the desktop Tkinter UI (`src/main.py`) still typically requires GUI forwarding (X11/Wayland), so CLI/test workflows are the default in Docker.

## Tests

Install dependencies and run pytest from the repo root:

```bash
pip install -r requirements.txt
pytest
```

For a reproducible tiered local run that mirrors CI markers and emits artifacts:

```bash
make test-matrix
# or
./scripts/run_test_matrix.sh
```

Artifacts are written under `reports/test_matrix/<tier>/`:

- `stdout.log` (captured per-tier output)
- `junit.xml` (per-tier JUnit report)
- optional `coverage.xml` and `coverage_html/` when coverage is enabled.

Enable per-tier coverage artifacts with:

```bash
TEST_MATRIX_COVERAGE=1 ./scripts/run_test_matrix.sh
```

Expected runtime (machine-dependent):

- `smoke`: typically short (fast PR sanity checks)
- `core`: moderate (PR correctness + governance)
- `core or slow`: longest run (main/nightly-style comprehensive pass)

CI mapping:

- PR fast gate → `smoke`
- PR full correctness gate → `core`
- `main` / nightly comprehensive gate → `core or slow`


## CI test tiers and markers

CI is split into explicit tiers with clear pass criteria:

- **Smoke (`smoke`)**: fast checks on every PR.
- **Core (`core`)**: correctness + governance checks on every PR.
- **Slow (`slow`)**: longer regression checks on `main` and nightly schedule.

Run locally with markers:

```bash
pytest -m smoke -ra
pytest -m core -ra
pytest -m "core or slow" -ra
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

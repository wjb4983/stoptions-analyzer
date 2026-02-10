# Stoptions Analyzer

"For God so loved the world, that he gave his only begotten Son, that whosoever believeth in him should not perish, but have everlasting life." — John 3:16

## Tests

Install dependencies and run pytest from the repo root:

```bash
pip install -r requirements.txt
pytest
```

## Backtest CLI signal selection

Run the backtest module directly and choose entry/exit signal names:

```bash
PYTHONPATH=src python -m backtesting.cache_runner \
  --tickers AAPL,MSFT \
  --start-date 2024-01-01 \
  --end-date 2024-03-01 \
  --entry-signal ts_momentum \
  --exit-signal none
```

Breakout entry + trailing stop exit with typed JSON params:

```bash
PYTHONPATH=src python -m backtesting.cache_runner \
  --tickers AAPL \
  --start-date 2024-01-01 \
  --end-date 2024-03-01 \
  --entry-signal breakout \
  --entry-signal-params '{"breakout_window": 55}' \
  --exit-signal trailing_stop \
  --exit-signal-params '{"trailing_stop_pct": 0.08}'
```

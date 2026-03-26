#!/usr/bin/env bash
set -euo pipefail

TICKER="${1:-NVDA}"
TIMEFRAME="${2:-1D}"

PYTHONPATH=src python -m snn_bench.scripts.smoke_pipeline --ticker "$TICKER" --timeframe "$TIMEFRAME"

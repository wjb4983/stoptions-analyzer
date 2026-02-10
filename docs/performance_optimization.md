# Performance optimization and parity guide

## Hotspot profiling scope

The main hotspots now profiled are:

1. **Signal rolling computations** in time-series momentum array generation.
2. **Slippage/fee execution loops** in vectorized backtesting.
3. **Metric computations** in vectorized backtesting (`_compute_metrics`) under both modes.

Profiling entry points:

- `src.analysis.time_series.perf.profile_signal_rolling_hotspots`
- `src.backtesting.perf.profile_backtest_hotspots`

Both profile functions run in:

- `reference` mode: loop-oriented baseline.
- `optimized` mode: vectorized / numba-accelerated kernels where applicable.

## Optimization model

### Reference mode

- Keeps explicit Python loops for trade-cost and rolling-volatility calculations.
- Intended as a correctness baseline for regression/parity checks.

### Optimized mode

- Uses vectorized NumPy operations for momentum raw-score computation.
- Uses `numba.njit` kernels for remaining heavy loops:
  - rolling volatility + volatility targeting
  - BPS slippage aggregation for large universes
  - fixed-commission aggregation for large universes
- Falls back to reference behavior automatically when an execution model is not a known optimized form.

## Numerical tolerances for parity checks

Parity tests use:

- **Absolute tolerance (`atol`)**: `1e-10`
- **Relative tolerance (`rtol`)**: `1e-8`

Rationale:

- Most operations are deterministic and algebraically equivalent between paths.
- Small floating-point drift may appear in cumulative calculations (e.g., equity curve), so a relative tolerance is also enforced.

See `tests/test_optimized_parity.py` for the executable parity policy.

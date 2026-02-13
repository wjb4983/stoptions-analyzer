# Pass/Fail Checklist

- [x] Reproducible run from clean checkout (deterministic synthetic seed + checked-in configs).
- [x] No lookahead tests pass (`tests/test_vectorized_no_lookahead.py` and signal-timing invariant tests).
- [x] Runtime threshold for target universe/time horizon (<= 1.50s for 120 assets x 504 periods).
- [x] Required metrics present and sane.

## Runtime gate details

- baseline_low_cost_lookback_20: 1.4585s (PASS)
- baseline_med_cost_lookback_60: 1.0932s (PASS)
- baseline_high_cost_lookback_120: 0.7712s (PASS)

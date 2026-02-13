# Causal Validation Templates for Event-Driven Hypotheses

This folder standardizes causal validation so candidate signals are promoted only when they are likely **causal** (not just correlated).

## Required methods

Every event-driven hypothesis should run all three templates:

1. Difference-in-differences (`templates/difference_in_differences.yaml`)
2. Synthetic control (`templates/synthetic_control.yaml`)
3. Propensity score matching (`templates/propensity_score_matching.yaml`)

## Required automated checks

Each template records:

- **Pre-trend check p-value** (must not reject parallel trends)
- **Placebo test p-value** (must fail to show effect on placebo windows/events)
- **Effect t-stat** (must exceed minimum signal strength)
- **Relative attenuation** in stress/placebo variants (must remain bounded)

These fields are consumed by governance causal robustness gating in `src/backtesting/cache_runner.py`.

## Suggested workflow

1. Copy template.
2. Fill dataset snapshot, treatment definition, and event windows.
3. Run estimation notebook/script.
4. Write method outputs into the run's `governance.causal_validation.methods` payload.
5. Promote only if all required methods pass thresholds.

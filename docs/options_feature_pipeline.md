# Options Feature Pipeline

This document defines the options feature set used by the directional and volatility strategy models.

## Raw feature formulas

Given per-date/per-asset option observables:

- `IV_25P`, `IV_25C`, `IV_ATM`, `IV_10P`, `IV_10C`
- `IV_1M`, `IV_3M`, `IV_6M`
- `PutVol`, `CallVol`, `OI`, `NetGammaNotional`, `MarketCap`, `SpotReturn`

we compute:

1. **Skew**: `skew = IV_25P - IV_25C`
2. **Convexity**: `convexity = 0.5 * (IV_25P + IV_25C) - IV_ATM`
3. **Term-structure curvature**: `term_structure_curvature = IV_6M - 2 * IV_3M + IV_1M`
4. **Local surface distortion**: `|IV_10P - 2*IV_25P + IV_ATM| + |IV_ATM - 2*IV_25C + IV_10C|`
5. **Put-call flow imbalance**: `(PutVol - CallVol) / max(PutVol + CallVol, 1)`
6. **OI changes**: `(OI_t - OI_{t-1}) / max(OI_{t-1}, 1)`
7. **Gamma exposure proxy**: `NetGammaNotional / max(MarketCap, 1)`
8. **Dealer positioning proxy**: `-(gamma_exposure_proxy * (1 + SpotReturn))`
9. **Unusual volume signature**: `(PutVol + CallVol) / max((PutVol + CallVol)_{t-k}, 1)` where `k ~= rolling_window / 2`.

## Normalization layers

For **every** raw options feature above we also compute:

- **Rolling z-score** over configurable window `W`:
  - `z_t = (x_t - mean(x_{t-W+1:t})) / std(x_{t-W+1:t})`
- **Cross-sectional rank** per date across assets scaled to `[0, 1]`.

These produce names like `skew_z`, `skew_rank`, ..., `unusual_volume_signature_z`, `unusual_volume_signature_rank`.

## Strategy-model integration

Two dedicated paradigms consume these features:

- `options_directional`
  - `skew_z`, `put_call_flow_imbalance_z`, `dealer_positioning_proxy_z`,
    `gamma_exposure_proxy_rank`, `unusual_volume_signature_z`
- `options_volatility`
  - `convexity_z`, `term_structure_curvature_z`, `local_surface_distortion_z`,
    `oi_changes_z`, `unusual_volume_signature_rank`

## Sanity checks against known behavior

The pipeline emits boolean diagnostics:

- `skew_mostly_put_rich`: at least half observations have positive skew (`IV_25P > IV_25C`).
- `flow_bounded`: put-call flow imbalance remains in `[-1, 1]`.
- `unusual_volume_positive`: unusual volume signature is non-negative.
- `dealer_proxy_inverse_to_gex`: dealer proxy is non-positively correlated with gamma exposure proxy.

These checks are lightweight guardrails and should be monitored alongside data-quality checks.

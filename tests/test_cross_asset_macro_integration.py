from __future__ import annotations

import numpy as np

from src.analysis.cross_asset import (
    build_release_aware_macro_features,
    compare_conditioned_models,
    compute_cross_market_transmission,
    join_cross_asset_indicators,
)


def test_release_aware_macro_join_prevents_lookahead() -> None:
    target_ts = np.array([1.0, 2.0, 3.0, 4.0, 5.0])
    events = [
        {"event_ts": 1.0, "release_ts": 1.5, "value": 10.0},
        {"event_ts": 3.0, "release_ts": 3.5, "value": 20.0},
    ]

    out = build_release_aware_macro_features(target_timestamps=target_ts, events=events, max_lag_events=1)

    assert np.isnan(out["macro_latest"][0])
    assert out["macro_latest"][1] == 10.0
    assert out["macro_latest"][2] == 10.0
    assert out["macro_latest"][3] == 20.0
    assert out["macro_lag_1"][3] == 10.0


def test_cross_asset_conditioned_model_outperforms_isolated() -> None:
    rng = np.random.default_rng(42)
    n = 400
    equity = rng.normal(size=(n, 2))
    options = rng.normal(size=(n, 2))
    cross = rng.normal(size=(n, 3))

    # target depends materially on cross-asset block
    noise = rng.normal(scale=0.05, size=n)
    target = 0.1 * equity[:, 0] - 0.1 * options[:, 1] + 0.8 * cross[:, 0] - 0.5 * cross[:, 1] + noise

    comparison = compare_conditioned_models(
        equity_features=equity,
        options_features=options,
        cross_asset_features=cross,
        target=target,
    )

    assert comparison.outperformance_vs_best_isolated_mse > 0
    assert comparison.outperformance_vs_best_isolated_r2 > 0


def test_lead_lag_and_transmission_diagnostics_identify_driver() -> None:
    rng = np.random.default_rng(13)
    n = 200
    rates = rng.normal(size=n)
    credit = rng.normal(size=n)
    target = np.zeros(n)
    target[1:] = 0.7 * rates[:-1] + 0.1 * credit[:-1] + rng.normal(scale=0.05, size=n - 1)

    diag = compute_cross_market_transmission(
        target_series=target,
        driver_series={"rates": rates, "credit": credit},
        max_lag=3,
    )

    assert diag.strongest_driver == "rates"
    assert diag.strongest_lag == 1
    assert abs(diag.transmission_betas["rates"]) > abs(diag.transmission_betas["credit"])


def test_join_cross_asset_indicators_asof_alignment() -> None:
    target_ts = np.array([1.0, 2.0, 3.0, 4.0])
    indicators = {
        "rates": {"timestamps": [0.5, 2.2, 3.7], "values": [1.0, 2.0, 3.0]},
        "fx": {"timestamps": [1.2, 2.8], "values": [10.0, 20.0]},
    }
    out = join_cross_asset_indicators(target_timestamps=target_ts, indicators=indicators)

    assert out.shape == (4, 2)
    assert out[0, 0] == 1.0
    assert np.isnan(out[0, 1])
    assert out[2, 1] == 20.0

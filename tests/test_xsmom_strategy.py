from __future__ import annotations

import numpy as np

from src.backtesting.strategies.xsmom import (
    CrossSectionalMomentumConfig,
    assign_rank_buckets,
    build_cross_sectional_momentum_targets,
    compute_risk_normalized_weights,
)


def test_compute_risk_normalized_weights_sums_to_unit_gross() -> None:
    raw = np.array([2.0, -1.0, 1.0], dtype=float)
    weights = compute_risk_normalized_weights(raw)

    assert np.isclose(np.sum(np.abs(weights)), 1.0)
    assert np.isclose(weights[0], 0.5)
    assert np.isclose(weights[1], -0.25)


def test_assign_rank_buckets_respects_quantiles() -> None:
    scores = np.array([0.3, -0.2, 0.1, -0.4, 0.8], dtype=float)
    top, bottom = assign_rank_buckets(scores, top_quantile=0.4, bottom_quantile=0.4, long_only=False)

    assert set(np.flatnonzero(top)) == {0, 4}
    assert set(np.flatnonzero(bottom)) == {1, 3}


def test_xsmom_handles_small_universe_and_missing_data() -> None:
    prices = np.array(
        [
            [100.0, 50.0],
            [101.0, 50.5],
            [102.0, np.nan],
            [103.0, 49.0],
            [104.0, 48.0],
            [105.0, 47.0],
        ],
        dtype=float,
    )
    missing = ~np.isfinite(prices)
    prices = np.where(np.isfinite(prices), prices, np.nan)

    config = CrossSectionalMomentumConfig(
        lookback_days=2,
        skip_days=0,
        top_quantile=0.5,
        bottom_quantile=0.5,
        long_only=False,
        vol_lookback_days=2,
        rebalance_interval=1,
    )
    targets = build_cross_sectional_momentum_targets(
        close_prices=prices,
        missing_mask=missing,
        config=config,
    )

    assert targets.shape == prices.shape
    # no NaNs should leak into portfolio targets
    assert np.isfinite(targets).all()
    # at each bar, gross exposure is either flat or normalized
    gross = np.sum(np.abs(targets), axis=1)
    assert np.all((np.isclose(gross, 0.0)) | (np.isclose(gross, 1.0)))

from __future__ import annotations

import numpy as np

from src.analysis.factor_exposure import (
    FACTOR_NAMES,
    build_factor_exposure_model,
    decompose_alpha,
    residualize_alpha_signals,
)


def test_build_factor_exposure_model_shapes() -> None:
    prices = np.array(
        [
            [100.0, 80.0, 120.0],
            [101.0, 79.0, 121.0],
            [102.0, 78.0, 119.0],
            [103.0, 77.5, 118.0],
        ]
    )
    result = build_factor_exposure_model(prices=prices, lookback=3)
    assert result.factor_names == FACTOR_NAMES
    assert result.factor_returns.shape == (4, len(FACTOR_NAMES))
    assert result.exposures_by_asset.shape == (4, 3, len(FACTOR_NAMES))


def test_residualize_and_decompose_consistency() -> None:
    raw = np.array(
        [
            [1.0, 0.5, -0.4],
            [0.9, 0.2, -0.3],
        ]
    )
    exposures = np.array(
        [
            [[1.0, 0.0], [0.5, 0.2], [-0.7, 0.1]],
            [[1.0, 0.0], [0.5, 0.2], [-0.7, 0.1]],
        ]
    )
    resid = residualize_alpha_signals(raw_signals=raw, factor_exposures=exposures)
    assert resid.shape == raw.shape
    assert np.max(np.abs(np.einsum("tak,ta->tk", exposures, resid))) < 1e-4

    returns = np.array(
        [
            [0.0, 0.0, 0.0],
            [0.01, -0.005, 0.002],
        ]
    )
    weights = np.array(
        [
            [0.1, -0.1, 0.0],
            [0.2, -0.2, 0.1],
        ]
    )
    factor_returns = np.array(
        [
            [0.0, 0.0],
            [0.003, -0.002],
        ]
    )
    dec = decompose_alpha(
        weights=weights,
        asset_returns=returns,
        factor_exposures=exposures,
        factor_returns=factor_returns,
    )
    np.testing.assert_allclose(dec["gross_alpha"], dec["factor_beta_contribution"] + dec["residual_alpha"])

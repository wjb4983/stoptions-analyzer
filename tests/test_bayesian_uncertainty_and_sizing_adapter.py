from __future__ import annotations

import numpy as np

from src.modeling_nextgen.adapters.backtesting_adapter import BacktestingBridge
from src.modeling_nextgen.calibration.bayesian_uncertainty import BayesianUncertaintyCalibrator


def test_bayesian_uncertainty_outputs_intervals_and_decomposition() -> None:
    samples = np.array(
        [
            [0.10, 0.20, 0.40],
            [0.20, 0.25, 0.30],
            [0.15, 0.30, 0.35],
        ],
        dtype=float,
    )
    aleatoric = np.array([0.05, 0.02, 0.01], dtype=float)
    realized = np.array([0.12, 0.28, 0.32], dtype=float)

    calibrator = BayesianUncertaintyCalibrator(mode="bma", confidence_level=0.9)
    calibrator.fit(samples, realized=realized, aleatoric_std=aleatoric)
    posterior = calibrator.transform(samples, aleatoric_std=aleatoric)

    assert posterior.posterior_mean.shape == (3,)
    assert posterior.confidence_lower.shape == (3,)
    assert posterior.confidence_upper.shape == (3,)
    assert np.all(posterior.confidence_upper >= posterior.confidence_lower)
    assert np.all(posterior.total_std >= posterior.epistemic_std)
    assert np.all(posterior.aleatoric_std > 0.0)


def test_interval_aware_position_size_throttles_wide_intervals() -> None:
    bridge = BacktestingBridge()
    probabilities = np.array([0.9, 0.9], dtype=float)

    tight = bridge.interval_aware_position_size(
        probabilities,
        confidence_lower=np.array([0.85, 0.55], dtype=float),
        confidence_upper=np.array([0.95, 1.0], dtype=float),
        max_leverage=1.0,
        interval_aversion=1.0,
    )

    assert tight[0] > tight[1]
    assert np.all(np.abs(tight) <= 1.0)

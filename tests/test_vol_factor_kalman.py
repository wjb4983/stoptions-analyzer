from __future__ import annotations

import numpy as np

from src.modeling_nextgen.models.state_space.vol_factor_kalman import (
    infer_observation_matrix,
    run_vol_factor_kalman,
)


def test_infer_observation_matrix_respects_surface_feature_semantics() -> None:
    names = ["atm_level", "skew_slope", "term_curvature", "unseen_feature"]
    h = infer_observation_matrix(names)

    assert h.shape == (4, 2)
    assert h[0, 0] > h[0, 1]
    assert h[2, 1] > h[2, 0]
    assert np.isfinite(h).all()


def test_run_vol_factor_kalman_returns_filter_and_smoother_posteriors_as_features() -> None:
    rng = np.random.default_rng(42)
    n_steps = 30

    true_states = np.zeros((n_steps, 2), dtype=np.float64)
    for t in range(1, n_steps):
        true_states[t, 0] = 0.95 * true_states[t - 1, 0] + 0.05 * true_states[t - 1, 1] + rng.normal(0.0, 0.02)
        true_states[t, 1] = 0.90 * true_states[t - 1, 1] + rng.normal(0.0, 0.03)

    h = np.array([[1.0, 0.2], [0.6, 0.4], [0.2, 1.1]], dtype=np.float64)
    observations = true_states @ h.T + rng.normal(0.0, 0.03, size=(n_steps, 3))
    observations[5, 1] = np.nan

    output = run_vol_factor_kalman(observations, observation_matrix=h)
    features = output.as_feature_dict(prefix="state_space")

    assert output.filtered_mean.shape == (n_steps, 2)
    assert output.smoothed_mean.shape == (n_steps, 2)
    assert output.smoothed_covariance.shape == (n_steps, 2, 2)
    assert np.isfinite(output.filtered_mean).all()

    expected_keys = {
        "state_space_vol_level_filtered",
        "state_space_vol_of_vol_filtered",
        "state_space_vol_level_smoothed",
        "state_space_vol_of_vol_smoothed",
        "state_space_vol_level_var",
        "state_space_vol_of_vol_var",
        "state_space_cross_cov",
        "state_space_innovation_norm",
    }
    assert expected_keys == set(features)
    assert all(v.shape == (n_steps,) for v in features.values())
    assert np.all(features["state_space_vol_level_var"] >= -1e-10)
    assert np.all(features["state_space_vol_of_vol_var"] >= -1e-10)

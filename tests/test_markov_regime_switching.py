import numpy as np

from modeling_nextgen.models.markov.regime_switching import (
    RegimeSwitchingConfig,
    estimate_transition_matrix,
    fit_regime_switching_model,
)
from regime.classifier import (
    RegimeMarkovAdapterConfig,
    classify_regimes_with_markov_adapter,
)


def test_transition_matrix_estimation_normalizes_rows() -> None:
    posterior = np.array(
        [
            [0.8, 0.1, 0.1],
            [0.7, 0.2, 0.1],
            [0.2, 0.6, 0.2],
            [0.1, 0.3, 0.6],
        ],
        dtype=float,
    )
    transition = estimate_transition_matrix(posterior)
    assert transition.shape == (3, 3)
    assert np.allclose(np.sum(transition, axis=1), 1.0)


def test_markov_regime_switching_returns_posterior_per_date() -> None:
    observations = np.array(
        [
            [-1.0, -0.8],
            [-0.9, -1.1],
            [0.1, 0.2],
            [0.3, 0.0],
            [1.2, 0.9],
            [1.4, 1.1],
        ],
        dtype=float,
    )
    dates = np.array([f"2024-01-0{i}" for i in range(1, 7)], dtype=object)
    fit = fit_regime_switching_model(
        observations,
        dates=dates,
        config=RegimeSwitchingConfig(em_iterations=10),
    )

    assert fit.posterior_probabilities.shape == (observations.shape[0], 3)
    assert np.allclose(np.sum(fit.posterior_probabilities, axis=1), 1.0)
    assert np.array_equal(fit.dates, dates)
    assert set(fit.regime_names.tolist()) == {"calm", "stressed", "dislocated"}


def test_classifier_markov_adapter_emits_ensemble_signals() -> None:
    rng = np.random.default_rng(7)
    features = rng.normal(size=(12, 4))
    raw_feature_map = {
        "vix_term_structure": rng.normal(size=12),
        "realized_vol": np.abs(rng.normal(size=12)),
        "yield_curve_slope": rng.normal(size=12),
        "breadth": rng.normal(size=12),
        "liquidity_proxy": rng.normal(size=12),
        "credit_spread": np.abs(rng.normal(size=12)),
    }

    out = classify_regimes_with_markov_adapter(
        features=features,
        feature_names=("f1", "f2", "f3", "f4"),
        raw_feature_map=raw_feature_map,
        dates=np.arange(features.shape[0]),
        adapter_config=RegimeMarkovAdapterConfig(ensemble_weight_markov=0.6, markov_em_iterations=8),
    )

    assert out["markov_regime_probabilities"].shape == (features.shape[0], 3)
    assert out["markov_transition_matrix"].shape == (3, 3)
    assert out["ensemble_regime_labels"].shape[0] == features.shape[0]
    assert np.all((out["regime_signal_agreement"] >= 0.0) & (out["regime_signal_agreement"] <= 1.0))

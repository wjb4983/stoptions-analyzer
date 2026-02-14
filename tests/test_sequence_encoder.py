import numpy as np

from src.modeling_nextgen.models.deep.sequence_encoder import SequenceEncoder


def _features(n: int = 16, t: int = 12) -> dict[str, np.ndarray]:
    rng = np.random.default_rng(42)
    return {
        "order_flow": rng.normal(size=(n, t, 2)),
        "imbalance": rng.normal(size=(n, t)),
        "gamma_proxy": rng.normal(size=(n, t)),
        "realized_vol": np.abs(rng.normal(size=n)),
    }


def test_sequence_encoder_predict_proba_shape_by_architecture() -> None:
    features = _features()
    labels = np.random.default_rng(7).integers(0, 3, size=(16, 3))

    for architecture in ("tcn", "lstm", "transformer"):
        model = SequenceEncoder(architecture=architecture, epochs=20, seed=5)
        model.fit(features, labels)
        probs = model.predict_proba(features)
        assert probs.shape == (16, 3, 3)
        assert np.allclose(np.sum(probs, axis=-1), 1.0, atol=1e-5)


def test_sequence_encoder_horizon_slice() -> None:
    features = _features()
    labels = np.random.default_rng(9).integers(0, 3, size=16)

    model = SequenceEncoder(architecture="tcn", horizons=("5m", "30m"), epochs=15)
    model.fit(features, labels)

    horizon_probs = model.predict_proba(features, horizon="30m")
    assert horizon_probs.shape == (16, 3)
    assert np.allclose(np.sum(horizon_probs, axis=-1), 1.0, atol=1e-5)

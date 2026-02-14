from __future__ import annotations

import numpy as np
import pytest

from src.modeling_nextgen.models.ml.multitask import MultiTaskRiskModel


def _dataset(n_samples: int = 32, input_dim: int = 5) -> tuple[np.ndarray, dict[str, np.ndarray]]:
    rng = np.random.default_rng(123)
    x = rng.normal(size=(n_samples, input_dim))
    labels = {
        "directional_return": rng.integers(0, 3, size=n_samples),
        "realized_vol_bucket": rng.integers(0, 3, size=n_samples),
        "drawdown_risk": rng.integers(0, 3, size=n_samples),
    }
    return x, labels


def test_multitask_validates_input_shape() -> None:
    x, y_by_task = _dataset()
    model = MultiTaskRiskModel(input_dim=5, epochs=5)

    with pytest.raises(ValueError, match="X must have shape"):
        model.fit(x[:, :4], y_by_task)

    with pytest.raises(ValueError, match="X must have shape"):
        model.predict_proba(np.ones((10, 4)))


def test_multitask_task_head_output_consistency() -> None:
    x, y_by_task = _dataset()
    model = MultiTaskRiskModel(input_dim=5, epochs=15, seed=77)
    model.fit(x, y_by_task)

    probs = model.predict_proba(x)
    preds = model.predict(x)

    assert set(probs) == {"directional_return", "realized_vol_bucket", "drawdown_risk"}
    assert set(preds) == set(probs)

    for task_name, task_probs in probs.items():
        assert task_probs.shape == (x.shape[0], 3)
        np.testing.assert_allclose(task_probs.sum(axis=1), np.ones(x.shape[0]), atol=1e-6)
        np.testing.assert_array_equal(np.argmax(task_probs, axis=1), preds[task_name])


def test_multitask_seeded_runs_are_deterministic() -> None:
    x, y_by_task = _dataset(n_samples=24)

    model_a = MultiTaskRiskModel(input_dim=5, epochs=20, seed=999)
    model_b = MultiTaskRiskModel(input_dim=5, epochs=20, seed=999)

    model_a.fit(x, y_by_task)
    model_b.fit(x, y_by_task)

    probs_a = model_a.predict_proba(x)
    probs_b = model_b.predict_proba(x)

    for task in probs_a:
        np.testing.assert_allclose(probs_a[task], probs_b[task], atol=1e-10)

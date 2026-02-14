from __future__ import annotations

import numpy as np
import pytest

from src.modeling_nextgen.core.contracts import PredictionResult
from src.modeling_nextgen.serving.inference import InferenceService


class _SchemaCheckedModel:
    name = "schema_checked"

    def fit(self, features: dict[str, np.ndarray], labels: np.ndarray) -> None:
        _ = (features, labels)

    def predict(self, features: dict[str, np.ndarray]) -> PredictionResult:
        if "x" not in features:
            raise ValueError("missing 'x' feature")
        x = np.asarray(features["x"], dtype=float)
        if x.ndim != 2:
            raise ValueError("x must be a 2D array")
        preds = x.sum(axis=1)
        return PredictionResult(
            predictions=preds,
            uncertainty=np.full_like(preds, 0.1, dtype=float),
            metadata={"schema_version": "1.0", "n_samples": int(x.shape[0])},
        )


def test_inference_request_response_schema_contract() -> None:
    service = InferenceService(_SchemaCheckedModel())

    result = service.predict({"x": np.array([[1.0, 2.0], [3.0, 4.0]])})

    assert isinstance(result, PredictionResult)
    assert result.predictions.shape == (2,)
    assert result.uncertainty is not None and result.uncertainty.shape == (2,)
    assert result.metadata is not None
    assert result.metadata["schema_version"] == "1.0"
    assert result.metadata["n_samples"] == 2


def test_inference_invalid_input_is_propagated() -> None:
    service = InferenceService(_SchemaCheckedModel())

    with pytest.raises(ValueError, match="missing 'x' feature"):
        service.predict({})

    with pytest.raises(ValueError, match="x must be a 2D array"):
        service.predict({"x": np.array([1.0, 2.0, 3.0])})


def test_inference_batch_vs_single_parity() -> None:
    service = InferenceService(_SchemaCheckedModel())
    batch = np.array([[1.0, 2.0], [5.0, -1.0], [3.5, 0.5]])

    batch_result = service.predict({"x": batch}).predictions
    single_stitched = np.array([
        service.predict({"x": row[None, :]}).predictions[0] for row in batch
    ])

    np.testing.assert_allclose(batch_result, single_stitched)

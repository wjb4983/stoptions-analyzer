from __future__ import annotations

import numpy as np

from ..core.contracts import Model, PredictionResult


class InferenceService:
    """Lightweight wrapper for serving next-gen models in-process."""

    def __init__(self, model: Model) -> None:
        self._model = model

    def predict(self, features: dict[str, np.ndarray]) -> PredictionResult:
        return self._model.predict(features)

from __future__ import annotations

from dataclasses import dataclass

import numpy as np

from ..core.contracts import PredictionResult


@dataclass
class NextGenModelBase:
    name: str

    def fit(self, features: dict[str, np.ndarray], labels: np.ndarray) -> None:
        _ = (features, labels)

    def predict(self, features: dict[str, np.ndarray]) -> PredictionResult:
        sample_count = 0
        if features:
            first = next(iter(features.values()))
            sample_count = int(np.asarray(first).shape[0])
        predictions = np.zeros(sample_count, dtype=float)
        return PredictionResult(predictions=predictions)

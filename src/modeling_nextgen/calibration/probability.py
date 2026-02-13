from __future__ import annotations

import numpy as np


class IdentityProbabilityCalibrator:
    name = "identity_probability"

    def fit(self, probabilities: np.ndarray, labels: np.ndarray) -> None:
        _ = (probabilities, labels)

    def transform(self, probabilities: np.ndarray) -> np.ndarray:
        return np.asarray(probabilities, dtype=float)

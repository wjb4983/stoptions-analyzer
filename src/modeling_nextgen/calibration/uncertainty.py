from __future__ import annotations

import numpy as np


class IdentityUncertaintyCalibrator:
    name = "identity_uncertainty"

    def fit(self, uncertainty: np.ndarray) -> None:
        _ = uncertainty

    def transform(self, uncertainty: np.ndarray) -> np.ndarray:
        return np.asarray(uncertainty, dtype=float)

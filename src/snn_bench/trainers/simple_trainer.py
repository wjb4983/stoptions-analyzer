from __future__ import annotations

import numpy as np


class SimpleTrainer:
    """Thin trainer wrapper to standardize fit/predict flow."""

    def __init__(self, model) -> None:
        self.model = model

    def run(self, x_train: np.ndarray, y_train: np.ndarray, x_eval: np.ndarray) -> np.ndarray:
        self.model.fit(x_train, y_train)
        return self.model.predict(x_eval)

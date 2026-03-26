from __future__ import annotations

import numpy as np


class DummySpikingModel:
    """Minimal linear classifier placeholder for SNN experimentation."""

    def __init__(self, seed: int = 42) -> None:
        self.seed = seed
        self.weights: np.ndarray | None = None

    def fit(self, x: np.ndarray, y: np.ndarray) -> None:
        if x.size == 0:
            self.weights = np.array([])
            return
        rng = np.random.default_rng(self.seed)
        self.weights = rng.normal(0, 0.1, size=x.shape[1])

    def predict(self, x: np.ndarray) -> np.ndarray:
        if self.weights is None or x.size == 0:
            return np.array([], dtype=int)
        logits = x @ self.weights
        return (logits > 0).astype(int)

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Protocol

import numpy as np


@dataclass(frozen=True)
class PredictionResult:
    predictions: np.ndarray
    probabilities: np.ndarray | None = None
    uncertainty: np.ndarray | None = None
    metadata: dict[str, Any] | None = None


class FeatureBuilder(Protocol):
    name: str

    def build(self, frame: Any) -> dict[str, np.ndarray]:
        ...


class Model(Protocol):
    name: str

    def fit(self, features: dict[str, np.ndarray], labels: np.ndarray) -> None:
        ...

    def predict(self, features: dict[str, np.ndarray]) -> PredictionResult:
        ...


class ProbabilisticModel(Model, Protocol):
    def predict_proba(self, features: dict[str, np.ndarray]) -> np.ndarray:
        ...


class Validator(Protocol):
    name: str

    def evaluate(self, model: Model, features: dict[str, np.ndarray], labels: np.ndarray) -> dict[str, float]:
        ...

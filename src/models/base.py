from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import Any, Protocol

import numpy as np


@dataclass(frozen=True)
class ModelExplanation:
    model_name: str
    feature_importances: dict[str, float]
    confidence_scores: np.ndarray
    metadata: dict[str, Any] | None = None


class ModelInterface(Protocol):
    name: str
    feature_importances_: dict[str, float]
    confidence_scores_: np.ndarray

    def fit(self, features: dict[str, np.ndarray], labels: np.ndarray) -> None:
        ...

    def predict(self, features: dict[str, np.ndarray]) -> np.ndarray:
        ...

    def predict_proba(self, features: dict[str, np.ndarray]) -> np.ndarray:
        ...

    def explain(self, features: dict[str, np.ndarray]) -> ModelExplanation:
        ...


class BaseParadigmModel(ABC):
    name: str = "base"
    required_features: tuple[str, ...] = ()

    def __init__(self) -> None:
        self.feature_importances_: dict[str, float] = {}
        self.confidence_scores_: np.ndarray = np.array([], dtype=float)
        self._weights: np.ndarray | None = None

    @abstractmethod
    def required_feature_names(self) -> tuple[str, ...]:
        ...

    def _build_feature_matrix(self, features: dict[str, np.ndarray]) -> np.ndarray:
        names = self.required_feature_names()
        if not names:
            raise ValueError(f"{self.name} must define at least one feature")
        columns: list[np.ndarray] = []
        expected_len: int | None = None
        for feature_name in names:
            if feature_name not in features:
                raise KeyError(f"Feature '{feature_name}' required by '{self.name}' is missing")
            values = np.asarray(features[feature_name], dtype=float)
            if values.ndim != 1:
                raise ValueError(f"Feature '{feature_name}' must be a 1D array")
            if expected_len is None:
                expected_len = values.shape[0]
            elif values.shape[0] != expected_len:
                raise ValueError("All features must have the same number of samples")
            columns.append(values)
        return np.column_stack(columns)

    def fit(self, features: dict[str, np.ndarray], labels: np.ndarray) -> None:
        x = self._build_feature_matrix(features)
        y = np.asarray(labels, dtype=float)
        if y.ndim != 1 or y.shape[0] != x.shape[0]:
            raise ValueError("labels must be 1D and aligned to feature rows")

        centered_x = x - np.mean(x, axis=0)
        centered_y = y - np.mean(y)
        raw_importances = np.abs(np.dot(centered_x.T, centered_y))
        total = float(np.sum(raw_importances))
        if total <= 0.0:
            weights = np.full(raw_importances.shape, 1.0 / max(1, raw_importances.size), dtype=float)
        else:
            weights = raw_importances / total
        self._weights = weights
        self.feature_importances_ = dict(zip(self.required_feature_names(), weights, strict=True))

        probs = self.predict_proba(features)
        self.confidence_scores_ = np.abs(probs - 0.5) * 2.0

    def predict_proba(self, features: dict[str, np.ndarray]) -> np.ndarray:
        if self._weights is None:
            raise RuntimeError(f"Model '{self.name}' must be fit before predict_proba")
        x = self._build_feature_matrix(features)
        logits = np.dot(x, self._weights)
        logits = logits - np.mean(logits)
        return 1.0 / (1.0 + np.exp(-logits))

    def predict(self, features: dict[str, np.ndarray]) -> np.ndarray:
        probs = self.predict_proba(features)
        self.confidence_scores_ = np.abs(probs - 0.5) * 2.0
        return np.where(probs >= 0.5, 1.0, -1.0)

    def explain(self, features: dict[str, np.ndarray]) -> ModelExplanation:
        probs = self.predict_proba(features)
        confidence = np.abs(probs - 0.5) * 2.0
        self.confidence_scores_ = confidence
        return ModelExplanation(
            model_name=self.name,
            feature_importances=dict(self.feature_importances_),
            confidence_scores=confidence,
            metadata={"required_features": self.required_feature_names()},
        )

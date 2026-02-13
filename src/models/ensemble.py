from __future__ import annotations

import logging
from dataclasses import dataclass

import numpy as np

from .base import ModelInterface

LOGGER = logging.getLogger(__name__)


@dataclass(frozen=True)
class EnsembleOutput:
    signal: np.ndarray
    probability: np.ndarray
    confidence_scores: np.ndarray
    feature_importances: dict[str, dict[str, float]]


class ModelEnsembler:
    def __init__(self, models: list[tuple[ModelInterface, float]]) -> None:
        if not models:
            raise ValueError("models must be non-empty")
        self.models = models
        self.stacking_weights_: np.ndarray | None = None

    def fit(self, features: dict[str, np.ndarray], labels: np.ndarray) -> None:
        for model, _ in self.models:
            model.fit(features, labels)
            LOGGER.info("fit model=%s importances=%s", model.name, model.feature_importances_)

    def weighted_vote(self, features: dict[str, np.ndarray]) -> EnsembleOutput:
        probs = []
        weights = np.asarray([weight for _, weight in self.models], dtype=float)
        gross = float(np.sum(np.abs(weights)))
        if gross <= 0.0:
            raise ValueError("ensemble weights must have non-zero gross")
        weights = weights / gross

        feature_importances: dict[str, dict[str, float]] = {}
        confidences = []
        for model, _ in self.models:
            model_probs = model.predict_proba(features)
            probs.append(model_probs)
            feature_importances[model.name] = dict(model.feature_importances_)
            confidences.append(np.abs(model_probs - 0.5) * 2.0)
            LOGGER.info("predict model=%s mean_confidence=%.4f", model.name, float(np.mean(confidences[-1])))

        prob_matrix = np.column_stack(probs)
        blended_prob = np.dot(prob_matrix, weights)
        confidence = np.dot(np.column_stack(confidences), weights)
        signal = np.where(blended_prob >= 0.5, 1.0, -1.0)
        return EnsembleOutput(
            signal=signal,
            probability=blended_prob,
            confidence_scores=confidence,
            feature_importances=feature_importances,
        )

    def fit_stacking(self, features: dict[str, np.ndarray], labels: np.ndarray) -> None:
        self.fit(features, labels)
        y = np.asarray(labels, dtype=float)
        base_prob_matrix = np.column_stack([model.predict_proba(features) for model, _ in self.models])
        centered_y = y - np.mean(y)
        centered_x = base_prob_matrix - np.mean(base_prob_matrix, axis=0)
        coeffs, *_ = np.linalg.lstsq(centered_x, centered_y, rcond=None)
        if np.allclose(coeffs, 0.0):
            coeffs = np.full(coeffs.shape, 1.0 / max(1, coeffs.size), dtype=float)
        self.stacking_weights_ = coeffs
        LOGGER.info("stacking weights=%s", coeffs)

    def stacking_predict(self, features: dict[str, np.ndarray]) -> EnsembleOutput:
        if self.stacking_weights_ is None:
            raise RuntimeError("fit_stacking must be called before stacking_predict")

        probs = np.column_stack([model.predict_proba(features) for model, _ in self.models])
        logits = np.dot(probs, self.stacking_weights_)
        centered = logits - np.mean(logits)
        ensemble_prob = 1.0 / (1.0 + np.exp(-centered))
        signal = np.where(ensemble_prob >= 0.5, 1.0, -1.0)

        feature_importances = {model.name: dict(model.feature_importances_) for model, _ in self.models}
        confidence = np.abs(ensemble_prob - 0.5) * 2.0
        return EnsembleOutput(
            signal=signal,
            probability=ensemble_prob,
            confidence_scores=confidence,
            feature_importances=feature_importances,
        )

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Protocol

import numpy as np


@dataclass(frozen=True)
class FeatureBatch:
    """Feature matrix and optional metadata used by alpha strategy plugins."""

    values: np.ndarray
    feature_names: tuple[str, ...]
    timestamps: np.ndarray | None = None
    metadata: dict[str, Any] | None = None


@dataclass(frozen=True)
class LabelSpec:
    """Label horizon and thresholding for supervised alpha targets."""

    horizon: int
    return_threshold: float = 0.0
    label_positive: int = 1
    label_negative: int = 0


@dataclass(frozen=True)
class ExplainabilityPayload:
    """Per-observation explainability payload for model diagnostics."""

    importances: np.ndarray
    contributions: np.ndarray | None = None
    metadata: dict[str, Any] | None = None


class AlphaStrategyPlugin(Protocol):
    """Contract for alpha model strategy plugins in backtesting."""

    def generate_features(self, raw_inputs: dict[str, np.ndarray]) -> FeatureBatch:
        """Build model-ready features from raw market inputs."""

    def define_label_horizon(self) -> LabelSpec:
        """Declare supervised label horizon and transformation defaults."""

    def fit(self, features: FeatureBatch, labels: np.ndarray, *, sample_weight: np.ndarray | None = None) -> None:
        """Fit the plugin model."""

    def predict(self, features: FeatureBatch) -> np.ndarray:
        """Predict signed alpha signal."""

    def predict_proba(self, features: FeatureBatch) -> np.ndarray:
        """Predict probability of positive label; shape = (n_samples,)."""

    def explain(self, features: FeatureBatch) -> ExplainabilityPayload:
        """Return explainability payload aligned to feature rows."""


@dataclass(frozen=True)
class MetaLabelingResult:
    base_signal: np.ndarray
    confidence: np.ndarray
    gated_signal: np.ndarray
    gate_mask: np.ndarray


def apply_meta_labeling(
    base_signal: np.ndarray,
    classifier_confidence: np.ndarray,
    *,
    confidence_threshold: float = 0.55,
) -> MetaLabelingResult:
    """Gate base signals with classifier confidence for meta-labeling workflows."""

    signal = np.asarray(base_signal, dtype=float)
    confidence = np.asarray(classifier_confidence, dtype=float)
    if signal.shape != confidence.shape:
        raise ValueError("base_signal and classifier_confidence must have matching shapes")

    threshold = float(confidence_threshold)
    gate = confidence >= threshold
    gated = np.where(gate, signal, 0.0)
    return MetaLabelingResult(
        base_signal=signal,
        confidence=confidence,
        gated_signal=gated,
        gate_mask=gate,
    )


def probability_calibrated_position_size(
    probabilities: np.ndarray,
    *,
    neutral_probability: float = 0.5,
    max_leverage: float = 1.0,
    gamma: float = 1.0,
) -> np.ndarray:
    """Map calibrated probabilities to signed position sizes in [-max_leverage, max_leverage]."""

    probs = np.asarray(probabilities, dtype=float)
    neutral = float(neutral_probability)
    if not 0.0 < neutral < 1.0:
        raise ValueError("neutral_probability must be in (0, 1)")

    clipped = np.clip(probs, 0.0, 1.0)
    centered = clipped - neutral
    scale = max(neutral, 1.0 - neutral)
    signed = np.divide(centered, scale, out=np.zeros_like(centered), where=scale > 0.0)
    transformed = np.sign(signed) * (np.abs(signed) ** max(float(gamma), 1e-8))
    return np.clip(transformed * float(max_leverage), -float(max_leverage), float(max_leverage))

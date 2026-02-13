from __future__ import annotations

from dataclasses import dataclass

import numpy as np

from src.analysis.reporting import compute_model_drift_diagnostics
from src.backtesting.strategies.alpha_model import (
    ExplainabilityPayload,
    FeatureBatch,
    LabelSpec,
    apply_meta_labeling,
    probability_calibrated_position_size,
)
from src.backtesting.strategies.ensemble import dynamic_model_weights, rolling_dynamic_ensemble


@dataclass
class DeterministicLinearPlugin:
    seed: int = 123

    def __post_init__(self) -> None:
        self._coef: np.ndarray | None = None

    def generate_features(self, raw_inputs: dict[str, np.ndarray]) -> FeatureBatch:
        x = np.column_stack([raw_inputs["x1"], raw_inputs["x2"]]).astype(float)
        return FeatureBatch(values=x, feature_names=("x1", "x2"))

    def define_label_horizon(self) -> LabelSpec:
        return LabelSpec(horizon=5, return_threshold=0.001)

    def fit(self, features: FeatureBatch, labels: np.ndarray, *, sample_weight: np.ndarray | None = None) -> None:
        x = np.asarray(features.values, dtype=float)
        y = np.asarray(labels, dtype=float).reshape(-1)
        ridge = 1e-6 * np.eye(x.shape[1])
        self._coef = np.linalg.solve(x.T @ x + ridge, x.T @ y)

    def predict(self, features: FeatureBatch) -> np.ndarray:
        if self._coef is None:
            raise RuntimeError("model not fit")
        return np.asarray(features.values, dtype=float) @ self._coef

    def predict_proba(self, features: FeatureBatch) -> np.ndarray:
        raw = self.predict(features)
        return 1.0 / (1.0 + np.exp(-raw))

    def explain(self, features: FeatureBatch) -> ExplainabilityPayload:
        if self._coef is None:
            raise RuntimeError("model not fit")
        vals = np.asarray(features.values, dtype=float)
        contrib = vals * self._coef[None, :]
        return ExplainabilityPayload(importances=np.abs(self._coef), contributions=contrib)


def test_fit_predict_reproducibility() -> None:
    rng = np.random.default_rng(7)
    raw = {
        "x1": rng.normal(size=64),
        "x2": rng.normal(size=64),
    }
    labels = 0.6 * raw["x1"] - 0.2 * raw["x2"] + rng.normal(scale=0.01, size=64)

    plugin_a = DeterministicLinearPlugin()
    plugin_b = DeterministicLinearPlugin()

    feats_a = plugin_a.generate_features(raw)
    feats_b = plugin_b.generate_features(raw)
    plugin_a.fit(feats_a, labels)
    plugin_b.fit(feats_b, labels)

    pred_a = plugin_a.predict(feats_a)
    pred_b = plugin_b.predict(feats_b)

    assert np.allclose(pred_a, pred_b)
    assert np.allclose(plugin_a.predict_proba(feats_a), plugin_b.predict_proba(feats_b))


def test_meta_labeling_and_probability_sizing() -> None:
    base = np.array([1.0, -1.0, 0.5, -0.25], dtype=float)
    conf = np.array([0.8, 0.3, 0.7, 0.49], dtype=float)
    labeled = apply_meta_labeling(base, conf, confidence_threshold=0.5)

    assert np.array_equal(labeled.gate_mask, np.array([True, False, True, False]))
    assert np.allclose(labeled.gated_signal, np.array([1.0, 0.0, 0.5, 0.0]))

    size = probability_calibrated_position_size(np.array([0.1, 0.5, 0.9]), max_leverage=1.5, gamma=1.0)
    assert np.allclose(size, np.array([-1.2, 0.0, 1.2]))


def test_dynamic_ensemble_weights_follow_quality() -> None:
    quality = np.array(
        [
            [0.2, 0.1, 0.0],
            [0.3, 0.1, 0.0],
            [0.4, 0.05, 0.0],
            [0.5, 0.0, 0.0],
        ],
        dtype=float,
    )
    weights = dynamic_model_weights(quality, lookback=3)

    assert np.isclose(weights.sum(), 1.0)
    assert weights[0] > weights[1] >= weights[2]

    signals = np.array(
        [
            [1.0, 0.0, -1.0],
            [1.0, 0.2, -1.0],
            [1.0, 0.3, -1.0],
            [1.0, 0.5, -1.0],
        ],
        dtype=float,
    )
    blended = rolling_dynamic_ensemble(signals, quality, lookback=2)
    assert blended.shape == (signals.shape[0],)
    assert blended[-1] > blended[0]


def test_model_drift_diagnostics_trigger() -> None:
    returns = np.array([0.001] * 20 + [-0.01] * 20, dtype=float)
    drift = compute_model_drift_diagnostics(returns=returns, drift_z_threshold=1.0)

    assert "retraining_triggered" in drift
    assert bool(drift["retraining_triggered"]) is True

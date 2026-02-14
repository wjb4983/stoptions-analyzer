from __future__ import annotations

import numpy as np

from modeling_nextgen.models.ml.meta_label_conformal import AcceptancePolicy, MetaLabelConformalModel
from models.paradigms import MetaLabelClassifierModel


def _features(n: int = 64) -> dict[str, np.ndarray]:
    idx = np.linspace(-1.0, 1.0, n)
    return {
        "base_signal": np.where(idx >= 0.0, 1.0, -1.0),
        "base_confidence": np.abs(idx),
        "risk_filter_score": 0.5 + 0.5 * idx,
    }


def test_meta_label_conformal_generates_valid_p_values_and_policy_outputs() -> None:
    features = _features(128)
    labels = (features["base_confidence"] + 0.2 * features["risk_filter_score"] > 0.6).astype(float)

    model = MetaLabelConformalModel(
        acceptance_policy=AcceptancePolicy(min_base_confidence=0.35, min_risk_filter_score=0.25),
        target_risk=0.35,
        min_coverage=0.1,
    )
    model.fit(features, labels)

    probs = model.predict_proba(features)
    p_values = model.conformal_p_values(probs)
    decision = model.apply_policy(features)

    assert probs.shape == labels.shape
    assert p_values.shape == labels.shape
    assert np.all((0.0 <= p_values) & (p_values <= 1.0))
    assert decision.accepted_mask.shape == labels.shape
    assert decision.rejected_mask.shape == labels.shape
    assert decision.gated_signal.shape == labels.shape
    assert np.isclose(decision.empirical_coverage, np.mean(decision.accepted_mask))


def test_legacy_meta_label_classifier_uses_compatibility_adapter() -> None:
    features = _features(80)
    labels = (features["base_confidence"] > 0.45).astype(float)

    model = MetaLabelClassifierModel()
    model.fit(features, labels)
    probs = model.predict_proba(features)

    assert probs.shape == labels.shape
    assert np.all((0.0 <= probs) & (probs <= 1.0))
    assert "base_signal" in model.feature_importances_
    assert np.any(np.isclose(probs, 0.5))

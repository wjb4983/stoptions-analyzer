from __future__ import annotations

import numpy as np

from modeling_nextgen.validation.adversarial import AdversarialValidationConfig, build_adversarial_fragility_scorecards


class StableModel:
    def predict_proba(self, features: dict[str, np.ndarray]) -> np.ndarray:
        x = np.asarray(features["signal"], dtype=float)
        # Intentionally insensitive to sign flips.
        z = np.abs(x)
        return 1.0 / (1.0 + np.exp(-z))


class BrittleModel:
    def predict_proba(self, features: dict[str, np.ndarray]) -> np.ndarray:
        x = np.asarray(features["signal"], dtype=float)
        return 1.0 / (1.0 + np.exp(-6.0 * x))


def test_adversarial_fragility_scorecards_include_required_scenarios_and_gates() -> None:
    signal = np.linspace(-1.0, 1.0, 128)
    payloads = {
        "stable": {"signal": signal},
        "brittle": {"signal": signal},
    }
    models = {
        "stable": StableModel(),
        "brittle": BrittleModel(),
    }

    report = build_adversarial_fragility_scorecards(
        models=models,
        feature_payloads=payloads,
        config=AdversarialValidationConfig(max_fragility_score=0.18, max_worst_case_degradation=0.3),
        seed=7,
    )

    by_id = {row["model_id"]: row for row in report["models"]}
    assert set(by_id) == {"stable", "brittle"}
    assert report["pass_rate"] < 1.0

    for row in by_id.values():
        assert set(row["scenario_degradation"]) == {
            "feature_sign_flips",
            "temporal_jitter",
            "sparse_missing_block_corruption",
        }

    assert by_id["brittle"]["auto_fail_gate"] is True
    assert by_id["brittle"]["fragility_score"] > by_id["stable"]["fragility_score"]


def test_adversarial_fragility_all_models_pass_when_thresholds_relaxed() -> None:
    class ConstantModel:
        def predict_proba(self, features: dict[str, np.ndarray]) -> np.ndarray:
            n = len(next(iter(features.values())))
            return np.full(n, 0.5, dtype=float)

    report = build_adversarial_fragility_scorecards(
        models={"const": ConstantModel()},
        feature_payloads={"const": {"x": np.arange(32, dtype=float)}},
        seed=11,
    )

    assert report["all_models_pass"] is True
    assert report["models"][0]["auto_fail_gate"] is False

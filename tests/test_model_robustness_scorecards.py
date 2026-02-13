from __future__ import annotations

import numpy as np

from models.paradigms import MomentumModel
from models.robustness import RobustnessThresholds, build_robustness_scorecards


def _make_payload(n: int = 64) -> dict[str, np.ndarray]:
    x = np.linspace(-1.0, 1.0, n)
    return {
        "returns_1m": x,
        "returns_3m": x * 0.7 + 0.1,
        "returns_6m": x * -0.3 + 0.2,
    }


def test_build_robustness_scorecards_reports_scenarios_groups_and_flags() -> None:
    model = MomentumModel()
    features = _make_payload()
    labels = np.where(features["returns_1m"] > 0.0, 1.0, -1.0)
    model.fit(features, labels)

    report = build_robustness_scorecards(
        models={"momentum": model},
        feature_payloads={"momentum": features},
        feature_groups={"trend": ("returns_1m", "returns_3m"), "long_horizon": ("returns_6m",)},
        thresholds=RobustnessThresholds(min_model_score=0.1, min_feature_group_score=0.1, max_brittle_features=3),
        seed=7,
    )

    assert report["production_ready"] is True
    assert len(report["models"]) == 1
    row = report["models"][0]
    assert set(row["scenario_degradation"].keys()) == {"noise", "delayed_data", "missing_fields", "shifted_distribution"}
    assert "trend" in row["feature_groups"]
    assert "long_horizon" in row["feature_groups"]
    assert all("recommended_action" in f for f in row["feature_breakdown"])


def test_build_robustness_scorecards_enforces_thresholds_and_marks_brittle_features() -> None:
    model = MomentumModel()
    features = _make_payload()
    labels = np.where(features["returns_1m"] + features["returns_3m"] > 0.0, 1.0, -1.0)
    model.fit(features, labels)

    report = build_robustness_scorecards(
        models={"momentum": model},
        feature_payloads={"momentum": features},
        thresholds=RobustnessThresholds(min_model_score=0.99, min_feature_group_score=0.99, max_brittle_features=0),
        feature_brittleness_threshold=0.05,
        seed=11,
    )

    assert report["production_ready"] is False
    row = report["models"][0]
    assert row["meets_minimum_threshold"] is False
    assert isinstance(row["brittle_features"], list)
    assert len(report["recommended_feature_actions"]) == len(row["brittle_features"])

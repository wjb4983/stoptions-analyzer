from __future__ import annotations

import copy
import json

from reports import benchmark_bundle


DIMENSIONS = [
    "robust_oos_performance",
    "statistical_significance",
    "execution_realism",
    "stress_resilience",
    "reproducibility",
]


def test_benchmark_bundle_writes_scorecard_artifact(tmp_path) -> None:
    scorecard = benchmark_bundle.build_benchmark_scorecard()
    out_path = tmp_path / "benchmark_scorecard.json"

    written = benchmark_bundle.write_scorecard(scorecard, out_path)

    assert written == out_path
    payload = json.loads(out_path.read_text())
    assert set(payload["dimensions"]) == set(DIMENSIONS)
    assert payload["promotion_gate"]["pass"] is True
    assert payload["promotion_gate"]["failed_critical_dimensions"] == []



def test_benchmark_bundle_gates_promotion_on_critical_failure() -> None:
    dataset = copy.deepcopy(benchmark_bundle._load_json(benchmark_bundle.DATASET_PATH))
    dataset["oos_returns"] = [-0.04, -0.03, -0.02, -0.01, -0.02, -0.03]

    scorecard = benchmark_bundle.build_benchmark_scorecard(dataset=dataset)

    assert scorecard["dimensions"]["robust_oos_performance"]["pass"] is False
    assert scorecard["promotion_gate"]["pass"] is False
    assert "robust_oos_performance" in scorecard["promotion_gate"]["failed_critical_dimensions"]

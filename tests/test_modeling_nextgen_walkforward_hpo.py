from __future__ import annotations

import json

from src.modeling_nextgen.validation.walkforward_hpo import (
    build_walkforward_windows,
    export_walkforward_hpo_reports,
    run_walkforward_hpo,
)


def test_run_walkforward_hpo_nested_tuning_and_outer_eval() -> None:
    windows = build_walkforward_windows(total_samples=30, train_size=10, test_size=5, step_size=5)
    candidates = [
        {"alpha": 1, "beta": 0.1},
        {"alpha": 2, "beta": 0.3},
    ]

    def inner_objective(params: dict[str, float], train_start: int, train_end: int) -> float:
        train_len = train_end - train_start
        return float(params["alpha"] * 0.5 + train_len * params["beta"])

    def outer_evaluator(params: dict[str, float], test_start: int, test_end: int) -> dict[str, object]:
        test_len = test_end - test_start
        sharpe = params["alpha"] - 0.1 * test_len
        return {"metrics": {"sharpe": float(sharpe), "return": float(test_len * 0.01)}}

    summary = run_walkforward_hpo(
        windows=windows,
        param_candidates=candidates,
        inner_objective=inner_objective,
        outer_evaluator=outer_evaluator,
        primary_metric="sharpe",
    )

    assert len(summary.folds) == 4
    assert all(row["selected_params"]["alpha"] == 2 for row in summary.folds)
    assert "transition_rate" in summary.parameter_stability
    assert "sharpe_mean" in summary.aggregate_metrics
    assert summary.model_card["card_version"] == "1.0"


def test_export_walkforward_hpo_reports_model_card_schema(tmp_path) -> None:
    windows = build_walkforward_windows(total_samples=24, train_size=8, test_size=4, step_size=4)
    candidates = [{"lookback": 5}, {"lookback": 10}]

    def inner_objective(params: dict[str, int], _train_start: int, _train_end: int) -> float:
        return float(params["lookback"])

    def outer_evaluator(params: dict[str, int], _test_start: int, _test_end: int) -> dict[str, object]:
        return {"metrics": {"sharpe": float(params["lookback"]) / 10.0}}

    summary = run_walkforward_hpo(
        windows=windows,
        param_candidates=candidates,
        inner_objective=inner_objective,
        outer_evaluator=outer_evaluator,
        primary_metric="sharpe",
    )
    paths = export_walkforward_hpo_reports(summary=summary, reports_dir=tmp_path, artifact_prefix="wf_test")

    assert paths["summary"].exists()
    assert paths["model_card"].exists()
    assert paths["folds_csv"].exists()
    assert paths["stability"].exists()

    model_card = json.loads(paths["model_card"].read_text())
    assert model_card["card_version"] == "1.0"
    assert "validation_summary" in model_card
    assert "quality_gates" in model_card
    assert "deployment_readiness" in model_card

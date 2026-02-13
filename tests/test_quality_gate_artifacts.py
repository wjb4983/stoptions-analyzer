import copy

from reports import generate_cards, nightly_benchmark_report, quality_gates


def test_quality_gates_pass_with_repo_artifacts() -> None:
    scorecard = quality_gates._load_json(quality_gates.DEFAULT_SCORECARD_PATH)
    calibration = quality_gates._load_json(quality_gates.DEFAULT_CALIBRATION_PATH)
    baseline = quality_gates._load_json(quality_gates.DEFAULT_BASELINE_PATH)

    payload = quality_gates.evaluate_quality_gates(
        scorecard=scorecard,
        calibration_report=calibration,
        baseline_rows=baseline,
    )
    assert payload["all_gates_pass"] is True
    assert set(payload["gates"].keys()) == {
        "data_quality",
        "leakage_tests",
        "validation_integrity",
        "calibration",
        "friction_adjusted_performance",
    }


def test_quality_gates_fail_when_scorecard_fails() -> None:
    scorecard = quality_gates._load_json(quality_gates.DEFAULT_SCORECARD_PATH)
    calibration = quality_gates._load_json(quality_gates.DEFAULT_CALIBRATION_PATH)
    baseline = quality_gates._load_json(quality_gates.DEFAULT_BASELINE_PATH)

    broken = copy.deepcopy(scorecard)
    broken["promotion_gate"] = {"pass": False, "failed_critical_dimensions": ["robust_oos_performance"]}

    payload = quality_gates.evaluate_quality_gates(
        scorecard=broken,
        calibration_report=calibration,
        baseline_rows=baseline,
    )
    assert payload["all_gates_pass"] is False
    assert payload["gates"]["validation_integrity"]["pass"] is False


def test_generate_cards_and_nightly_comparison_shapes() -> None:
    scorecard = quality_gates._load_json(quality_gates.DEFAULT_SCORECARD_PATH)
    calibration = quality_gates._load_json(quality_gates.DEFAULT_CALIBRATION_PATH)
    baseline = quality_gates._load_json(quality_gates.DEFAULT_BASELINE_PATH)
    gates = quality_gates.evaluate_quality_gates(
        scorecard=scorecard,
        calibration_report=calibration,
        baseline_rows=baseline,
    )

    model_card = generate_cards.build_model_card(scorecard=scorecard, quality_gates=gates)
    strategy_card = generate_cards.build_strategy_card(baseline_rows=baseline, quality_gates=gates)

    assert model_card["card_version"] == "1.0"
    assert strategy_card["card_version"] == "1.0"
    assert strategy_card["merge_deploy_gates_passed"] is True

    comparison = nightly_benchmark_report.build_comparison(mainline=baseline, prior=baseline)
    assert comparison["scenarios_compared"] == len(baseline)
    assert comparison["mainline_beats_or_matches_prior_on_sharpe"] == len(baseline)

from __future__ import annotations

import json
from datetime import date

from ui.research_lab_page import ResearchLabPage


def _build_record(hypothesis_id: str) -> dict[str, object]:
    return {
        "hypothesis_id": hypothesis_id,
        "idea": {
            "title": "Test hypothesis",
            "description": "Test description",
            "submitter": "research_lab_ui",
            "submitted_at": date.today().isoformat(),
        },
        "economic_rationale": {"market_inefficiency": "test"},
        "data_requirements": {"assets": ["AAPL"]},
        "test_design": {"primary_test": "walk_forward"},
        "results_review": {
            "reviewer": "research_lab_ui",
            "status": "pending",
            "notes": "Awaiting execution",
        },
        "promotion_or_rejection": {
            "state": "promoted_to_experiment",
            "approval_status": "pending",
            "reason": "Clears minimum weighted rubric thresholds",
        },
        "decision": "accept",
        "decision_reason": "Clears minimum weighted rubric thresholds",
        "rubric": {
            "total_score": 4.1,
            "novelty": 4,
        },
        "lineage": {
            "hypothesis_id": hypothesis_id,
            "optimization_run_id": "opt_1",
            "walk_forward_run_id": "wf_1",
            "stress_run_id": "stress_1",
        },
    }


def _build_context() -> dict[str, object]:
    return {
        "tickers": ["AAPL", "MSFT"],
        "start_date": date(2024, 1, 1),
        "end_date": date(2024, 12, 31),
        "lookback": 90,
        "skip": 5,
        "costs_bps": 5.0,
    }


def test_write_experiment_skeleton_writes_research_project_manifest(tmp_path):
    page = object.__new__(ResearchLabPage)
    page._research_lab_dir = tmp_path / "research_lab"

    output_path = ResearchLabPage._write_experiment_skeleton(page, _build_record("hyp_abc"), _build_context())

    assert output_path == tmp_path / "research_lab" / "experiment_skeletons" / "hyp_abc" / "research_project.json"
    payload = json.loads(output_path.read_text(encoding="utf-8"))

    assert payload["schema_version"] == "1.0"
    assert payload["hypothesis"]["hypothesis_id"] == "hyp_abc"
    assert payload["run_references"][0]["step"] == "parameter_optimization"
    assert payload["gate_outcomes"][0]["gate"] == "rubric_score"
    assert payload["reviewer_comments"][0]["reviewer"] == "research_lab_ui"
    assert payload["promotion_history"][0]["state"] == "promoted_to_experiment"


def test_write_experiment_skeleton_appends_manifest_sections_without_duplication(tmp_path):
    page = object.__new__(ResearchLabPage)
    page._research_lab_dir = tmp_path / "research_lab"

    record = _build_record("hyp_repeat")
    context = _build_context()

    output_path = ResearchLabPage._write_experiment_skeleton(page, record, context)
    ResearchLabPage._write_experiment_skeleton(page, record, context)

    payload = json.loads(output_path.read_text(encoding="utf-8"))
    assert len(payload["run_references"]) == 3
    assert len(payload["gate_outcomes"]) == 1
    assert len(payload["reviewer_comments"]) == 1
    assert len(payload["promotion_history"]) == 1

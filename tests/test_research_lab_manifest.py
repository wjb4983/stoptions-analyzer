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
        "universe_filters": {
            "sector": {"include": ["Technology"], "exclude": ["Financials"]},
            "liquidity": {"min_adv": 1_000_000.0, "min_liquidity": 250_000.0},
            "price_market_cap": {"min_price": 5.0, "max_price": 500.0, "min_market_cap": 1_000_000_000.0, "max_market_cap": 5_000_000_000_000.0},
            "options_eligibility": {"min_open_interest": 500, "min_option_volume": 100, "min_days_to_expiration": 7, "require_weeklies": True},
        },
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
    assert payload["context"]["universe_filters"]["sector"]["include"] == ["Technology"]


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


def test_build_signal_grids_serializes_universe_filters_into_core_grid():
    page = object.__new__(ResearchLabPage)
    context = _build_context()
    config = type(
        "_Cfg",
        (),
        {"entry_signals": ["ts_momentum"], "exit_signals": ["none"]},
    )()

    entry_grid, exit_grid, core_grid = ResearchLabPage._build_signal_grids(page, context, config)

    assert entry_grid == {"ts_momentum": [{}]}
    assert exit_grid == {"none": [{}]}
    assert core_grid["universe_filters"][0]["options_eligibility"]["require_weeklies"] is True


class _Var:
    def __init__(self) -> None:
        self.value = ""

    def set(self, value: str) -> None:
        self.value = value


def test_extract_manifest_path_from_output_resolves_saved_outputs_line(tmp_path):
    page = object.__new__(ResearchLabPage)
    run_dir = tmp_path / "wf_run"
    run_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = run_dir / "manifest.json"
    manifest_path.write_text("{}", encoding="utf-8")

    output = f"Walk-forward complete. Saved outputs to: {run_dir}"
    extracted = ResearchLabPage._extract_manifest_path_from_output(page, output)

    assert extracted == manifest_path


def test_refresh_governance_dashboard_updates_summary_fields():
    page = object.__new__(ResearchLabPage)
    page._governance_gate_counts_var = _Var()
    page._governance_missing_checks_var = _Var()
    page._governance_promotion_ready_var = _Var()
    page._governance_approval_status_var = _Var()
    page._load_governance_payload_from_output = lambda _output: {
        "gate_checks": {"a": True, "b": False, "c": True},
        "missing_required_checks": ["b"],
        "is_promotion_ready": False,
        "approval_status": "pending",
        "promotion_state": "research",
    }

    ResearchLabPage._refresh_governance_dashboard_from_output(page, "Saved outputs to: /tmp/path")

    assert page._governance_gate_counts_var.value == "2 passed / 1 failed (3 total)"
    assert page._governance_missing_checks_var.value == "b"
    assert page._governance_promotion_ready_var.value == "Not ready"
    assert page._governance_approval_status_var.value == "pending (research)"

from __future__ import annotations

import json

from src.ui.research_lab_page import ResearchLabPage


def _page() -> ResearchLabPage:
    return object.__new__(ResearchLabPage)


def test_top_loss_contributors_card_aggregates_negative_drivers() -> None:
    page = _page()
    rows = [
        {"top_drivers": [{"feature": "trade_cost_bps", "attribution_score": -2.0}]},
        {"top_drivers": [{"feature": "trade_cost_bps", "attribution_score": -1.0}]},
        {"top_drivers": [{"feature": "running_pnl", "attribution_score": -0.5}]},
    ]

    card = page._build_top_loss_contributors_card(rows)

    assert "trade_cost_bps (3.00)" in card
    assert "running_pnl (0.50)" in card


def test_regime_failure_windows_card_groups_flagged_sequences() -> None:
    page = _page()
    rows = [
        {"timestamp": "2024-01-01T00:00:00", "red_flags": [{"flag": "x"}]},
        {"timestamp": "2024-01-01T00:01:00", "red_flags": [{"flag": "x"}]},
        {"timestamp": "2024-01-01T00:02:00", "red_flags": []},
        {"timestamp": "2024-01-01T00:03:00", "red_flags": [{"flag": "x"}]},
    ]
    regime_map = {
        "2024-01-01T00:00:00": "risk_off",
        "2024-01-01T00:01:00": "risk_off",
        "2024-01-01T00:03:00": "risk_on",
    }

    card = page._build_regime_failure_windows_card(rows, regime_map)

    assert "risk_off: 2 flagged trades" in card
    assert "risk_on: 1 flagged trades" in card


def test_slippage_sensitivity_card_participation_buckets() -> None:
    page = _page()
    rows = [
        {
            "fill_context": [
                {"requested_size": 10.0, "residual_size": 1.0, "participation_rate": 0.2},
                {"requested_size": 10.0, "residual_size": 2.0, "participation_rate": 0.4},
                {"requested_size": 10.0, "residual_size": 3.0, "participation_rate": 0.8},
            ]
        }
    ]

    card = page._build_slippage_sensitivity_card(rows)

    assert "low participation residual=10.00%" in card
    assert "mid participation residual=20.00%" in card
    assert "high participation residual=30.00%" in card


def test_load_regime_by_timestamp_from_regimes_json(tmp_path) -> None:
    page = _page()
    run_dir = tmp_path / "run"
    run_dir.mkdir()
    (run_dir / "regimes.json").write_text(
        json.dumps([
            {"timestamp": "2024-01-01T00:00:00", "regime": "risk_off"},
            {"timestamp": "2024-01-01T00:01:00", "regime": "risk_on"},
        ])
    )

    mapping = page._load_regime_by_timestamp(run_dir)

    assert mapping == {
        "2024-01-01T00:00:00": "risk_off",
        "2024-01-01T00:01:00": "risk_on",
    }

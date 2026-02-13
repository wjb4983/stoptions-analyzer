from __future__ import annotations

import json
from pathlib import Path

import numpy as np

from src.backtesting.scenario_toolkit import (
    ScenarioSpec,
    build_custom_scenarios,
    export_scenario_comparison_report,
    optimize_hedges_for_scenarios,
    project_strategy_under_scenarios,
)


def _fixture() -> tuple[np.ndarray, np.ndarray, list[str], dict[str, np.ndarray]]:
    returns = np.array(
        [
            [0.01, -0.002, 0.003],
            [-0.004, 0.005, -0.002],
            [0.006, -0.001, 0.004],
            [-0.02, -0.015, -0.01],
            [0.012, 0.008, 0.006],
            [0.004, -0.002, 0.001],
        ],
        dtype=float,
    )
    weights = np.array(
        [
            [0.3, -0.2, 0.1],
            [0.25, -0.15, 0.1],
            [0.2, -0.1, 0.05],
            [0.2, -0.1, 0.05],
            [0.18, -0.08, 0.06],
            [0.15, -0.05, 0.05],
        ],
        dtype=float,
    )
    sectors = ["tech", "financials", "energy"]
    liquidity = {
        "available_bar_volume": np.full_like(weights, 1_000_000.0),
        "spread_bps": np.full_like(weights, 12.0),
    }
    return weights, returns, sectors, liquidity


def test_build_custom_scenarios_supports_requested_types() -> None:
    specs = [
        ScenarioSpec(name="rate", scenario_type="rate_shock", params={"rate_shift": -0.003}),
        ScenarioSpec(name="vol", scenario_type="vol_shock", params={"jump_multiplier": 2.0}),
        ScenarioSpec(name="gap", scenario_type="gap_risk", params={"gap_jump": 0.05}),
        ScenarioSpec(
            name="rotation",
            scenario_type="sector_rotation",
            params={"favored_sectors": ["energy"], "penalized_sectors": ["tech"]},
        ),
        ScenarioSpec(name="crash_rebound", scenario_type="crash_rebound_path", params={"crash_periods": 2, "rebound_periods": 2}),
    ]
    payloads = build_custom_scenarios(specs=specs, n_assets=3, sector_by_asset=["tech", "financials", "energy"])
    assert len(payloads) == 5
    assert {p["scenario"].name for p in payloads} == {"rate", "vol", "gap", "rotation", "crash_rebound"}


def test_project_strategy_under_scenarios_emits_pnl_risk_and_exposure() -> None:
    weights, returns, sectors, liquidity = _fixture()
    specs = [
        ScenarioSpec(name="rate", scenario_type="rate_shock", params={"rate_shift": -0.002}),
        ScenarioSpec(name="rotation", scenario_type="sector_rotation", params={"favored_sectors": ["financials"], "penalized_sectors": ["tech"]}),
    ]
    payloads = build_custom_scenarios(specs=specs, n_assets=3, sector_by_asset=sectors)
    rows = project_strategy_under_scenarios(
        base_weights=weights,
        asset_returns=returns,
        scenario_payloads=payloads,
        sector_by_asset=sectors,
        liquidity_context=liquidity,
    )
    assert len(rows) == 2
    assert all("pnl_total" in row and "max_drawdown" in row and "gross_exposure" in row for row in rows)
    assert all(set(row["sector_exposure"].keys()) == set(sectors) for row in rows)


def test_optimize_hedges_for_selected_scenarios_returns_solution() -> None:
    weights, returns, sectors, _ = _fixture()
    payloads = build_custom_scenarios(
        specs=[
            ScenarioSpec(name="vol", scenario_type="vol_shock", params={"jump_multiplier": 2.1}),
            ScenarioSpec(name="crash", scenario_type="crash_rebound_path", params={"crash_periods": 3}),
        ],
        n_assets=3,
        sector_by_asset=sectors,
    )
    hedge_returns = np.column_stack([returns[:, 0], -returns[:, 1]])
    result = optimize_hedges_for_scenarios(
        base_weights=weights,
        asset_returns=returns,
        scenario_payloads=payloads,
        hedge_returns=hedge_returns,
        selected_scenarios=["vol"],
        n_trials=80,
        random_seed=7,
    )
    assert result["selected_scenarios"] == ["vol"]
    assert len(result["optimal_hedge_weights"]) == 2
    assert len(result["scenario_breakdown"]) == 1


def test_export_scenario_comparison_report_writes_artifacts(tmp_path: Path) -> None:
    rows = [
        {
            "scenario": "vol",
            "scenario_type": "vol_shock",
            "pnl_total": -0.03,
            "max_drawdown": -0.2,
            "cvar_95": 0.04,
            "liquidity_breach_count": 1,
            "gross_exposure": 0.35,
            "net_exposure": 0.12,
            "sector_exposure": {"tech": 0.2},
        }
    ]
    optimization = {"objective": 0.5, "selected_scenarios": ["vol"], "optimal_hedge_weights": [0.1, -0.2]}
    outputs = export_scenario_comparison_report(
        scenario_projection_rows=rows,
        hedge_optimization_result=optimization,
        output_dir=tmp_path,
        basename="meeting_pack",
    )

    assert Path(outputs["json"]).exists()
    assert Path(outputs["csv"]).exists()
    assert Path(outputs["markdown"]).exists()

    payload = json.loads(Path(outputs["json"]).read_text())
    assert payload["hedge_optimization"]["selected_scenarios"] == ["vol"]

from __future__ import annotations

import numpy as np

from src.analysis.reporting import build_scenario_attribution_and_guardrails
from src.backtesting import cache_runner


def _fixture_returns() -> np.ndarray:
    return np.array([0.01, -0.004, 0.006, -0.002, 0.003, -0.001, 0.002], dtype=float)


def _fixture_prices() -> np.ndarray:
    return np.array(
        [
            [100.0, 80.0],
            [101.0, 79.5],
            [100.5, 80.2],
            [101.2, 79.8],
            [101.0, 80.5],
            [100.8, 80.1],
            [101.1, 80.3],
        ],
        dtype=float,
    )


def test_stress_scenario_definitions_include_expected_shocks() -> None:
    definitions = cache_runner._build_stress_scenario_definitions(
        timestamps=np.arange(_fixture_returns().size, dtype=np.int64),
        returns=_fixture_returns(),
    )
    names = {row["name"] for row in definitions}
    assert "historical_recent_window" in names
    assert "historical_first_window" in names
    assert "synthetic_returns_crash" in names
    assert "synthetic_spread_widening" in names
    assert "synthetic_liquidity_drought" in names
    assert "synthetic_borrow_unavailable" in names
    assert "synthetic_correlation_breakdown" in names


def test_scenario_wrappers_are_deterministic_for_fixture_inputs() -> None:
    definitions = cache_runner._build_stress_scenario_definitions(
        timestamps=np.arange(_fixture_returns().size, dtype=np.int64),
        returns=_fixture_returns(),
    )
    first = cache_runner._run_stress_scenario_wrappers(
        returns=_fixture_returns(),
        prices=_fixture_prices(),
        scenario_definitions=definitions,
    )
    second = cache_runner._run_stress_scenario_wrappers(
        returns=_fixture_returns(),
        prices=_fixture_prices(),
        scenario_definitions=definitions,
    )
    assert first == second


def test_scenario_attribution_and_guardrails_emit_rows() -> None:
    baseline = {"sharpe": 1.1, "max_drawdown": -0.08, "total_return": 0.12}
    scenario_results = [
        {
            "name": "synthetic_returns_crash",
            "pnl_total": -0.03,
            "metrics": {"sharpe": -0.2, "max_drawdown": -0.35, "total_return": -0.07},
        },
        {
            "name": "synthetic_liquidity_drought",
            "pnl_total": 0.01,
            "metrics": {"sharpe": 0.5, "max_drawdown": -0.14, "total_return": 0.04},
        },
    ]

    payload = build_scenario_attribution_and_guardrails(
        baseline_metrics=baseline,
        scenario_results=scenario_results,
    )

    attribution = payload["scenario_attribution"]
    guardrails = payload["scenario_guardrails"]
    assert len(attribution) == 2
    assert len(guardrails) == 2
    assert any(not row["passed"] for row in guardrails)


def test_stress_scenario_controls_enable_regime_replay_and_overlays() -> None:
    definitions = cache_runner._build_stress_scenario_definitions(
        timestamps=np.arange(_fixture_returns().size, dtype=np.int64),
        returns=_fixture_returns(),
        controls={
            "enable_historical_replay_regimes": True,
            "historical_replay_window_bars": 3,
            "synthetic_jump_magnitude": 0.05,
            "overlay_spread_multiplier": 3.0,
            "overlay_liquidity_multiplier": 0.3,
        },
    )
    names = {row["name"] for row in definitions}
    assert "historical_volatility_shock_window" in names
    assert "synthetic_jump_cluster" in names
    assert "synthetic_liquidity_spread_overlay" in names


def test_stress_gate_summary_flags_failures() -> None:
    payload = {
        "scenario_guardrails": [
            {"scenario": "a", "passed": True},
            {"scenario": "b", "passed": False},
        ]
    }
    summary = cache_runner._stress_gate_summary(payload)
    assert summary["stress_passed"] is False
    assert summary["stress_failed_scenarios"] == 1
    assert summary["stress_total_scenarios"] == 2


def test_stress_scenario_definitions_include_synthetic_regime_switch_paths() -> None:
    definitions = cache_runner._build_stress_scenario_definitions(
        timestamps=np.arange(_fixture_returns().size, dtype=np.int64),
        returns=_fixture_returns(),
        controls={"synthetic_path_count": 2, "synthetic_path_seed": 7},
    )
    synthetic = [row for row in definitions if str(row.get("type")) == "synthetic_path"]
    assert len(synthetic) == 2
    assert all("regime_switches" in dict(row.get("stress_characteristics", {})) for row in synthetic)


def test_stress_gate_summary_emits_survivability_and_failures() -> None:
    payload = {
        "scenario_guardrails": [
            {"scenario": "a", "passed": True},
            {"scenario": "b", "passed": False},
        ],
        "scenario_attribution": [
            {"delta_max_drawdown": -0.2, "delta_sharpe": -0.5, "delta_total_return": -0.1},
            {"delta_max_drawdown": -0.1, "delta_sharpe": -0.2, "delta_total_return": -0.05},
        ],
    }
    summary = cache_runner._stress_gate_summary(payload, controls={"stress_survivability_min": 0.8})
    assert summary["stress_failed_scenario_names"] == ["b"]
    assert 0.0 <= summary["stress_survivability_score"] <= 1.0
    assert summary["stress_model_gate_passed"] is False

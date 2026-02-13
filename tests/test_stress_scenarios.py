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

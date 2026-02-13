from __future__ import annotations

import numpy as np

from src.backtesting.allocation_optimizer import AllocationOptimizationConfig, optimize_allocation


def _sample_inputs() -> tuple[np.ndarray, np.ndarray, np.ndarray, list[str], list[str]]:
    mu = np.array([0.10, 0.08, 0.04, 0.03], dtype=float)
    cov = np.array(
        [
            [0.08, 0.02, 0.01, 0.00],
            [0.02, 0.07, 0.00, 0.01],
            [0.01, 0.00, 0.06, 0.02],
            [0.00, 0.01, 0.02, 0.05],
        ],
        dtype=float,
    )
    scenarios = np.random.default_rng(7).normal(0.0, 0.02, size=(24, 12, 4))
    sectors = ["tech", "tech", "fin", "fin"]
    strategies = ["carry", "carry", "trend", "trend"]
    return mu, cov, scenarios, sectors, strategies


def test_mean_variance_turnover_sector_and_regime_leverage() -> None:
    mu, cov, scenarios, sectors, strategies = _sample_inputs()
    prev = np.array([0.30, -0.20, 0.30, -0.20], dtype=float)
    result = optimize_allocation(
        expected_returns=mu,
        covariance=cov,
        previous_weights=prev,
        uncertainty=np.array([0.06, 0.04, 0.02, 0.01], dtype=float),
        sectors=sectors,
        instrument_to_strategy=strategies,
        regime="risk_off",
        return_scenarios=scenarios,
        config=AllocationOptimizationConfig(
            objective="mean_variance",
            turnover_penalty=2.5,
            sector_neutrality=True,
            max_drawdown=0.15,
            leverage_caps_by_regime={"risk_off": 0.6},
        ),
    )

    assert float(np.sum(np.abs(result.weights))) <= 0.6 + 1e-6
    assert result.diagnostics["regime_leverage_cap"] == 0.6
    sector_exp = result.diagnostics["sector_net_exposure"]
    assert isinstance(sector_exp, dict)
    assert abs(float(sector_exp["tech"])) < 1e-2
    assert abs(float(sector_exp["fin"])) < 1e-2
    assert "gross_leverage_cap" in result.shadow_prices


def test_cvar_objective_and_drawdown_guardrail() -> None:
    mu, cov, scenarios, sectors, strategies = _sample_inputs()
    result = optimize_allocation(
        expected_returns=mu,
        covariance=cov,
        previous_weights=np.zeros(4, dtype=float),
        sectors=sectors,
        instrument_to_strategy=strategies,
        return_scenarios=scenarios,
        config=AllocationOptimizationConfig(
            objective="cvar_minimization",
            max_drawdown=0.10,
            default_leverage_cap=0.9,
            cvar_confidence=0.9,
        ),
    )

    assert float(result.diagnostics["cvar"]) >= 0.0
    assert float(result.diagnostics["max_drawdown"]) >= -0.10 - 1e-6


def test_risk_parity_strategy_level_outputs_strategy_allocations() -> None:
    mu, cov, scenarios, sectors, strategies = _sample_inputs()
    result = optimize_allocation(
        expected_returns=mu,
        covariance=cov,
        previous_weights=np.zeros(4, dtype=float),
        sectors=sectors,
        instrument_to_strategy=strategies,
        return_scenarios=scenarios,
        config=AllocationOptimizationConfig(
            objective="risk_parity",
            level="strategy",
            strategy_gross_caps={"carry": 0.35, "trend": 0.35},
            default_leverage_cap=0.7,
        ),
    )

    assert set(result.strategy_weights) == {"carry", "trend"}
    assert float(np.sum(np.abs(result.weights))) <= 0.7 + 1e-6
    assert "turnover_penalty" in result.shadow_prices

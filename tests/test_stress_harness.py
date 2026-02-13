from __future__ import annotations

import numpy as np

from src.analysis.reporting import export_stress_harness_report
from src.backtesting.portfolio import replay_weights_under_stress
from src.backtesting.vectorized import (
    StressCorrelationBreak,
    StressLiquidityDrought,
    StressScenario,
    StressShockVector,
    StressVolatilityJump,
    apply_stress_scenario,
)


def test_apply_stress_scenario_adjusts_returns_and_liquidity() -> None:
    asset_returns = np.array(
        [
            [0.01, -0.005],
            [0.005, 0.002],
            [-0.02, -0.01],
            [0.004, 0.003],
        ],
        dtype=float,
    )
    liquidity = {
        "available_bar_volume": np.full_like(asset_returns, 100.0),
        "spread_bps": np.full_like(asset_returns, 5.0),
        "max_participation_per_bar": np.full_like(asset_returns, 0.25),
    }
    scenario = StressScenario(
        name="combo",
        shock_vector=StressShockVector(
            returns_multiplier_by_asset=(1.2, 0.8),
            returns_shift_by_asset=(-0.001, -0.0005),
        ),
        correlation_break=StressCorrelationBreak(break_strength=0.5, idiosyncratic_amplifier=1.3),
        volatility_jump=StressVolatilityJump(jump_multiplier=1.5, trigger_quantile=0.5),
        liquidity_drought=StressLiquidityDrought(
            available_volume_multiplier=0.4,
            spread_multiplier=2.0,
            max_participation_multiplier=0.5,
        ),
    )

    out = apply_stress_scenario(asset_returns=asset_returns, scenario=scenario, liquidity_context=liquidity)

    shocked = out["asset_returns"]
    assert shocked.shape == asset_returns.shape
    assert np.any(np.abs(shocked - asset_returns) > 1e-10)
    assert np.allclose(out["liquidity"]["available_bar_volume"], 40.0)
    assert np.allclose(out["liquidity"]["spread_bps"], 10.0)
    assert np.allclose(out["liquidity"]["max_participation_per_bar"], 0.125)


def test_replay_weights_under_stress_emits_core_metrics() -> None:
    base_weights = np.array(
        [
            [0.5, -0.2],
            [0.5, -0.2],
            [0.45, -0.15],
            [0.40, -0.10],
        ],
        dtype=float,
    )
    stressed_returns = np.array(
        [
            [0.01, -0.01],
            [-0.03, -0.01],
            [-0.02, 0.005],
            [0.01, 0.002],
        ],
        dtype=float,
    )
    avail = np.full_like(base_weights, 0.04)
    spread = np.full_like(base_weights, 35.0)

    result = replay_weights_under_stress(
        base_weights=base_weights,
        stressed_asset_returns=stressed_returns,
        stressed_available_volume=avail,
        stressed_spread_bps=spread,
        liquidity_turnover_threshold=0.02,
        scenario_name="liquidity_crunch",
    )

    assert result.scenario == "liquidity_crunch"
    assert result.portfolio_returns.shape[0] == base_weights.shape[0]
    assert result.max_drawdown <= 0.0
    assert result.cvar_95 >= 0.0
    assert result.attribution_by_asset.shape[0] == base_weights.shape[1]
    assert result.liquidity_breach_count >= 1


def test_export_stress_harness_report_picks_worst_case() -> None:
    report = export_stress_harness_report(
        scenario_replays=[
            {
                "scenario": "mild",
                "max_drawdown": -0.10,
                "cvar_95": 0.03,
                "liquidity_breach_count": 2,
                "attribution_by_asset": [0.01, -0.02],
                "worst_bar": 3,
                "worst_bar_contribution_by_asset": [0.002, -0.009],
            },
            {
                "scenario": "severe",
                "max_drawdown": -0.30,
                "cvar_95": 0.08,
                "liquidity_breach_count": 5,
                "attribution_by_asset": [-0.02, -0.04],
                "worst_bar": 1,
                "worst_bar_contribution_by_asset": [-0.01, -0.03],
            },
        ]
    )

    assert len(report["scenario_reports"]) == 2
    assert report["worst_case_decomposition"]["scenario"] == "severe"

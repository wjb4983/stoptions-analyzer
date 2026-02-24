from __future__ import annotations

import numpy as np
import pytest

from src.backtesting.event_driven import EventDrivenBacktester, Order
from src.backtesting.strategies import __all__ as strategy_public_api
from src.backtesting.strategies.alpha_model import probability_calibrated_position_size
from src.backtesting.strategies.dsl import list_template_names
from src.backtesting.strategies import HedgeRebalanceConfig, VolHedgingPolicy
from src.backtesting.strategies.greek_targets import (
    GreekNeutralTargetRequest,
    build_greek_neutral_targets,
    compute_aggregate_greek_exposures,
)
from src.backtesting.vectorized import backtest_vectorized
from src.ui.backtesting_page import STRATEGIES


class _TargetPositionStrategy:
    def __init__(self, symbol: str, target_positions: np.ndarray) -> None:
        self.symbol = symbol
        self.target_positions = target_positions
        self.idx = 0

    def on_bar(self, bar: dict[str, float | str], portfolio) -> list[Order]:
        target = float(self.target_positions[self.idx])
        current = portfolio.positions.get(self.symbol).quantity if self.symbol in portfolio.positions else 0.0
        delta = target - current
        self.idx += 1
        if delta > 0:
            return [Order(symbol=self.symbol, quantity=delta, side="buy", timestamp=str(bar.get("timestamp")))]
        if delta < 0:
            return [Order(symbol=self.symbol, quantity=abs(delta), side="sell", timestamp=str(bar.get("timestamp")))]
        return []


@pytest.fixture
def strategy_discovery_catalog() -> dict[str, object]:
    return {
        "ui_strategies": tuple(STRATEGIES),
        "dsl_templates": list_template_names(),
        "public_strategy_api": tuple(strategy_public_api),
    }


def test_signal_generation_invariants_delta_neutral_and_gross_normalized() -> None:
    base = np.array([[0.5, 0.3, -0.2], [0.1, -0.2, 0.7]], dtype=float)
    delta = np.array([[1.0, 0.5, 1.2], [0.8, 1.3, 0.4]], dtype=float)

    targets = build_greek_neutral_targets(
        GreekNeutralTargetRequest(base_targets=base, delta=delta, neutralize=("delta",))
    )

    assert targets.shape == base.shape
    assert np.allclose(np.sum(targets * delta, axis=1), 0.0, atol=1e-10)
    assert np.allclose(np.sum(np.abs(targets), axis=1), 1.0, atol=1e-10)


def test_boundary_conditions_zero_iv_missing_greeks_and_sparse_rows() -> None:
    base = np.array([[1.0, -1.0], [0.0, 0.0], [0.4, -0.4]], dtype=float)
    delta = np.array([[1.0, 1.0], [1.0, 1.0], [0.0, 0.0]], dtype=float)

    targets = build_greek_neutral_targets(
        GreekNeutralTargetRequest(
            base_targets=base,
            delta=delta,
            vega=np.zeros_like(base),  # zero IV proxy
            gamma=None,  # missing greek should be ignored
            neutralize=("delta", "gamma", "vega"),
        )
    )

    assert np.allclose(targets[1], 0.0)
    assert np.allclose(np.sum(targets[0] * delta[0]), 0.0, atol=1e-10)
    assert np.allclose(targets[2], np.array([0.5, -0.5]))


def test_position_sizing_compatibility_with_probability_calibration() -> None:
    probabilities = np.array([0.2, 0.5, 0.9], dtype=float)
    sized = probability_calibrated_position_size(probabilities, max_leverage=1.5, gamma=1.0)
    base_targets = np.column_stack([sized, -0.5 * sized])
    delta = np.array([[1.0, 1.0], [1.2, 0.8], [0.9, 1.1]], dtype=float)

    neutral = build_greek_neutral_targets(
        GreekNeutralTargetRequest(base_targets=base_targets, delta=delta, neutralize=("delta",))
    )

    assert neutral.shape == base_targets.shape
    assert np.max(np.abs(neutral)) <= 1.0
    exposures = compute_aggregate_greek_exposures(
        positions=np.ones_like(neutral),
        portfolio_weights=np.array([0.6, 0.4]),
        sensitivities={"delta": delta, "vega": np.zeros_like(delta)},
    )
    assert set(exposures.keys()) == {"delta", "gamma", "vega", "theta"}


def test_vectorized_and_event_driven_parity_with_same_target_config() -> None:
    prices = np.array([100.0, 101.0, 103.0, 102.0, 104.0], dtype=float)
    base = np.array([[0.0], [0.7], [0.4], [-0.5], [0.2]], dtype=float)
    delta = np.ones_like(base)
    signals = build_greek_neutral_targets(
        GreekNeutralTargetRequest(base_targets=base, delta=delta, neutralize=("delta", "vega"))
    )[:, 0]

    vec = backtest_vectorized(prices=prices, signals=signals, initial_equity=1_000.0)

    bars = [
        {"open": float(price), "timestamp": f"2024-01-0{i+1}"}
        for i, price in enumerate(prices)
    ]
    ev = EventDrivenBacktester(initial_cash=1_000.0)
    strategy = _TargetPositionStrategy(symbol="ABC", target_positions=signals)
    event = ev.run(bars, strategy)

    final_qty = event.portfolio.positions.get("ABC").quantity if "ABC" in event.portfolio.positions else 0.0
    event_final_equity = event.portfolio.cash + final_qty * prices[-1]
    vectorized_final_equity = float(vec.equity_curve[-1])

    assert event_final_equity == pytest.approx(vectorized_final_equity, rel=1e-6, abs=1e-6)
    event_total_return = event_final_equity / 1_000.0 - 1.0
    vectorized_total_return = vectorized_final_equity / 1_000.0 - 1.0
    assert event_total_return == pytest.approx(vectorized_total_return, rel=1e-6, abs=1e-6)


def test_strategy_discovery_and_cli_listing_include_expected_entries(strategy_discovery_catalog: dict[str, object]) -> None:
    ui_strategies = strategy_discovery_catalog["ui_strategies"]
    dsl_templates = strategy_discovery_catalog["dsl_templates"]
    public_strategy_api = strategy_discovery_catalog["public_strategy_api"]

    assert "xsmom" in ui_strategies
    assert "ts_momentum_core" in dsl_templates
    assert "GreekNeutralTargetRequest" in public_strategy_api
    assert "build_greek_neutral_targets" in public_strategy_api


def test_vol_hedging_policy_reduces_delta_exposure_and_reports_drag() -> None:
    prices = np.array([
        [100.0, 100.0],
        [101.0, 99.0],
        [102.0, 98.0],
        [103.0, 97.0],
        [104.0, 96.0],
    ])
    signals = np.array([
        [0.0, 0.0],
        [0.9, 0.0],
        [0.9, 0.0],
        [0.9, 0.0],
        [0.9, 0.0],
    ])
    delta = np.array([
        [1.0, -1.0],
        [1.0, -1.0],
        [1.0, -1.0],
        [1.0, -1.0],
        [1.0, -1.0],
    ])

    unhedged = backtest_vectorized(prices=prices, signals=signals, greek_sensitivities={"delta": delta})
    hedged = backtest_vectorized(
        prices=prices,
        signals=signals,
        greek_sensitivities={"delta": delta},
        hedging_policy=VolHedgingPolicy(
            hedge_delta=True,
            rebalance=HedgeRebalanceConfig(every_n_bars=1),
            max_hedge_notional=1.0,
            max_hedge_turnover=2.0,
        ),
    )

    unhedged_delta = np.asarray(unhedged.cost_breakdown["greek_diagnostics"]["snapshots"]["delta"], dtype=float)
    hedged_delta = np.asarray(hedged.cost_breakdown["greek_diagnostics"]["snapshots"]["delta"], dtype=float)
    assert np.max(np.abs(hedged_delta[1:])) < np.max(np.abs(unhedged_delta[1:]))

    assert "hedge_cost" in hedged.cost_breakdown
    assert "hedge_pnl" in hedged.cost_breakdown
    assert "hedge_cost" in hedged.cost_breakdown["totals"]
    assert np.isfinite(float(hedged.cost_breakdown["totals"]["hedge_pnl"]))


def test_vol_hedging_respects_margin_limits() -> None:
    prices = np.full((5, 2), 100.0, dtype=float)
    signals = np.array([
        [0.0, 0.0],
        [3.0, 0.0],
        [3.0, 0.0],
        [3.0, 0.0],
        [3.0, 0.0],
    ])
    delta = np.ones_like(signals)

    result = backtest_vectorized(
        prices=prices,
        signals=signals,
        greek_sensitivities={"delta": delta},
        hedging_policy=VolHedgingPolicy(hedge_delta=True, rebalance=HedgeRebalanceConfig(every_n_bars=1)),
        margin_schedule_by_asset=np.array([0.95, 0.95]),
        stress_addon_by_asset=np.array([0.3, 0.3]),
        concentration_addon=0.2,
    )

    margin_util = np.asarray(result.cost_breakdown["account_state"]["margin_utilization"], dtype=float)
    assert np.all(margin_util <= 1.0 + 1e-9)

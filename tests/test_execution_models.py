from __future__ import annotations

import json

import numpy as np

from src.backtesting.execution import (
    AssetClassCarryCost,
    ParticipationImpactSlippage,
    SpreadSlippage,
    VolatilityScaledSlippage,
    calibrate_impact_coefficient_bps,
    load_impact_calibration_buckets,
)
from src.backtesting.vectorized import backtest_vectorized


def test_spread_slippage_is_side_aware() -> None:
    model = SpreadSlippage(spread_bps=10.0)
    buy = model.apply(100.0, 1.0)
    sell = model.apply(100.0, -1.0)
    assert buy > 100.0
    assert sell < 100.0
    assert np.isclose((buy - 100.0) / 100.0, (100.0 - sell) / 100.0)


def test_spread_slippage_increases_with_worse_queue_rank() -> None:
    model = SpreadSlippage(spread_bps=10.0)
    front = model.apply(100.0, 1.0, {"queue_rank_proxy": 0.0})
    back = model.apply(100.0, 1.0, {"queue_rank_proxy": 1.0})
    assert back > front


def test_participation_impact_cost_monotonicity() -> None:
    prices = np.array([100.0, 100.0, 100.0, 100.0])
    low_signals = np.array([0.0, 0.1, 0.1, 0.1])
    high_signals = np.array([0.0, 0.8, 0.8, 0.8])
    model = ParticipationImpactSlippage(base_bps=0.0, impact_coefficient_bps=50.0, participation_exponent=1.0)
    common_kwargs = {
        "slippage_model": model,
        "volumes": np.full_like(prices, 1.0),
        "adv": np.full_like(prices, 1.0),
    }

    low = backtest_vectorized(prices, low_signals, **common_kwargs)
    high = backtest_vectorized(prices, high_signals, **common_kwargs)

    assert high.cost_breakdown["totals"]["slippage"] > low.cost_breakdown["totals"]["slippage"]


def test_participation_impact_uses_realized_participation_when_provided() -> None:
    model = ParticipationImpactSlippage(base_bps=0.0, impact_coefficient_bps=100.0, participation_exponent=1.0)
    low = model.apply(100.0, 10.0, {"realized_participation": 0.1, "queue_rank_proxy": 0.0})
    high = model.apply(100.0, 10.0, {"realized_participation": 0.4, "queue_rank_proxy": 0.0})
    assert high > low


def test_volatility_scaled_slippage_increases_with_volatility() -> None:
    model = VolatilityScaledSlippage(base_bps=5.0, target_volatility=0.01)
    low = model.apply(100.0, 1.0, {"volatility": 0.01})
    high = model.apply(100.0, 1.0, {"volatility": 0.05})
    assert high > low


def test_volatility_scaled_slippage_penalizes_queue_and_participation() -> None:
    model = VolatilityScaledSlippage(base_bps=5.0, target_volatility=0.01)
    relaxed = model.apply(100.0, 1.0, {"volatility": 0.02, "realized_participation": 0.0, "queue_rank_proxy": 0.0})
    stressed = model.apply(100.0, 1.0, {"volatility": 0.02, "realized_participation": 0.5, "queue_rank_proxy": 1.0})
    assert stressed > relaxed


def test_calibration_helpers_roundtrip(tmp_path) -> None:
    buckets = [
        {"participation": 0.1, "slippage_bps": 2.0, "count": 10},
        {"participation": 0.2, "slippage_bps": 4.0, "count": 10},
    ]
    path = tmp_path / "impact_buckets.json"
    path.write_text(json.dumps(buckets))
    loaded = load_impact_calibration_buckets(path)
    coeff = calibrate_impact_coefficient_bps(loaded)
    model = ParticipationImpactSlippage.from_calibration_buckets(loaded)
    assert np.isclose(coeff, 20.0)
    assert np.isclose(model.impact_coefficient_bps, 20.0)


def test_asset_class_carry_respects_class_rates() -> None:
    model = AssetClassCarryCost(
        asset_classes=["equity", "etf"],
        annual_short_borrow_rates={"equity": 0.06, "etf": 0.01},
        annual_long_financing_rates={"equity": 0.02, "etf": 0.00},
        periods_per_year=252.0,
    )
    positions = np.array([[1.0, -1.0]])
    carry = model.calculate(positions)
    assert carry.shape == positions.shape
    assert carry[0, 0] > carry[0, 1]

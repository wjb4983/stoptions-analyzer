from __future__ import annotations

import numpy as np

from src.backtesting.execution import (
    AssetClassCarryCost,
    ParticipationImpactSlippage,
    SpreadSlippage,
    VolatilityScaledSlippage,
)
from src.backtesting.vectorized import backtest_vectorized


def test_spread_slippage_is_side_aware() -> None:
    model = SpreadSlippage(spread_bps=10.0)
    buy = model.apply(100.0, 1.0)
    sell = model.apply(100.0, -1.0)
    assert buy > 100.0
    assert sell < 100.0
    assert np.isclose((buy - 100.0) / 100.0, (100.0 - sell) / 100.0)


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


def test_volatility_scaled_slippage_increases_with_volatility() -> None:
    model = VolatilityScaledSlippage(base_bps=5.0, target_volatility=0.01)
    low = model.apply(100.0, 1.0, {"volatility": 0.01})
    high = model.apply(100.0, 1.0, {"volatility": 0.05})
    assert high > low


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

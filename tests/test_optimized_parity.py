from __future__ import annotations

import numpy as np

from src.analysis.time_series.momentum import (
    MomentumHyperparameters,
    TimeSeriesMomentumSettings,
    build_time_series_momentum_arrays,
)
from src.backtesting.execution import BpsSlippage, FixedCommission, ShortBorrowCost
from src.backtesting.vectorized import backtest_vectorized

# Numerical tolerance policy for parity checks:
# - absolute tolerance: 1e-10 for deterministic vectorized terms
# - relative tolerance: 1e-8 for compounded equity curves
ABS_TOL = 1e-10
REL_TOL = 1e-8


def test_time_series_momentum_reference_vs_optimized_parity() -> None:
    rng = np.random.default_rng(11)
    prices = (100.0 * np.cumprod(1.0 + rng.normal(0.0002, 0.02, size=2_000))).tolist()
    settings = TimeSeriesMomentumSettings(
        hyperparameters=MomentumHyperparameters(
            lookback_days=63,
            skip_days=1,
            vol_window_days=20,
            target_volatility=0.2,
            max_leverage=1.5,
        )
    )

    ref = build_time_series_momentum_arrays(prices, settings, mode="reference")
    opt = build_time_series_momentum_arrays(prices, settings, mode="optimized")

    assert np.allclose(ref.raw_score, opt.raw_score, rtol=REL_TOL, atol=ABS_TOL, equal_nan=True)
    assert np.allclose(ref.position_signal, opt.position_signal, rtol=REL_TOL, atol=ABS_TOL)
    assert np.allclose(ref.confidence, opt.confidence, rtol=REL_TOL, atol=ABS_TOL)
    assert np.allclose(ref.scaled_weight, opt.scaled_weight, rtol=REL_TOL, atol=ABS_TOL)
    assert np.allclose(ref.tradable_position, opt.tradable_position, rtol=REL_TOL, atol=ABS_TOL)


def test_backtest_reference_vs_optimized_parity() -> None:
    rng = np.random.default_rng(12)
    n_periods, n_assets = 1_000, 16
    returns = rng.normal(0.0001, 0.01, size=(n_periods, n_assets))
    prices = 100.0 * np.cumprod(1.0 + returns, axis=0)
    signals = np.sign(rng.normal(size=(n_periods, n_assets)))
    weights = np.full(n_assets, 1.0 / n_assets)

    kwargs = dict(
        prices=prices,
        signals=signals,
        slippage_model=BpsSlippage(7.5),
        fee_model=FixedCommission(0.0003),
        borrow_cost_model=ShortBorrowCost(0.03),
        weights=weights,
        initial_equity=1.0,
    )

    ref = backtest_vectorized(**kwargs, execution_mode="reference")
    opt = backtest_vectorized(**kwargs, execution_mode="optimized")

    assert np.allclose(ref.positions, opt.positions, rtol=REL_TOL, atol=ABS_TOL)
    assert np.allclose(ref.trades, opt.trades, rtol=REL_TOL, atol=ABS_TOL)
    assert np.allclose(ref.returns, opt.returns, rtol=REL_TOL, atol=ABS_TOL)
    assert np.allclose(ref.turnover, opt.turnover, rtol=REL_TOL, atol=ABS_TOL)
    assert np.allclose(ref.equity_curve, opt.equity_curve, rtol=REL_TOL, atol=ABS_TOL)
    assert np.allclose(ref.cost_breakdown["total"], opt.cost_breakdown["total"], rtol=REL_TOL, atol=ABS_TOL)

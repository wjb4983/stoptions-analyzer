"""Performance profiling helpers for reference vs optimized execution paths."""

from __future__ import annotations

import cProfile
import pstats
from dataclasses import dataclass
from io import StringIO

import numpy as np

from .execution import BpsSlippage, FixedCommission, ShortBorrowCost
from .vectorized import backtest_vectorized


@dataclass(frozen=True)
class ProfileSummary:
    mode: str
    output: str


def profile_backtest_hotspots(n_periods: int = 20_000, n_assets: int = 128) -> list[ProfileSummary]:
    rng = np.random.default_rng(42)
    returns = rng.normal(loc=0.0002, scale=0.01, size=(n_periods, n_assets))
    prices = 100.0 * np.cumprod(1.0 + returns, axis=0)
    signals = np.sign(rng.normal(size=(n_periods, n_assets)))
    weights = np.full(n_assets, 1.0 / n_assets)

    outputs: list[ProfileSummary] = []
    for mode in ("reference", "optimized"):
        pr = cProfile.Profile()
        pr.enable()
        backtest_vectorized(
            prices,
            signals,
            slippage_model=BpsSlippage(5.0),
            fee_model=FixedCommission(0.0005),
            borrow_cost_model=ShortBorrowCost(0.03),
            weights=weights,
            execution_mode=mode,
        )
        pr.disable()

        buf = StringIO()
        stats = pstats.Stats(pr, stream=buf).sort_stats("cumtime")
        stats.print_stats(15)
        outputs.append(ProfileSummary(mode=mode, output=buf.getvalue()))
    return outputs

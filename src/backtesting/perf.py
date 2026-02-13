"""Performance profiling helpers for reference vs optimized execution paths."""

from __future__ import annotations

import cProfile
import pstats
from dataclasses import dataclass
from io import StringIO
from statistics import NormalDist

import numpy as np

from .execution import BpsSlippage, FixedCommission, ShortBorrowCost
from .vectorized import backtest_vectorized


@dataclass(frozen=True)
class ProfileSummary:
    mode: str
    output: str


def deflated_sharpe_ratio(
    *,
    observed_sharpe: float,
    n_returns: int,
    skew: float,
    kurtosis: float,
    n_trials: int,
) -> float:
    """Compute Deflated Sharpe Ratio probability after multiple testing correction."""
    if n_returns <= 2:
        return 0.0

    sr = float(observed_sharpe)
    variance_term = 1.0 - skew * sr + ((kurtosis - 1.0) / 4.0) * (sr**2)
    sr_std = np.sqrt(max(variance_term, 1e-12) / max(1, n_returns - 1))

    effective_trials = max(1, int(n_trials))
    if effective_trials == 1:
        sr_star = 0.0
    else:
        norm = NormalDist()
        q1 = norm.inv_cdf(1.0 - 1.0 / effective_trials)
        q2 = norm.inv_cdf(1.0 - 1.0 / (effective_trials * np.e))
        euler_gamma = 0.5772156649
        sr_star = (1.0 - euler_gamma) * q1 + euler_gamma * q2
    return float(NormalDist().cdf((sr - sr_star) / max(sr_std, 1e-12)))


def probabilistic_sharpe_ratio(
    *,
    observed_sharpe: float,
    benchmark_sharpe: float,
    n_returns: int,
    skew: float = 0.0,
    kurtosis: float = 3.0,
) -> float:
    """Return Probabilistic Sharpe Ratio: P(SR > benchmark_sharpe)."""
    if n_returns <= 2:
        return 0.0
    sr = float(observed_sharpe)
    sr_ref = float(benchmark_sharpe)
    variance_term = 1.0 - skew * sr + ((kurtosis - 1.0) / 4.0) * (sr**2)
    sr_std = np.sqrt(max(variance_term, 1e-12) / max(1, n_returns - 1))
    z_score = (sr - sr_ref) / max(sr_std, 1e-12)
    return float(NormalDist().cdf(z_score))


def benjamini_hochberg_adjusted_pvalues(p_values: list[float]) -> list[float]:
    """Compute Benjamini-Hochberg adjusted p-values (q-values)."""
    m = len(p_values)
    if m == 0:
        return []

    clipped = np.clip(np.asarray(p_values, dtype=float), 0.0, 1.0)
    order = np.argsort(clipped)
    ranked = clipped[order]

    adjusted = np.empty(m, dtype=float)
    running = 1.0
    for idx in range(m - 1, -1, -1):
        rank = idx + 1
        candidate = (ranked[idx] * m) / rank
        running = min(running, float(candidate))
        adjusted[idx] = running

    out = np.empty(m, dtype=float)
    out[order] = np.clip(adjusted, 0.0, 1.0)
    return out.tolist()


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

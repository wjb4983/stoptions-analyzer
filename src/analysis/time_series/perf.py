"""Profiling helpers for time-series momentum kernels."""

from __future__ import annotations

import cProfile
import pstats
from dataclasses import dataclass
from io import StringIO

import numpy as np

from .momentum import MomentumHyperparameters, TimeSeriesMomentumSettings, build_time_series_momentum_arrays


@dataclass(frozen=True)
class ProfileSummary:
    mode: str
    output: str


def profile_signal_rolling_hotspots(n_periods: int = 250_000) -> list[ProfileSummary]:
    rng = np.random.default_rng(7)
    prices = 100.0 * np.cumprod(1.0 + rng.normal(0.0001, 0.01, size=n_periods))
    settings = TimeSeriesMomentumSettings(
        hyperparameters=MomentumHyperparameters(
            lookback_days=63,
            skip_days=1,
            vol_window_days=20,
            target_volatility=0.2,
            max_leverage=2.0,
        )
    )

    outputs: list[ProfileSummary] = []
    for mode in ("reference", "optimized"):
        pr = cProfile.Profile()
        pr.enable()
        build_time_series_momentum_arrays(prices.tolist(), settings, mode=mode)
        pr.disable()

        buf = StringIO()
        pstats.Stats(pr, stream=buf).sort_stats("cumtime").print_stats(15)
        outputs.append(ProfileSummary(mode=mode, output=buf.getvalue()))
    return outputs

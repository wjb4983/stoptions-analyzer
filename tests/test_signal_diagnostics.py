from __future__ import annotations

import numpy as np

from src.analysis.cross_sectional.momentum import MomentumSettings, compute_cross_sectional_momentum
from src.analysis.diagnostics import compute_signal_diagnostics, validate_signal_diagnostics
from src.analysis.time_series.momentum import TimeSeriesMomentumSettings, compute_time_series_momentum
from src.backtesting.strategies.alpha_model import build_signal_diagnostics


def _prices() -> dict[str, list[dict[str, float]]]:
    base = np.linspace(100, 120, 80)
    return {
        "AAA": [{"close": float(v), "volume": 1_000_000.0} for v in base],
        "BBB": [{"close": float(v), "volume": 800_000.0} for v in base * 0.98],
        "CCC": [{"close": float(v), "volume": 500_000.0} for v in base[::-1] * 1.02],
        "DDD": [{"close": float(v), "volume": 600_000.0} for v in base * 1.01],
    }


def test_compute_signal_diagnostics_sections_present() -> None:
    prices = _prices()
    diagnostics = compute_signal_diagnostics(
        scores={"AAA": 0.4, "BBB": 0.2, "CCC": -0.1, "DDD": -0.3},
        weights={"AAA": 0.25, "BBB": 0.25, "CCC": -0.25, "DDD": -0.25},
        prices_by_ticker=prices,
    )

    assert validate_signal_diagnostics(diagnostics) is True
    assert "information_coefficient" in diagnostics
    assert "rank_stability" in diagnostics
    assert "exposure" in diagnostics


def test_time_series_and_cross_sectional_attach_diagnostics() -> None:
    prices = _prices()
    ts = compute_time_series_momentum(prices, None, TimeSeriesMomentumSettings(lookback_days=20, skip_days=2))
    cs = compute_cross_sectional_momentum(prices, None, MomentumSettings(lookback_days=20, skip_days=2))

    assert "diagnostics" in ts.metadata
    assert "diagnostics" in cs.metadata
    assert validate_signal_diagnostics(ts.metadata["diagnostics"]) is True
    assert validate_signal_diagnostics(cs.metadata["diagnostics"]) is True


def test_alpha_model_build_signal_diagnostics_ready_flag() -> None:
    prices = _prices()
    out = build_signal_diagnostics(
        signal_by_ticker={"AAA": 0.8, "BBB": 0.3, "CCC": -0.2, "DDD": -0.6},
        weights_by_ticker={"AAA": 0.25, "BBB": 0.25, "CCC": -0.25, "DDD": -0.25},
        prices_by_ticker=prices,
    )

    assert out.diagnostics_ready is True
    assert out.diagnostics["exposure"]["gross_exposure"] > 0

import numpy as np

from backtesting.regimes import (
    RegimeFeatureConfig,
    apply_regime_risk_overlays,
    compute_regime_labels,
)


def test_compute_regime_labels_is_deterministic() -> None:
    prices = np.array(
        [
            [100.0, 200.0],
            [101.0, 198.0],
            [102.0, 197.0],
            [103.0, 196.0],
            [102.0, 197.0],
            [101.0, 198.0],
        ],
        dtype=float,
    )
    cfg = RegimeFeatureConfig(
        trend_fast_window=2,
        trend_slow_window=3,
        vol_window=2,
        liquidity_window=2,
        macro_window=2,
    )

    first = compute_regime_labels(prices, config=cfg)
    second = compute_regime_labels(prices, config=cfg)

    for key in ("trend", "volatility", "liquidity", "macro", "labels"):
        assert np.array_equal(first[key], second[key])


def test_regime_risk_overlay_caps_exposure_by_regime() -> None:
    weights = np.array(
        [
            [0.6, -0.6],
            [0.5, -0.5],
            [0.4, -0.4],
        ],
        dtype=float,
    )
    regime_labels = np.array(
        [
            "trend_up|vol_high|liq_thin|macro_risk_off",
            "trend_up|vol_low|liq_normal|macro_risk_on",
            "trend_up|vol_high|liq_thin|macro_risk_off",
        ],
        dtype=object,
    )
    risk_map = {
        "default": {"risk_multiplier": 1.0, "max_gross_exposure": 1.0},
        "trend_up|vol_high|liq_thin|macro_risk_off": {
            "risk_multiplier": 0.5,
            "max_gross_exposure": 0.4,
        },
    }

    adjusted, diagnostics = apply_regime_risk_overlays(
        weights=weights,
        regime_labels=regime_labels,
        risk_map=risk_map,
    )

    gross = np.sum(np.abs(adjusted), axis=1)
    assert gross[0] <= 0.4000001
    assert gross[2] <= 0.4000001
    assert abs(gross[1] - 1.0) < 1e-9

    multipliers = diagnostics["regime_risk_multiplier"]
    assert np.array_equal(multipliers, np.array([0.5, 1.0, 0.5], dtype=float))

from __future__ import annotations

import numpy as np

from src.analysis.options import compute_options_feature_pipeline


def _build_raw_inputs(n_time: int = 30, n_assets: int = 4) -> dict[str, np.ndarray]:
    t = np.linspace(0.0, 1.0, n_time)[:, None]
    a = np.linspace(-0.5, 0.5, n_assets)[None, :]

    iv_atm = 0.2 + 0.03 * t + 0.01 * a
    iv_25d_put = iv_atm + 0.03 + 0.005 * a
    iv_25d_call = iv_atm - 0.01 + 0.004 * a

    return {
        "iv_atm": iv_atm,
        "iv_25d_put": iv_25d_put,
        "iv_25d_call": iv_25d_call,
        "iv_10d_put": iv_25d_put + 0.01,
        "iv_10d_call": iv_25d_call + 0.005,
        "iv_1m": iv_atm,
        "iv_3m": iv_atm - 0.01,
        "iv_6m": iv_atm - 0.015,
        "put_volume": 80 + 10 * t + 4 * a,
        "call_volume": 70 + 8 * t - 3 * a,
        "open_interest": 1000 + 15 * np.arange(n_time)[:, None] + 10 * a,
        "net_gamma_notional": 3e6 + 2e5 * np.sin(2 * np.pi * t) + 5e4 * a,
        "underlying_market_cap": 1e10 + 5e8 * a,
        "spot_return": 0.01 * np.sin(4 * np.pi * t) + 0.002 * a,
    }


def test_feature_pipeline_outputs_expected_features_and_shapes() -> None:
    raw = _build_raw_inputs()
    result = compute_options_feature_pipeline(raw, rolling_window=10)

    expected = {
        "skew",
        "convexity",
        "term_structure_curvature",
        "local_surface_distortion",
        "put_call_flow_imbalance",
        "oi_changes",
        "gamma_exposure_proxy",
        "dealer_positioning_proxy",
        "unusual_volume_signature",
    }

    for name in expected:
        assert name in result.features
        assert f"{name}_z" in result.features
        assert f"{name}_rank" in result.features
        assert result.features[name].shape == raw["iv_atm"].shape


def test_pipeline_sanity_checks_follow_expected_market_behavior() -> None:
    raw = _build_raw_inputs()
    result = compute_options_feature_pipeline(raw, rolling_window=8)

    assert result.sanity_checks["flow_bounded"] is True
    assert result.sanity_checks["unusual_volume_positive"] is True
    assert result.sanity_checks["skew_mostly_put_rich"] is True

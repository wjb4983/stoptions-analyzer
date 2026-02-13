from __future__ import annotations

import numpy as np

from src.backtesting import perf
from src.backtesting.regimes import _rolling_mean as regime_rolling_mean
from src.backtesting.regimes import _rolling_std as regime_rolling_std
from src.backtesting.vectorized import _rolling_mean as vec_rolling_mean
from src.backtesting.vectorized import _rolling_volatility


def _ref_rolling_mean_1d(values: np.ndarray, window: int) -> np.ndarray:
    out = np.zeros_like(values, dtype=float)
    for idx in range(values.size):
        start = max(0, idx - window + 1)
        out[idx] = float(np.mean(values[start : idx + 1]))
    return out


def _ref_rolling_std_1d(values: np.ndarray, window: int, ddof: int = 1) -> np.ndarray:
    out = np.zeros_like(values, dtype=float)
    for idx in range(values.size):
        start = max(0, idx - window + 1)
        segment = values[start : idx + 1]
        out[idx] = float(np.std(segment, ddof=ddof)) if segment.size > ddof else 0.0
    return out


def _ref_rolling_mean_2d(values: np.ndarray, window: int) -> np.ndarray:
    out = np.zeros_like(values, dtype=float)
    for col in range(values.shape[1]):
        for idx in range(values.shape[0]):
            start = max(0, idx - window + 1)
            out[idx, col] = float(np.mean(values[start : idx + 1, col]))
    return out


def _ref_rolling_volatility(prices: np.ndarray, window: int) -> np.ndarray:
    rets = np.zeros_like(prices, dtype=float)
    rets[1:] = prices[1:] / np.where(prices[:-1] == 0.0, 1.0, prices[:-1]) - 1.0
    out = np.zeros_like(prices, dtype=float)
    for col in range(prices.shape[1]):
        for idx in range(prices.shape[0]):
            start = max(0, idx - window + 1)
            segment = rets[start : idx + 1, col]
            out[idx, col] = float(np.std(segment, ddof=1)) if segment.size > 1 else 0.0
    return out


def test_regime_rolling_functions_match_reference() -> None:
    rng = np.random.default_rng(7)
    values = rng.normal(size=257)

    for window in (1, 2, 5, 32, 400):
        got_mean = regime_rolling_mean(values, window)
        exp_mean = _ref_rolling_mean_1d(values, window)
        assert np.allclose(got_mean, exp_mean, rtol=1e-12, atol=1e-12)

        got_std = regime_rolling_std(values, window)
        exp_std = _ref_rolling_std_1d(values, window, ddof=1)
        assert np.allclose(got_std, exp_std, rtol=1e-12, atol=1e-12)


def test_vectorized_rolling_functions_match_reference() -> None:
    rng = np.random.default_rng(11)
    values = rng.normal(size=(251, 6))
    prices = 100.0 * np.cumprod(1.0 + 0.01 * rng.normal(size=(251, 6)), axis=0)

    for window in (1, 2, 5, 20, 1000):
        got_mean = vec_rolling_mean(values, window)
        exp_mean = _ref_rolling_mean_2d(values, window)
        assert np.allclose(got_mean, exp_mean, rtol=1e-12, atol=1e-12)

        got_vol = _rolling_volatility(prices, window)
        exp_vol = _ref_rolling_volatility(prices, window)
        assert np.allclose(got_vol, exp_vol, rtol=1e-12, atol=1e-12)


def test_profile_regression_thresholds() -> None:
    summaries = [
        perf.ProfileSummary(mode="reference", output="", elapsed_seconds=1.0),
        perf.ProfileSummary(mode="optimized", output="", elapsed_seconds=1.1),
    ]
    report = perf.check_profile_regression(summaries)
    assert report["pass"] is True
    assert float(report["optimized_to_reference_ratio"]) <= float(report["optimized_to_reference_max_ratio"])

    fail_report = perf.check_profile_regression(
        summaries,
        thresholds={"optimized_to_reference_max_ratio": 1.05},
    )
    assert fail_report["pass"] is False


def test_serialization_boundary_benchmark_shapes() -> None:
    report = perf.benchmark_serialization_boundaries(n_periods=250, n_assets=4)
    assert report["n_periods"] == 250
    assert report["n_assets"] == 4
    results = report["results"]
    assert "npz" in results and "json" in results and "parquet" in results
    assert float(results["npz"]["write_seconds"]) >= 0.0
    assert float(results["npz"]["read_seconds"]) >= 0.0
    assert float(results["json"]["size_bytes"]) > 0.0

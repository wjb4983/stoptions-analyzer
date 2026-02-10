from __future__ import annotations

import numpy as np

from src.analysis.time_series.momentum import (
    MomentumHyperparameters,
    TimeSeriesMomentumSettings,
    build_time_series_momentum_arrays,
    compute_time_series_momentum,
)


def test_momentum_arrays_respect_skip_and_no_lookahead_shift() -> None:
    closes = [100.0, 102.0, 104.0, 108.0, 110.0]
    settings = TimeSeriesMomentumSettings(
        hyperparameters=MomentumHyperparameters(lookback_days=2, skip_days=1)
    )

    arrays = build_time_series_momentum_arrays(closes, settings)

    expected_raw = np.array([np.nan, np.nan, np.nan, 0.04, 0.05882353])
    assert np.allclose(arrays.raw_score[3:], expected_raw[3:], atol=1e-8)
    assert np.all(arrays.position_signal[:3] == 0)
    assert np.array_equal(arrays.position_signal[3:], np.array([1.0, 1.0]))
    # t close signal becomes tradable at t+1 open.
    assert arrays.tradable_position[3] == 0.0
    assert arrays.tradable_position[4] == arrays.scaled_weight[3]


def test_momentum_arrays_apply_vol_target_and_max_leverage() -> None:
    closes = [100.0, 105.0, 110.0, 100.0, 102.0, 104.0]
    settings = TimeSeriesMomentumSettings(
        hyperparameters=MomentumHyperparameters(
            lookback_days=1,
            skip_days=0,
            vol_window_days=3,
            target_volatility=0.2,
            max_leverage=0.5,
        )
    )

    arrays = build_time_series_momentum_arrays(closes, settings)

    assert np.max(np.abs(arrays.scaled_weight)) <= 0.5 + 1e-12
    assert arrays.scaled_weight[-1] > 0


def test_compute_time_series_momentum_emits_standardized_metrics() -> None:
    settings = TimeSeriesMomentumSettings(
        hyperparameters=MomentumHyperparameters(lookback_days=2, skip_days=1),
        top_quantile=0.5,
        bottom_quantile=0.5,
    )
    prices_by_ticker = {
        "AAA": [100.0, 102.0, 104.0, 108.0, 110.0],
        "BBB": [110.0, 108.0, 106.0, 104.0, 102.0],
    }

    result = compute_time_series_momentum(prices_by_ticker, None, settings)

    assert set(result.metrics["AAA"].keys()) >= {
        "base",
        "latest_position_signal",
        "latest_confidence",
        "latest_scaled_weight",
        "latest_tradable_position",
    }
    assert result.metrics["AAA"]["latest_position_signal"] == 1.0
    assert result.metrics["BBB"]["latest_position_signal"] == -1.0
    assert result.metadata["vol_window_days"] == 20
    assert result.metadata["max_leverage"] == 1.0

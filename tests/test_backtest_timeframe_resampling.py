from __future__ import annotations

import numpy as np

from src.backtesting import cache_runner
from src.data_access.engine_loader import EngineArrayBundle, EngineArrayMetadata


def _sample_bundle(rows: int = 12) -> EngineArrayBundle:
    date_index = np.arange(rows, dtype=np.int64)
    open_prices = np.arange(rows * 2, dtype=float).reshape(rows, 2)
    close_prices = open_prices + 0.5
    missing_mask = np.zeros((rows, 2), dtype=bool)
    metadata = EngineArrayMetadata(
        symbol_to_column={"AAA": 0, "BBB": 1},
        date_index=date_index,
        missingness_ratio=0.0,
        missingness_by_symbol={"AAA": 0.0, "BBB": 0.0},
        coverage_by_symbol={"AAA": 1.0, "BBB": 1.0},
        tradable_ratio_by_symbol={"AAA": 1.0, "BBB": 1.0},
        excluded_symbols={},
        audit_summary_by_symbol={},
    )
    return EngineArrayBundle(
        date_index=date_index,
        open_prices=open_prices,
        close_prices=close_prices,
        missing_mask=missing_mask,
        metadata=metadata,
    )


def test_resample_from_1m_to_5m_uses_index_math() -> None:
    bundle = _sample_bundle(rows=12)

    resampled = cache_runner._resample_engine_bundle_from_1m(bundle, timeframe="5m")

    assert resampled.date_index.tolist() == [4, 9]
    assert resampled.close_prices.shape == (2, 2)
    assert np.allclose(resampled.close_prices[:, 0], [8.5, 18.5])


def test_resample_1m_is_noop() -> None:
    bundle = _sample_bundle(rows=7)

    resampled = cache_runner._resample_engine_bundle_from_1m(bundle, timeframe="1m")

    assert np.array_equal(resampled.date_index, bundle.date_index)
    assert np.array_equal(resampled.close_prices, bundle.close_prices)

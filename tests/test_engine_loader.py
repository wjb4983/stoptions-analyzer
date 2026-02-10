from __future__ import annotations

from datetime import datetime, timezone

import numpy as np
import pytest

from src.data_access.engine_loader import load_canonical_price_arrays


def _write_npz(cache_root, symbol: str, year: int, **arrays) -> None:
    safe = symbol.upper()
    folder = cache_root / safe / "1m"
    folder.mkdir(parents=True, exist_ok=True)
    np.savez_compressed(folder / f"{safe}_1m_{year}.npz", **arrays)


def _ms(year: int, month: int, day: int, hour: int, minute: int) -> int:
    return int(datetime(year, month, day, hour, minute, tzinfo=timezone.utc).timestamp() * 1000)


def test_loader_aligns_float64_and_metadata(tmp_path) -> None:
    ts1 = np.array([_ms(2024, 1, 2, 14, 30), _ms(2024, 1, 2, 14, 31)], dtype=np.int64)
    ts2 = np.array([_ms(2024, 1, 2, 14, 31), _ms(2024, 1, 2, 14, 32)], dtype=np.int64)

    _write_npz(tmp_path, "AAPL", 2024, t=ts1, o=np.array([100, 101]), c=np.array([100.5, 101.5]))
    _write_npz(tmp_path, "MSFT", 2024, t=ts2, o=np.array([200, 201]), c=np.array([200.5, 201.5]))

    bundle = load_canonical_price_arrays(
        symbols=["AAPL", "MSFT"],
        start="2024-01-02T14:30:00+00:00",
        end="2024-01-02T14:32:00+00:00",
        cache_root=tmp_path,
        lookback_window=1,
    )

    assert bundle.open_prices.shape == (3, 2)
    assert bundle.close_prices.shape == (3, 2)
    assert bundle.missing_mask.shape == (3, 2)
    assert bundle.open_prices.dtype == np.float64
    assert bundle.close_prices.dtype == np.float64
    assert bundle.metadata.symbol_to_column == {"AAPL": 0, "MSFT": 1}
    assert np.isnan(bundle.open_prices[0, 1])
    assert bundle.missing_mask[0, 1]
    assert bundle.metadata.missingness_ratio == pytest.approx(2 / 6)


def test_loader_rejects_duplicate_timestamps(tmp_path) -> None:
    ts = np.array([_ms(2024, 1, 2, 14, 30), _ms(2024, 1, 2, 14, 30)], dtype=np.int64)
    _write_npz(tmp_path, "AAPL", 2024, t=ts, o=np.array([100, 101]), c=np.array([100.5, 101.5]))

    with pytest.raises(ValueError, match="Duplicate timestamps"):
        load_canonical_price_arrays(
            symbols=["AAPL"],
            start="2024-01-02T14:30:00+00:00",
            end="2024-01-02T14:31:00+00:00",
            cache_root=tmp_path,
        )


def test_loader_rejects_non_monotonic_time(tmp_path) -> None:
    ts = np.array([_ms(2024, 1, 2, 14, 31), _ms(2024, 1, 2, 14, 30)], dtype=np.int64)
    _write_npz(tmp_path, "AAPL", 2024, t=ts, o=np.array([100, 101]), c=np.array([100.5, 101.5]))

    with pytest.raises(ValueError, match="Non-monotonic"):
        load_canonical_price_arrays(
            symbols=["AAPL"],
            start="2024-01-02T14:30:00+00:00",
            end="2024-01-02T14:31:00+00:00",
            cache_root=tmp_path,
        )


def test_loader_split_adjustment_validation(tmp_path) -> None:
    ts = np.array([_ms(2024, 1, 2, 14, 30), _ms(2024, 1, 2, 14, 31)], dtype=np.int64)
    _write_npz(
        tmp_path,
        "AAPL",
        2024,
        t=ts,
        o=np.array([100, 101]),
        c=np.array([100.5, 101.5]),
        split_factor=np.array([1.0, 0.0]),
    )

    with pytest.raises(ValueError, match="split factors"):
        load_canonical_price_arrays(
            symbols=["AAPL"],
            start="2024-01-02T14:30:00+00:00",
            end="2024-01-02T14:31:00+00:00",
            cache_root=tmp_path,
        )


def test_loader_minimum_lookback_validation(tmp_path) -> None:
    ts = np.array([_ms(2024, 1, 2, 14, 30)], dtype=np.int64)
    _write_npz(tmp_path, "AAPL", 2024, t=ts, o=np.array([100]), c=np.array([100.5]))

    with pytest.raises(ValueError, match="Insufficient history"):
        load_canonical_price_arrays(
            symbols=["AAPL"],
            start="2024-01-02T14:30:00+00:00",
            end="2024-01-02T14:31:00+00:00",
            cache_root=tmp_path,
            lookback_window=2,
        )

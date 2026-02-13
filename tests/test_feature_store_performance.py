from __future__ import annotations

from datetime import datetime, timedelta

import pytest

from data.feature_store import FeatureMetadata, FeatureStore


def _build_store() -> FeatureStore:
    store = FeatureStore(point_in_time_cache_size=2)
    store.register_snapshot(
        feature_name="rv",
        metadata=FeatureMetadata(source="vendor", lag=timedelta(minutes=30), refresh_cadence=timedelta(minutes=5)),
        records=[
            {"entity": "AAPL", "timestamp": datetime(2024, 1, 1, 9, 0), "value": 0.2},
            {"entity": "AAPL", "timestamp": datetime(2024, 1, 1, 9, 30), "value": 0.25},
            {"entity": "MSFT", "timestamp": datetime(2024, 1, 1, 9, 0), "value": 0.18},
        ],
    )
    return store


def test_point_in_time_join_cache_populates_and_reuses() -> None:
    store = _build_store()
    obs = [{"entity": "AAPL", "timestamp": datetime(2024, 1, 1, 10, 5)}]

    first = store.point_in_time_join(observations=obs)
    info_after_first = store.point_in_time_cache_info()
    second = store.point_in_time_join(observations=obs)
    info_after_second = store.point_in_time_cache_info()

    assert first == second
    assert info_after_first["entries"] == 1
    assert info_after_second["entries"] == 1


def test_register_snapshot_invalidates_point_in_time_cache() -> None:
    store = _build_store()
    obs = [{"entity": "AAPL", "timestamp": datetime(2024, 1, 1, 10, 5)}]
    store.point_in_time_join(observations=obs)
    assert store.point_in_time_cache_info()["entries"] == 1

    store.register_snapshot(
        feature_name="rv",
        metadata=FeatureMetadata(source="vendor", lag=timedelta(minutes=30), refresh_cadence=timedelta(minutes=5)),
        records=[{"entity": "AAPL", "timestamp": datetime(2024, 1, 1, 9, 45), "value": 0.28}],
    )

    assert store.point_in_time_cache_info()["entries"] == 0


def test_export_snapshot_columnar_requires_pyarrow_when_unavailable(tmp_path, monkeypatch) -> None:
    import data.feature_store.store as feature_store_module

    store = _build_store()
    monkeypatch.setattr(feature_store_module, "pa", None)

    with pytest.raises(RuntimeError):
        store.export_snapshot_columnar(feature_name="rv", output_path=tmp_path / "rv.parquet")

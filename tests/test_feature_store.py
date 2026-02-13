from __future__ import annotations

from datetime import date, datetime, timedelta

import pytest

from data.feature_store import FeatureLeakageError, FeatureMetadata, FeatureStore, generate_daily_feature_report


def test_point_in_time_join_prevents_future_leakage() -> None:
    store = FeatureStore()
    store.register_snapshot(
        feature_name="iv_rank",
        metadata=FeatureMetadata(source="vendor_a", lag=timedelta(hours=1), refresh_cadence=timedelta(hours=1)),
        records=[
            {"entity": "AAPL", "timestamp": datetime(2024, 1, 1, 9, 0), "value": 0.15},
            {"entity": "AAPL", "timestamp": datetime(2024, 1, 1, 11, 0), "value": 0.45},
        ],
    )

    joined = store.point_in_time_join(
        observations=[
            {"entity": "AAPL", "timestamp": datetime(2024, 1, 1, 10, 30)},
        ]
    )

    assert joined[0]["iv_rank"] == 0.15
    assert joined[0]["iv_rank__asof"] == datetime(2024, 1, 1, 9, 0)


def test_snapshot_versioning_uses_requested_version() -> None:
    store = FeatureStore()
    metadata = FeatureMetadata(source="vendor_a", lag=timedelta(minutes=30), refresh_cadence=timedelta(minutes=30))
    store.register_snapshot(
        feature_name="realized_vol",
        metadata=metadata,
        records=[{"entity": "SPY", "timestamp": datetime(2024, 1, 2, 10, 0), "value": 0.20}],
    )
    store.register_snapshot(
        feature_name="realized_vol",
        metadata=metadata,
        records=[{"entity": "SPY", "timestamp": datetime(2024, 1, 2, 10, 0), "value": 0.25}],
    )

    latest = store.point_in_time_join(observations=[{"entity": "SPY", "timestamp": datetime(2024, 1, 2, 11, 0)}])
    old = store.point_in_time_join(
        observations=[{"entity": "SPY", "timestamp": datetime(2024, 1, 2, 11, 0)}],
        feature_versions={"realized_vol": 1},
    )

    assert latest[0]["realized_vol"] == 0.25
    assert latest[0]["realized_vol__version"] == 2
    assert old[0]["realized_vol"] == 0.20
    assert old[0]["realized_vol__version"] == 1


def test_negative_lag_is_rejected_as_leakage_attempt() -> None:
    store = FeatureStore()

    with pytest.raises(FeatureLeakageError):
        store.register_snapshot(
            feature_name="term_structure",
            metadata=FeatureMetadata(source="vendor_b", lag=timedelta(minutes=-5), refresh_cadence=timedelta(minutes=5)),
            records=[{"entity": "QQQ", "timestamp": datetime(2024, 1, 3, 9, 35), "value": 1.2}],
        )


def test_daily_report_tracks_freshness_and_null_rate() -> None:
    joined_rows = [
        {
            "entity": "AAPL",
            "timestamp": datetime(2024, 1, 4, 10, 0),
            "iv_rank": 0.2,
            "iv_rank__asof": datetime(2024, 1, 4, 9, 0),
        },
        {
            "entity": "AAPL",
            "timestamp": datetime(2024, 1, 4, 11, 0),
            "iv_rank": None,
            "iv_rank__asof": None,
        },
    ]

    report = generate_daily_feature_report(
        report_date=date(2024, 1, 4),
        joined_rows=joined_rows,
        feature_names=["iv_rank"],
    )

    metrics = report["feature_health"]["iv_rank"]
    assert metrics["null_rate"] == 0.5
    assert metrics["freshness"]["mean_seconds"] == 3600.0
    assert metrics["non_null_count"] == 1

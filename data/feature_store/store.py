from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Any


class FeatureLeakageError(ValueError):
    """Raised when a point-in-time request would leak future information."""


@dataclass(frozen=True)
class FeatureMetadata:
    source: str
    lag: timedelta
    refresh_cadence: timedelta


@dataclass(frozen=True)
class FeatureSnapshot:
    feature_name: str
    version: int
    created_at: datetime
    metadata: FeatureMetadata
    entity_key: str
    timestamp_key: str
    value_key: str
    records: tuple[dict[str, Any], ...]


class FeatureStore:
    def __init__(self) -> None:
        self._snapshots: dict[str, list[FeatureSnapshot]] = {}

    def register_snapshot(
        self,
        *,
        feature_name: str,
        records: list[dict[str, Any]],
        metadata: FeatureMetadata,
        entity_key: str = "entity",
        timestamp_key: str = "timestamp",
        value_key: str = "value",
        created_at: datetime | None = None,
    ) -> FeatureSnapshot:
        versions = self._snapshots.setdefault(feature_name, [])
        if metadata.lag.total_seconds() < 0:
            raise FeatureLeakageError(f"Feature '{feature_name}' has negative lag; this permits future leakage")

        snapshot = FeatureSnapshot(
            feature_name=feature_name,
            version=len(versions) + 1,
            created_at=created_at or datetime.utcnow(),
            metadata=metadata,
            entity_key=entity_key,
            timestamp_key=timestamp_key,
            value_key=value_key,
            records=tuple(_normalize_record(r, entity_key, timestamp_key, value_key) for r in records),
        )
        versions.append(snapshot)
        return snapshot

    def get_snapshot(self, feature_name: str, version: int | None = None) -> FeatureSnapshot:
        if feature_name not in self._snapshots or not self._snapshots[feature_name]:
            raise KeyError(f"Unknown feature '{feature_name}'")
        versions = self._snapshots[feature_name]
        if version is None:
            return versions[-1]
        for snapshot in versions:
            if snapshot.version == version:
                return snapshot
        raise KeyError(f"Unknown version {version} for feature '{feature_name}'")

    def point_in_time_join(
        self,
        *,
        observations: list[dict[str, Any]],
        feature_versions: dict[str, int | None] | None = None,
        observation_entity_key: str = "entity",
        observation_timestamp_key: str = "timestamp",
        strict: bool = True,
    ) -> list[dict[str, Any]]:
        feature_versions = feature_versions or {}
        enriched: list[dict[str, Any]] = []

        snapshots = {
            feature_name: self.get_snapshot(feature_name, version)
            for feature_name, version in feature_versions.items()
        }
        if not snapshots:
            for feature_name, versions in self._snapshots.items():
                if versions:
                    snapshots[feature_name] = versions[-1]

        for obs in observations:
            row = dict(obs)
            obs_ts = _to_datetime(obs[observation_timestamp_key])
            entity = obs[observation_entity_key]

            for feature_name, snapshot in snapshots.items():
                cutoff = obs_ts - snapshot.metadata.lag
                candidate = _latest_record_for_entity(
                    records=snapshot.records,
                    entity_key=snapshot.entity_key,
                    timestamp_key=snapshot.timestamp_key,
                    value_key=snapshot.value_key,
                    entity=entity,
                    max_timestamp=cutoff,
                )

                if strict and candidate is not None and candidate["timestamp"] > cutoff:
                    raise FeatureLeakageError(
                        f"Leakage detected for '{feature_name}' entity={entity}: "
                        f"candidate timestamp {candidate['timestamp'].isoformat()} exceeds cutoff {cutoff.isoformat()}"
                    )

                row[feature_name] = None if candidate is None else candidate["value"]
                row[f"{feature_name}__asof"] = None if candidate is None else candidate["timestamp"]
                row[f"{feature_name}__version"] = snapshot.version
                row[f"{feature_name}__source"] = snapshot.metadata.source
                row[f"{feature_name}__lag_seconds"] = int(snapshot.metadata.lag.total_seconds())
                row[f"{feature_name}__refresh_cadence_seconds"] = int(snapshot.metadata.refresh_cadence.total_seconds())

            enriched.append(row)

        return enriched


def generate_daily_feature_report(
    *,
    report_date: date,
    joined_rows: list[dict[str, Any]],
    feature_names: list[str],
    observation_timestamp_key: str = "timestamp",
) -> dict[str, Any]:
    report: dict[str, Any] = {
        "report_date": report_date.isoformat(),
        "feature_health": {},
    }

    for feature_name in feature_names:
        values = [row.get(feature_name) for row in joined_rows]
        asofs = [row.get(f"{feature_name}__asof") for row in joined_rows]
        ts_values = [_to_datetime(row[observation_timestamp_key]) for row in joined_rows if observation_timestamp_key in row]

        non_null = sum(1 for value in values if value is not None)
        null_rate = 1.0 if not values else 1.0 - (non_null / len(values))

        freshness_seconds: list[float] = []
        for obs_ts, asof_ts in zip(ts_values, asofs, strict=False):
            if asof_ts is not None:
                freshness_seconds.append((obs_ts - _to_datetime(asof_ts)).total_seconds())

        freshness = {
            "mean_seconds": 0.0 if not freshness_seconds else float(sum(freshness_seconds) / len(freshness_seconds)),
            "max_seconds": 0.0 if not freshness_seconds else float(max(freshness_seconds)),
            "min_seconds": 0.0 if not freshness_seconds else float(min(freshness_seconds)),
        }

        report["feature_health"][feature_name] = {
            "null_rate": float(null_rate),
            "freshness": freshness,
            "sample_count": len(values),
            "non_null_count": non_null,
        }

    return report


def _normalize_record(record: dict[str, Any], entity_key: str, timestamp_key: str, value_key: str) -> dict[str, Any]:
    if entity_key not in record or timestamp_key not in record or value_key not in record:
        raise ValueError(f"Record is missing one of required keys: {entity_key}, {timestamp_key}, {value_key}")
    return {
        entity_key: record[entity_key],
        timestamp_key: _to_datetime(record[timestamp_key]),
        value_key: record[value_key],
    }


def _to_datetime(value: Any) -> datetime:
    if isinstance(value, datetime):
        return value
    if isinstance(value, str):
        return datetime.fromisoformat(value)
    raise TypeError(f"Expected datetime or ISO string, got {type(value)!r}")


def _latest_record_for_entity(
    *,
    records: tuple[dict[str, Any], ...],
    entity_key: str,
    timestamp_key: str,
    value_key: str,
    entity: Any,
    max_timestamp: datetime,
) -> dict[str, Any] | None:
    candidates = [
        r for r in records if r[entity_key] == entity and r[timestamp_key] <= max_timestamp
    ]
    if not candidates:
        return None
    chosen = max(candidates, key=lambda r: r[timestamp_key])
    return {"timestamp": chosen[timestamp_key], "value": chosen[value_key]}

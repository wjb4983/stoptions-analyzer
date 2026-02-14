from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any
import hashlib
import importlib.util
import json

import numpy as np


_pyarrow_spec = importlib.util.find_spec("pyarrow")
if _pyarrow_spec:
    import pyarrow as pa
    import pyarrow.feather as feather
    import pyarrow.parquet as pq
else:
    pa = None
    feather = None
    pq = None


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
    version_keys: "FeatureVersionKeys"


@dataclass(frozen=True)
class FeatureVersionKeys:
    data_source_hash: str
    transformation_graph_hash: str
    training_window_metadata: dict[str, Any]
    feature_set_id: str


@dataclass(frozen=True)
class _PreparedSnapshot:
    entities: dict[Any, tuple[np.ndarray, np.ndarray]]


class FeatureStore:
    def __init__(self, *, point_in_time_cache_size: int = 16) -> None:
        self._snapshots: dict[str, list[FeatureSnapshot]] = {}
        self._prepared: dict[tuple[str, int], _PreparedSnapshot] = {}
        self._point_in_time_cache: dict[tuple[Any, ...], list[dict[str, Any]]] = {}
        self._point_in_time_cache_size = max(1, int(point_in_time_cache_size))

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
        data_source_hash: str | None = None,
        transformation_graph_hash: str | None = None,
        training_window_metadata: dict[str, Any] | None = None,
    ) -> FeatureSnapshot:
        versions = self._snapshots.setdefault(feature_name, [])
        if metadata.lag.total_seconds() < 0:
            raise FeatureLeakageError(f"Feature '{feature_name}' has negative lag; this permits future leakage")

        normalized_training_window = _normalize_training_window_metadata(training_window_metadata)
        version_keys = _build_feature_version_keys(
            feature_name=feature_name,
            version=len(versions) + 1,
            metadata=metadata,
            data_source_hash=data_source_hash,
            transformation_graph_hash=transformation_graph_hash,
            training_window_metadata=normalized_training_window,
        )

        snapshot = FeatureSnapshot(
            feature_name=feature_name,
            version=len(versions) + 1,
            created_at=created_at or datetime.utcnow(),
            metadata=metadata,
            entity_key=entity_key,
            timestamp_key=timestamp_key,
            value_key=value_key,
            records=tuple(_normalize_record(r, entity_key, timestamp_key, value_key) for r in records),
            version_keys=version_keys,
        )
        versions.append(snapshot)
        self._prepared[(feature_name, snapshot.version)] = _prepare_snapshot(snapshot)
        self._point_in_time_cache.clear()
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

        snapshots = {
            feature_name: self.get_snapshot(feature_name, version)
            for feature_name, version in feature_versions.items()
        }
        if not snapshots:
            for feature_name, versions in self._snapshots.items():
                if versions:
                    snapshots[feature_name] = versions[-1]

        cache_key = _build_point_in_time_cache_key(
            observations=observations,
            snapshots=snapshots,
            observation_entity_key=observation_entity_key,
            observation_timestamp_key=observation_timestamp_key,
            strict=strict,
        )
        cached = self._point_in_time_cache.get(cache_key)
        if cached is not None:
            return [dict(row) for row in cached]

        enriched: list[dict[str, Any]] = [dict(obs) for obs in observations]
        obs_entities = [obs[observation_entity_key] for obs in observations]
        obs_ts_ns = np.asarray(
            [_to_datetime(obs[observation_timestamp_key]).timestamp() * 1_000_000_000 for obs in observations],
            dtype=np.int64,
        )

        if not snapshots:
            self._cache_point_in_time_join(cache_key, enriched)
            return enriched

        for feature_name, snapshot in snapshots.items():
            prepared = self._prepared.get((feature_name, snapshot.version))
            if prepared is None:
                prepared = _prepare_snapshot(snapshot)
                self._prepared[(feature_name, snapshot.version)] = prepared
            lag_ns = int(snapshot.metadata.lag.total_seconds() * 1_000_000_000)
            values, asofs = _vectorized_asof_lookup(
                prepared=prepared,
                entities=obs_entities,
                cutoffs_ns=obs_ts_ns - lag_ns,
            )

            for idx, row in enumerate(enriched):
                if strict and asofs[idx] is not None and asofs[idx] > _ns_to_datetime(obs_ts_ns[idx] - lag_ns):
                    raise FeatureLeakageError(
                        f"Leakage detected for '{feature_name}' entity={obs_entities[idx]}: "
                        f"candidate timestamp {asofs[idx].isoformat()} exceeds cutoff "
                        f"{_ns_to_datetime(obs_ts_ns[idx] - lag_ns).isoformat()}"
                    )
                row[feature_name] = values[idx]
                row[f"{feature_name}__asof"] = asofs[idx]
                row[f"{feature_name}__version"] = snapshot.version
                row[f"{feature_name}__source"] = snapshot.metadata.source
                row[f"{feature_name}__lag_seconds"] = int(snapshot.metadata.lag.total_seconds())
                row[f"{feature_name}__refresh_cadence_seconds"] = int(snapshot.metadata.refresh_cadence.total_seconds())
                row[f"{feature_name}__data_source_hash"] = snapshot.version_keys.data_source_hash
                row[f"{feature_name}__transformation_graph_hash"] = snapshot.version_keys.transformation_graph_hash
                row[f"{feature_name}__training_window_metadata"] = dict(snapshot.version_keys.training_window_metadata)
                row[f"{feature_name}__feature_set_id"] = snapshot.version_keys.feature_set_id

        self._cache_point_in_time_join(cache_key, enriched)
        return enriched

    def build_nextgen_feature_set_registry(
        self,
        *,
        feature_versions: dict[str, int | None] | None = None,
    ) -> dict[str, Any]:
        requested = feature_versions or {}
        names = sorted(set(requested) if requested else set(self._snapshots))
        feature_set_ids: dict[str, str] = {}
        version_keys: dict[str, dict[str, Any]] = {}
        for name in names:
            snapshot = self.get_snapshot(name, requested.get(name))
            feature_set_ids[name] = snapshot.version_keys.feature_set_id
            version_keys[name] = {
                "version": snapshot.version,
                "data_source_hash": snapshot.version_keys.data_source_hash,
                "transformation_graph_hash": snapshot.version_keys.transformation_graph_hash,
                "training_window_metadata": dict(snapshot.version_keys.training_window_metadata),
                "feature_set_id": snapshot.version_keys.feature_set_id,
            }
        return {
            "immutable_feature_set_ids": feature_set_ids,
            "feature_version_keys": version_keys,
        }

    def export_snapshot_columnar(
        self,
        *,
        feature_name: str,
        version: int | None = None,
        output_path: str | Path,
        format: str = "parquet",
    ) -> Path:
        if pa is None:
            raise RuntimeError("pyarrow is required for Parquet/Arrow exports")

        snapshot = self.get_snapshot(feature_name, version)
        path = Path(output_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        table = pa.Table.from_pylist(list(snapshot.records))
        fmt = format.lower().strip()
        if fmt == "parquet":
            pq.write_table(table, path)
        elif fmt in {"arrow", "feather"}:
            feather.write_feather(table, path)
        else:
            raise ValueError("format must be one of: parquet, arrow, feather")
        return path


    def point_in_time_cache_info(self) -> dict[str, int]:
        return {
            "entries": len(self._point_in_time_cache),
            "capacity": self._point_in_time_cache_size,
        }

    def _cache_point_in_time_join(self, key: tuple[Any, ...], rows: list[dict[str, Any]]) -> None:
        self._point_in_time_cache[key] = [dict(item) for item in rows]
        if len(self._point_in_time_cache) > self._point_in_time_cache_size:
            oldest = next(iter(self._point_in_time_cache))
            self._point_in_time_cache.pop(oldest, None)


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


def _prepare_snapshot(snapshot: FeatureSnapshot) -> _PreparedSnapshot:
    by_entity: dict[Any, list[tuple[int, Any]]] = {}
    for record in snapshot.records:
        entity = record[snapshot.entity_key]
        ts = _to_datetime(record[snapshot.timestamp_key])
        ns = int(ts.timestamp() * 1_000_000_000)
        by_entity.setdefault(entity, []).append((ns, record[snapshot.value_key]))

    prepared: dict[Any, tuple[np.ndarray, np.ndarray]] = {}
    for entity, pairs in by_entity.items():
        pairs.sort(key=lambda item: item[0])
        ts = np.asarray([item[0] for item in pairs], dtype=np.int64)
        values = np.asarray([item[1] for item in pairs], dtype=object)
        prepared[entity] = (ts, values)
    return _PreparedSnapshot(entities=prepared)


def _vectorized_asof_lookup(*, prepared: _PreparedSnapshot, entities: list[Any], cutoffs_ns: np.ndarray) -> tuple[list[Any], list[datetime | None]]:
    values: list[Any] = [None] * len(entities)
    asofs: list[datetime | None] = [None] * len(entities)

    idx_by_entity: dict[Any, list[int]] = {}
    for idx, entity in enumerate(entities):
        idx_by_entity.setdefault(entity, []).append(idx)

    for entity, idxs in idx_by_entity.items():
        payload = prepared.entities.get(entity)
        if payload is None:
            continue
        entity_ts, entity_values = payload
        entity_cutoffs = cutoffs_ns[np.asarray(idxs, dtype=int)]
        pos = np.searchsorted(entity_ts, entity_cutoffs, side="right") - 1
        valid = pos >= 0
        for local, row_idx in enumerate(idxs):
            if not valid[local]:
                continue
            pick = int(pos[local])
            values[row_idx] = entity_values[pick].item() if hasattr(entity_values[pick], "item") else entity_values[pick]
            asofs[row_idx] = _ns_to_datetime(int(entity_ts[pick]))
    return values, asofs


def _build_point_in_time_cache_key(
    *,
    observations: list[dict[str, Any]],
    snapshots: dict[str, FeatureSnapshot],
    observation_entity_key: str,
    observation_timestamp_key: str,
    strict: bool,
) -> tuple[Any, ...]:
    obs_sig = tuple(
        (obs.get(observation_entity_key), _to_datetime(obs[observation_timestamp_key]).isoformat())
        for obs in observations
    )
    snap_sig = tuple(sorted((name, snap.version) for name, snap in snapshots.items()))
    return (observation_entity_key, observation_timestamp_key, strict, obs_sig, snap_sig)


def _normalize_training_window_metadata(training_window_metadata: dict[str, Any] | None) -> dict[str, Any]:
    payload = dict(training_window_metadata or {})
    normalized: dict[str, Any] = {}
    for key, value in payload.items():
        if isinstance(value, datetime):
            normalized[str(key)] = value.isoformat()
        elif isinstance(value, date):
            normalized[str(key)] = value.isoformat()
        else:
            normalized[str(key)] = value
    return normalized


def _build_feature_version_keys(
    *,
    feature_name: str,
    version: int,
    metadata: FeatureMetadata,
    data_source_hash: str | None,
    transformation_graph_hash: str | None,
    training_window_metadata: dict[str, Any],
) -> FeatureVersionKeys:
    resolved_data_source_hash = str(data_source_hash or _stable_hash({"source": metadata.source}))
    resolved_transformation_graph_hash = str(
        transformation_graph_hash
        or _stable_hash(
            {
                "feature_name": feature_name,
                "lag_seconds": metadata.lag.total_seconds(),
                "refresh_cadence_seconds": metadata.refresh_cadence.total_seconds(),
            }
        )
    )
    feature_set_id = _stable_hash(
        {
            "feature_name": feature_name,
            "version": version,
            "source": metadata.source,
            "data_source_hash": resolved_data_source_hash,
            "transformation_graph_hash": resolved_transformation_graph_hash,
            "training_window_metadata": training_window_metadata,
        }
    )
    return FeatureVersionKeys(
        data_source_hash=resolved_data_source_hash,
        transformation_graph_hash=resolved_transformation_graph_hash,
        training_window_metadata=training_window_metadata,
        feature_set_id=feature_set_id,
    )


def _stable_hash(payload: dict[str, Any]) -> str:
    encoded = json.dumps(payload, sort_keys=True, default=str, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def _ns_to_datetime(value: int) -> datetime:
    return datetime.utcfromtimestamp(value / 1_000_000_000)

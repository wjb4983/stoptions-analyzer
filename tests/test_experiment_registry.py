from __future__ import annotations

import csv
import json
from pathlib import Path

from src.backtesting.experiment_registry import (
    append_experiment_entry,
    read_registry,
    registry_db_path,
)


def _entry(*, run_id: str, timestamp: str, metric_schema_version: str, score: float) -> dict[str, object]:
    return {
        "timestamp": timestamp,
        "run_type": "unit",
        "run_id": run_id,
        "run_dir": f"runs/{run_id}",
        "code_version": "abc123",
        "metric_schema_version": metric_schema_version,
        "random_seed": 7,
        "primary_metric": "sharpe",
        "primary_metric_value": score,
        "config_hash": f"cfg-{run_id}",
        "config_checksum": f"cfgchk-{run_id}",
        "data_snapshot_checksum": f"snap-{run_id}",
        "manifest_path": f"manifests/{run_id}.json",
        "reproducibility_fingerprint": f"fp-{run_id}",
        "manifest_checksum": f"mchk-{run_id}",
        "parameters": {"lookback": 30},
        "data_snapshot_identifiers": {"prices": "daily.v1"},
        "governance": {"approved": True},
        "metrics": {"sharpe": score, "max_drawdown": -0.1},
        "significance": {"p_value": 0.01},
        "model_artifacts": [f"models/{run_id}.pkl"],
        "plot_artifacts": [f"plots/{run_id}.png"],
        "metric_artifacts": [f"metrics/{run_id}.json"],
        "reproducibility_metadata": {"python": "3.11"},
    }


def test_registration_uniqueness_replaces_existing_run_id(tmp_path: Path) -> None:
    output_dir = tmp_path / "registry"
    first = _entry(run_id="run-001", timestamp="2024-01-01T00:00:00", metric_schema_version="v1", score=1.2)
    replacement = _entry(run_id="run-001", timestamp="2024-01-02T00:00:00", metric_schema_version="v2", score=1.8)

    append_experiment_entry(output_dir, first)
    append_experiment_entry(output_dir, replacement)

    rows = read_registry(output_dir)
    assert len(rows) == 1
    assert rows[0]["run_id"] == "run-001"
    assert rows[0]["metric_schema_version"] == "v2"
    assert rows[0]["metrics"]["sharpe"] == 1.8


def test_metadata_persistence_across_jsonl_csv_and_sqlite(tmp_path: Path) -> None:
    output_dir = tmp_path / "registry"
    entry = _entry(run_id="run-abc", timestamp="2024-01-03T12:30:00", metric_schema_version="v3", score=2.1)

    append_experiment_entry(output_dir, entry)

    jsonl_path = output_dir / "experiment_index.jsonl"
    parsed_jsonl = [json.loads(line) for line in jsonl_path.read_text(encoding="utf-8").splitlines() if line.strip()]
    assert parsed_jsonl[-1]["manifest_path"] == "manifests/run-abc.json"
    assert parsed_jsonl[-1]["reproducibility_metadata"]["python"] == "3.11"

    csv_path = output_dir / "experiment_index.csv"
    with csv_path.open("r", encoding="utf-8", newline="") as handle:
        rows = list(csv.DictReader(handle))
    assert rows[-1]["run_id"] == "run-abc"
    assert json.loads(rows[-1]["metrics_json"])["sharpe"] == 2.1
    assert json.loads(rows[-1]["governance_json"])["approved"] is True

    assert read_registry(output_dir)[0]["model_artifacts"] == ["models/run-abc.pkl"]


def test_lookup_and_versioning_behavior_prefers_sqlite_then_jsonl_fallback(tmp_path: Path) -> None:
    output_dir = tmp_path / "registry"
    append_experiment_entry(
        output_dir,
        _entry(run_id="run-v1", timestamp="2024-01-01T00:00:00", metric_schema_version="v1", score=0.9),
    )
    append_experiment_entry(
        output_dir,
        _entry(run_id="run-v2", timestamp="2024-01-02T00:00:00", metric_schema_version="v2", score=1.1),
    )

    sqlite_rows = read_registry(output_dir)
    assert [row["run_id"] for row in sqlite_rows] == ["run-v2", "run-v1"]
    assert [row["metric_schema_version"] for row in sqlite_rows] == ["v2", "v1"]

    registry_db_path(output_dir).unlink()
    jsonl_rows = read_registry(output_dir)
    assert [row["run_id"] for row in jsonl_rows] == ["run-v2", "run-v1"]
    assert [row["metric_schema_version"] for row in jsonl_rows] == ["v2", "v1"]

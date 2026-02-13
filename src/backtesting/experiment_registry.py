from __future__ import annotations

import csv
import json
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any


def registry_db_path(output_dir: Path) -> Path:
    return output_dir / "experiment_registry.sqlite3"


def _ensure_schema(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS experiments (
            run_id TEXT PRIMARY KEY,
            run_dir TEXT,
            run_type TEXT,
            timestamp TEXT,
            config_hash TEXT,
            config_checksum TEXT,
            data_snapshot_identifiers_json TEXT,
            data_snapshot_checksum TEXT,
            metrics_json TEXT,
            significance_json TEXT,
            governance_json TEXT,
            manifest_path TEXT,
            reproducibility_fingerprint TEXT,
            manifest_checksum TEXT,
            raw_entry_json TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS governance_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id TEXT NOT NULL,
            timestamp TEXT NOT NULL,
            action TEXT NOT NULL,
            reason TEXT NOT NULL,
            actor TEXT,
            resulting_promotion_state TEXT,
            resulting_approval_status TEXT,
            FOREIGN KEY(run_id) REFERENCES experiments(run_id)
        )
        """
    )


def append_experiment_entry(output_dir: Path, entry: dict[str, Any]) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    jsonl_path = output_dir / "experiment_index.jsonl"
    with jsonl_path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(entry, sort_keys=True) + "\n")

    csv_path = output_dir / "experiment_index.csv"
    row = {
        "timestamp": entry.get("timestamp"),
        "run_type": entry.get("run_type"),
        "run_id": entry.get("run_id"),
        "run_dir": entry.get("run_dir"),
        "code_version": entry.get("code_version"),
        "metric_schema_version": entry.get("metric_schema_version"),
        "random_seed": entry.get("random_seed"),
        "primary_metric": entry.get("primary_metric"),
        "primary_metric_value": entry.get("primary_metric_value"),
        "config_hash": entry.get("config_hash"),
        "config_checksum": entry.get("config_checksum"),
        "data_snapshot_checksum": entry.get("data_snapshot_checksum"),
        "manifest_path": entry.get("manifest_path"),
        "reproducibility_fingerprint": entry.get("reproducibility_fingerprint"),
        "manifest_checksum": entry.get("manifest_checksum"),
        "parameters_json": json.dumps(entry.get("parameters", {}), sort_keys=True),
        "data_snapshot_json": json.dumps(entry.get("data_snapshot_identifiers", {}), sort_keys=True),
        "governance_json": json.dumps(entry.get("governance", {}), sort_keys=True),
        "metrics_json": json.dumps(entry.get("metrics", {}), sort_keys=True),
        "significance_json": json.dumps(entry.get("significance", {}), sort_keys=True),
    }
    fieldnames = list(row.keys())
    write_header = not csv_path.exists()
    with csv_path.open("a", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        if write_header:
            writer.writeheader()
        writer.writerow(row)

    with sqlite3.connect(registry_db_path(output_dir)) as conn:
        _ensure_schema(conn)
        conn.execute(
            """
            INSERT OR REPLACE INTO experiments(
                run_id, run_dir, run_type, timestamp, config_hash, config_checksum,
                data_snapshot_identifiers_json, data_snapshot_checksum,
                metrics_json, significance_json, governance_json,
                manifest_path, reproducibility_fingerprint, manifest_checksum, raw_entry_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                str(entry.get("run_id", "")),
                str(entry.get("run_dir", "")),
                str(entry.get("run_type", "")),
                str(entry.get("timestamp", "")),
                str(entry.get("config_hash", "")),
                str(entry.get("config_checksum", "")),
                json.dumps(entry.get("data_snapshot_identifiers", {}), sort_keys=True),
                str(entry.get("data_snapshot_checksum", "")),
                json.dumps(entry.get("metrics", {}), sort_keys=True),
                json.dumps(entry.get("significance", {}), sort_keys=True),
                json.dumps(entry.get("governance", {}), sort_keys=True),
                str(entry.get("manifest_path", "")),
                str(entry.get("reproducibility_fingerprint", "")),
                str(entry.get("manifest_checksum", "")),
                json.dumps(entry, sort_keys=True),
            ),
        )
        conn.commit()


def read_registry(output_dir: Path) -> list[dict[str, Any]]:
    db_path = registry_db_path(output_dir)
    if db_path.exists():
        with sqlite3.connect(db_path) as conn:
            _ensure_schema(conn)
            rows = conn.execute(
                "SELECT raw_entry_json FROM experiments ORDER BY timestamp DESC"
            ).fetchall()
        result: list[dict[str, Any]] = []
        for (payload,) in rows:
            try:
                parsed = json.loads(str(payload))
            except json.JSONDecodeError:
                continue
            if isinstance(parsed, dict):
                result.append(parsed)
        return result

    jsonl_path = output_dir / "experiment_index.jsonl"
    rows: list[dict[str, Any]] = []
    if jsonl_path.exists():
        for line in jsonl_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line:
                continue
            try:
                parsed = json.loads(line)
            except json.JSONDecodeError:
                continue
            if isinstance(parsed, dict):
                rows.append(parsed)
    rows.reverse()
    return rows


def append_governance_event(
    output_dir: Path,
    *,
    run_id: str,
    action: str,
    reason: str,
    resulting_promotion_state: str,
    resulting_approval_status: str,
    actor: str = "ui",
) -> None:
    with sqlite3.connect(registry_db_path(output_dir)) as conn:
        _ensure_schema(conn)
        conn.execute(
            """
            INSERT INTO governance_events(run_id, timestamp, action, reason, actor, resulting_promotion_state, resulting_approval_status)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                datetime.now().isoformat(),
                action,
                reason,
                actor,
                resulting_promotion_state,
                resulting_approval_status,
            ),
        )
        conn.commit()


from __future__ import annotations

import json
from pathlib import Path

from execution.artifact_sync import sync_run_artifacts, update_remote_sync_mapping, write_artifact_manifest


def test_write_artifact_manifest_uses_relative_paths(tmp_path: Path) -> None:
    run_dir = tmp_path / "run_a"
    run_dir.mkdir()
    nested = run_dir / "sub" / "metrics.json"
    nested.parent.mkdir()
    nested.write_text('{"sharpe": 1.2}', encoding="utf-8")

    manifest_path = write_artifact_manifest(run_dir=run_dir, workflow="backtesting", run_id="run_a")
    payload = json.loads(manifest_path.read_text(encoding="utf-8"))

    assert payload["run_dir"] == "."
    assert payload["workflow"] == "backtesting"
    assert payload["artifact_count"] == 1
    assert payload["artifacts"][0]["path"] == "sub/metrics.json"


def test_sync_run_artifacts_summary_then_full(tmp_path: Path) -> None:
    source = tmp_path / "remote"
    source.mkdir()
    (source / "manifest.json").write_text("{}", encoding="utf-8")
    (source / "metrics.json").write_text("{}", encoding="utf-8")
    (source / "huge_blob.bin").write_bytes(b"x" * 4096)

    target_root = tmp_path / "local"
    summary_result = sync_run_artifacts(
        remote_job_id="job123",
        remote_run_dir=source,
        local_output_root=target_root,
        mode="summary",
    )
    assert (summary_result.local_synced_run_dir / "manifest.json").exists()
    assert not (summary_result.local_synced_run_dir / "huge_blob.bin").exists()

    full_result = sync_run_artifacts(
        remote_job_id="job123",
        remote_run_dir=source,
        local_output_root=target_root,
        mode="full",
    )
    assert (full_result.local_synced_run_dir / "huge_blob.bin").exists()


def test_update_remote_sync_mapping_overwrites_existing() -> None:
    merged = update_remote_sync_mapping(
        existing_mapping={"jobA": "/tmp/a"},
        remote_job_id="jobA",
        local_synced_run_dir="/tmp/new_a",
    )
    assert merged["jobA"] == str(Path("/tmp/new_a").resolve())

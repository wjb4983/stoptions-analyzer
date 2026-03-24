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


def test_sync_run_artifacts_summary_selected_and_full_modes(tmp_path: Path) -> None:
    source = tmp_path / "remote"
    source.mkdir()
    (source / "manifest.json").write_text("{}", encoding="utf-8")
    (source / "metrics.json").write_text("{}", encoding="utf-8")
    summary_dir = source / "summary"
    summary_dir.mkdir()
    (summary_dir / "leaderboard_stats.json").write_text("{}", encoding="utf-8")
    (source / "huge_blob.bin").write_bytes(b"x" * 4096)

    target_root = tmp_path / "local"
    summary_result = sync_run_artifacts(
        remote_job_id="job123",
        remote_run_dir=source,
        local_output_root=target_root,
        mode="summary_only",
    )
    assert (summary_result.local_synced_run_dir / "summary" / "leaderboard_stats.json").exists()
    assert not (summary_result.local_synced_run_dir / "huge_blob.bin").exists()

    selected_result = sync_run_artifacts(
        remote_job_id="job123",
        remote_run_dir=source,
        local_output_root=target_root,
        mode="selected_files",
        include_files=["metrics.json"],
    )
    assert (selected_result.local_synced_run_dir / "metrics.json").exists()

    full_result = sync_run_artifacts(
        remote_job_id="job123",
        remote_run_dir=source,
        local_output_root=target_root,
        mode="full_artifacts",
    )
    assert (full_result.local_synced_run_dir / "huge_blob.bin").exists()


def test_update_remote_sync_mapping_overwrites_existing() -> None:
    merged = update_remote_sync_mapping(
        existing_mapping={"jobA": "/tmp/a"},
        remote_job_id="jobA",
        local_synced_run_dir="/tmp/new_a",
    )
    assert merged["jobA"] == str(Path("/tmp/new_a").resolve())

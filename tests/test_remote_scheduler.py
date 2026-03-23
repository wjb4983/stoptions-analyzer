from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from remote.scheduler import SCHEDULER_STATE_FILENAME, tick_scheduler


def _write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def _make_job(remote_root: Path, *, job_id: str, job_type: str, status: str = "queued", blocked_by: str | None = None) -> Path:
    job_dir = remote_root / job_id
    envelope = {
        "schema_version": 1,
        "run_id": job_id,
        "job_id": job_id,
        "job_type": job_type,
        "params": {},
        "timestamps": {"created_at": "2026-01-01T00:00:00+00:00", "submitted_at": "2026-01-01T00:00:00+00:00"},
    }
    status_payload = {
        "schema_version": 1,
        "job_id": job_id,
        "run_id": job_id,
        "job_type": job_type,
        "status": status,
        "blocked_by": blocked_by,
        "timestamps": {"created_at": "2026-01-01T00:00:00+00:00", "submitted_at": "2026-01-01T00:00:00+00:00", "started_at": None, "completed_at": None},
        "error": None,
    }
    _write_json(job_dir / "job_request.json", envelope)
    _write_json(job_dir / "status.json", status_payload)
    _write_json(job_dir / "summary.json", status_payload)
    (job_dir / "logs.txt").write_text("", encoding="utf-8")
    return job_dir


def test_scheduler_blocks_only_exclusive_jobs(tmp_path: Path) -> None:
    remote_root = tmp_path / "remote"
    _make_job(remote_root, job_id="job-exclusive-1", job_type="backtest")
    _make_job(remote_root, job_id="job-exclusive-2", job_type="gpu_training")
    _make_job(remote_root, job_id="job-concurrent-1", job_type="general_analysis")

    launched: list[str] = []

    def _launcher(job_file: Path) -> int:
        launched.append(job_file.parent.name)
        payload = json.loads((job_file.parent / "status.json").read_text(encoding="utf-8"))
        payload["status"] = "running"
        payload["blocked_by"] = None
        (job_file.parent / "status.json").write_text(json.dumps(payload, indent=2), encoding="utf-8")
        return 4321

    tick_scheduler(remote_root, launcher=_launcher)

    assert "job-exclusive-1" in launched
    assert "job-concurrent-1" in launched
    assert "job-exclusive-2" not in launched

    blocked = json.loads((remote_root / "job-exclusive-2" / "status.json").read_text(encoding="utf-8"))
    assert blocked["status"] == "queued"
    assert blocked["blocked_by"] == "exclusive_job:job-exclusive-1"

    state = json.loads((remote_root / SCHEDULER_STATE_FILENAME).read_text(encoding="utf-8"))
    assert state["running_exclusive_job_id"] == "job-exclusive-1"
    assert "job-exclusive-2" in state["queue"]


def test_scheduler_releases_exclusive_slot_after_cancellation(tmp_path: Path) -> None:
    remote_root = tmp_path / "remote"
    _make_job(remote_root, job_id="job-exclusive-running", job_type="backtest", status="canceled")
    _make_job(remote_root, job_id="job-exclusive-queued", job_type="gpu_training", status="queued", blocked_by="exclusive_job:job-exclusive-running")

    launched: list[str] = []

    def _launcher(job_file: Path) -> int:
        launched.append(job_file.parent.name)
        payload = json.loads((job_file.parent / "status.json").read_text(encoding="utf-8"))
        payload["status"] = "running"
        payload["blocked_by"] = None
        (job_file.parent / "status.json").write_text(json.dumps(payload, indent=2), encoding="utf-8")
        return 999

    tick_scheduler(remote_root, launcher=_launcher)

    assert launched == ["job-exclusive-queued"]
    state = json.loads((remote_root / SCHEDULER_STATE_FILENAME).read_text(encoding="utf-8"))
    assert state["running_exclusive_job_id"] == "job-exclusive-queued"
    assert "job-exclusive-queued" not in state["queue"]

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import subprocess
import sys
import time
from typing import Any, Callable

from execution.contracts import SCHEMA_VERSION, normalize_job_state

EXCLUSIVE_JOB_TYPES = {"backtest", "gpu_training"}
SCHEDULER_STATE_FILENAME = "scheduler_state.json"
SCHEDULER_LOCK_FILENAME = "scheduler.lock"
REGISTRY_FILENAME = "registry.jsonl"
TERMINAL_STATUSES = {"succeeded", "failed", "canceled", "completed"}


Launcher = Callable[[Path], int | None]


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _normalize_job_class(job_type: str) -> str:
    return "exclusive" if str(job_type).strip().lower() in EXCLUSIVE_JOB_TYPES else "concurrent"


def _read_json(path: Path) -> dict[str, Any] | None:
    if not path.exists():
        return None
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return None
    return payload if isinstance(payload, dict) else None


def _write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, default=str), encoding="utf-8")


def _append_registry_event(remote_root: Path, payload: dict[str, Any]) -> None:
    registry_path = remote_root / REGISTRY_FILENAME
    registry_path.parent.mkdir(parents=True, exist_ok=True)
    with registry_path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps({"timestamp": _now(), **payload}, default=str))
        handle.write("\n")


def _status_payload(*, envelope: dict[str, Any], status: str, started_at: str | None, completed_at: str | None = None, error: str | None = None, blocked_by: str | None = None) -> dict[str, Any]:
    return {
        "schema_version": SCHEMA_VERSION,
        "job_id": envelope.get("job_id") or envelope.get("run_id"),
        "run_id": envelope.get("run_id"),
        "job_type": envelope.get("job_type"),
        "status": normalize_job_state(status).value,
        "blocked_by": blocked_by,
        "timestamps": {
            "created_at": envelope.get("timestamps", {}).get("created_at"),
            "submitted_at": envelope.get("timestamps", {}).get("submitted_at"),
            "started_at": started_at,
            "completed_at": completed_at,
        },
        "error": error,
    }


def _summary_payload(*, envelope: dict[str, Any], status: str, started_at: str | None, completed_at: str | None = None, error: str | None = None, blocked_by: str | None = None) -> dict[str, Any]:
    return {
        "schema_version": SCHEMA_VERSION,
        "job_id": envelope.get("job_id") or envelope.get("run_id"),
        "job_type": envelope.get("job_type"),
        "status": normalize_job_state(status).value,
        "blocked_by": blocked_by,
        "timestamps": {
            "created_at": envelope.get("timestamps", {}).get("created_at"),
            "submitted_at": envelope.get("timestamps", {}).get("submitted_at"),
            "started_at": started_at,
            "completed_at": completed_at,
        },
        "error": error,
    }


def _load_state(remote_root: Path) -> dict[str, Any]:
    payload = _read_json(remote_root / SCHEDULER_STATE_FILENAME) or {}
    jobs = payload.get("jobs")
    queue = payload.get("queue")
    return {
        "schema_version": SCHEMA_VERSION,
        "last_updated": payload.get("last_updated"),
        "running_exclusive_job_id": str(payload.get("running_exclusive_job_id", "")).strip() or None,
        "queue": [str(job_id) for job_id in queue] if isinstance(queue, list) else [],
        "jobs": jobs if isinstance(jobs, dict) else {},
    }


def _save_state(remote_root: Path, state: dict[str, Any]) -> None:
    state["schema_version"] = SCHEMA_VERSION
    state["last_updated"] = _now()
    path = remote_root / SCHEDULER_STATE_FILENAME
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(state, indent=2, default=str), encoding="utf-8")
    tmp.replace(path)


def _default_launcher(job_file: Path) -> int | None:
    logs_path = job_file.parent / "logs.txt"
    logs_path.touch(exist_ok=True)
    with logs_path.open("a", encoding="utf-8") as log_handle:
        process = subprocess.Popen(
            [sys.executable, "-m", "remote.worker", "--job-file", str(job_file)],
            cwd=str(job_file.parent),
            stdout=log_handle,
            stderr=subprocess.STDOUT,
            start_new_session=True,
        )
    return int(process.pid)


def _job_dirs(remote_root: Path) -> list[Path]:
    result: list[Path] = []
    for item in sorted(remote_root.iterdir()):
        if not item.is_dir():
            continue
        if (item / "job_request.json").exists():
            result.append(item)
    return result


def _acquire_lock(lock_path: Path, timeout_seconds: float = 5.0) -> bool:
    deadline = time.monotonic() + max(0.2, timeout_seconds)
    while time.monotonic() < deadline:
        try:
            fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(fd)
            return True
        except FileExistsError:
            time.sleep(0.05)
    return False


def _release_lock(lock_path: Path) -> None:
    try:
        lock_path.unlink()
    except FileNotFoundError:
        return


def tick_scheduler(remote_root: Path, *, launcher: Launcher | None = None) -> dict[str, Any]:
    remote_root = remote_root.expanduser().resolve()
    remote_root.mkdir(parents=True, exist_ok=True)
    lock_path = remote_root / SCHEDULER_LOCK_FILENAME
    if not _acquire_lock(lock_path):
        return {"status": "busy", "remote_root": str(remote_root)}

    active_launcher = launcher or _default_launcher
    try:
        state = _load_state(remote_root)
        jobs = state["jobs"]
        known_queue = list(state.get("queue", []))
        queue: list[str] = []
        queue_seen: set[str] = set()

        running_exclusive_job_id = None

        for job_dir in _job_dirs(remote_root):
            envelope = _read_json(job_dir / "job_request.json") or {}
            job_id = str(envelope.get("job_id") or envelope.get("run_id") or job_dir.name).strip()
            if not job_id:
                continue
            job_type = str(envelope.get("job_type", "")).strip()
            status_payload = _read_json(job_dir / "status.json") or {}
            summary_path = job_dir / "summary.json"
            cancel_path = job_dir / "cancel.requested"
            status = str(status_payload.get("status", "queued")).strip().lower() or "queued"
            blocked_by = status_payload.get("blocked_by")
            started_at = status_payload.get("timestamps", {}).get("started_at") if isinstance(status_payload.get("timestamps"), dict) else None
            completed_at = status_payload.get("timestamps", {}).get("completed_at") if isinstance(status_payload.get("timestamps"), dict) else None
            error = status_payload.get("error")

            job_record = jobs.get(job_id)
            if not isinstance(job_record, dict):
                job_record = {}
            previous_status = str(job_record.get("status", status)).strip().lower() or status
            job_class = _normalize_job_class(job_type)
            child_pid = job_record.get("child_pid")

            if status in TERMINAL_STATUSES:
                blocked_by = None
            elif cancel_path.exists() and status in {"queued", "canceling"}:
                completed_at = _now()
                status = "canceled"
                error = "Canceled before execution"
                blocked_by = None
                _write_json(job_dir / "status.json", _status_payload(envelope=envelope, status="canceled", started_at=started_at, completed_at=completed_at, error=error))
                _write_json(summary_path, _summary_payload(envelope=envelope, status="canceled", started_at=started_at, completed_at=completed_at, error=error))
                _append_registry_event(remote_root, {"event": "completed", "job_id": job_id, "job_type": job_type, "status": "canceled", "error": error})

            if status == "running" and job_class == "exclusive":
                running_exclusive_job_id = job_id

            if status == "queued" and job_id not in queue_seen:
                queue_seen.add(job_id)
                queue.append(job_id)

            jobs[job_id] = {
                "job_id": job_id,
                "job_type": job_type,
                "job_class": job_class,
                "status": status,
                "blocked_by": blocked_by,
                "remote_dir": str(job_dir),
                "child_pid": child_pid,
                "updated_at": _now(),
            }

            if previous_status != status:
                _append_registry_event(
                    remote_root,
                    {"event": "scheduler_state", "job_id": job_id, "job_type": job_type, "from_status": previous_status, "to_status": status},
                )

        for queued_job in known_queue:
            if queued_job in jobs and queued_job not in queue_seen and str(jobs[queued_job].get("status")) == "queued":
                queue.append(queued_job)
                queue_seen.add(queued_job)

        for job_id in queue:
            record = jobs.get(job_id)
            if not isinstance(record, dict):
                continue
            if str(record.get("status")) != "queued":
                continue
            job_class = str(record.get("job_class", "concurrent"))
            job_dir = Path(str(record.get("remote_dir", "")))
            envelope = _read_json(job_dir / "job_request.json") or {}
            if job_class == "exclusive" and running_exclusive_job_id and running_exclusive_job_id != job_id:
                blocked_by = f"exclusive_job:{running_exclusive_job_id}"
                if record.get("blocked_by") != blocked_by:
                    _write_json(
                        job_dir / "status.json",
                        _status_payload(
                            envelope=envelope,
                            status="queued",
                            started_at=None,
                            blocked_by=blocked_by,
                        ),
                    )
                    _write_json(
                        job_dir / "summary.json",
                        _summary_payload(envelope=envelope, status="queued", started_at=None, blocked_by=blocked_by),
                    )
                    _append_registry_event(remote_root, {"event": "queued_blocked", "job_id": job_id, "job_type": record.get("job_type"), "blocked_by": blocked_by})
                record["blocked_by"] = blocked_by
                continue

            child_pid = active_launcher(job_dir / "job_request.json")
            record["status"] = "running"
            record["blocked_by"] = None
            record["child_pid"] = child_pid
            record["started_at"] = _now()
            _append_registry_event(
                remote_root,
                {
                    "event": "scheduler_dispatch",
                    "job_id": job_id,
                    "job_type": record.get("job_type"),
                    "job_class": job_class,
                    "child_pid": child_pid,
                },
            )
            if job_class == "exclusive":
                running_exclusive_job_id = job_id

        state["queue"] = [job_id for job_id in queue if isinstance(jobs.get(job_id), dict) and str(jobs[job_id].get("status")) == "queued"]
        state["running_exclusive_job_id"] = running_exclusive_job_id
        state["jobs"] = jobs
        _save_state(remote_root, state)
        return {
            "status": "ok",
            "running_exclusive_job_id": running_exclusive_job_id,
            "queued": len(state["queue"]),
            "jobs": len(jobs),
        }
    finally:
        _release_lock(lock_path)


def main() -> int:
    parser = argparse.ArgumentParser(description="Run one scheduler tick for remote jobs")
    parser.add_argument("--remote-root", required=True, help="Remote jobs root directory")
    args = parser.parse_args()
    tick_scheduler(Path(args.remote_root))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import json
from pathlib import Path
import threading
import traceback
from typing import Any

from backtesting.cache_runner import (
    CancellationToken,
    TaskCancellationError,
    run_backtest_cache,
    run_multi_signal_backtest,
    run_strategy_optimization,
    run_time_series_momentum_backtest,
    run_trained_regime_backtest,
    run_walk_forward_backtest,
)
from backtesting.regime_training_pipeline import execute_regime_training_pipeline
from execution.backend import (
    JOB_ANALYSIS_CALLABLE,
    JOB_BACKTEST_CACHE,
    JOB_BACKTEST_MULTI_SIGNAL,
    JOB_BACKTEST_OPTIMIZATION,
    JOB_BACKTEST_TIME_SERIES,
    JOB_BACKTEST_TRAINED_REGIME,
    JOB_BACKTEST_WALK_FORWARD,
    JOB_REGIME_TRAINING,
)
from execution.contracts import SCHEMA_VERSION, ensure_schema_compatible, normalize_job_state
from execution.remote_payloads import deserialize_from_json, serialize_for_json
from execution.run_summary import build_workflow_summary
from remote.scheduler import tick_scheduler


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _append_log(log_path: Path, message: str) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with log_path.open("a", encoding="utf-8") as handle:
        handle.write(f"[{_now()}] {message}\n")


def _write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, default=str), encoding="utf-8")


def _status_payload(*, envelope: dict[str, Any], status: str, started_at: str, completed_at: str | None = None, error: str | None = None, blocked_by: str | None = None) -> dict[str, Any]:
    normalized_status = normalize_job_state(status).value
    return {
        "schema_version": SCHEMA_VERSION,
        "job_id": envelope.get("job_id") or envelope.get("run_id"),
        "run_id": envelope.get("run_id"),
        "job_type": envelope.get("job_type"),
        "status": normalized_status,
        "blocked_by": blocked_by,
        "timestamps": {
            "created_at": envelope.get("timestamps", {}).get("created_at"),
            "submitted_at": envelope.get("timestamps", {}).get("submitted_at"),
            "started_at": started_at,
            "completed_at": completed_at,
        },
        "error": error,
    }


def _summary_payload(*, envelope: dict[str, Any], status: str, started_at: str, completed_at: str | None = None, error: str | None = None, blocked_by: str | None = None) -> dict[str, Any]:
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


def _append_registry_event(job_dir: Path, payload: dict[str, Any]) -> None:
    registry_path = job_dir.parent / "registry.jsonl"
    registry_path.parent.mkdir(parents=True, exist_ok=True)
    with registry_path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps({"timestamp": _now(), **payload}, default=str))
        handle.write("\n")


def _artifact_manifest(job_dir: Path) -> dict[str, Any]:
    tracked = [
        "job_request.json",
        "status.json",
        "logs.txt",
        "result.json",
        "artifacts.json",
        "cancel.requested",
    ]
    artifacts: list[dict[str, Any]] = []
    for item in sorted(job_dir.iterdir()):
        if item.name.startswith(".") or not item.is_file():
            continue
        if item.name in tracked:
            continue
        artifacts.append({"path": item.name, "size_bytes": item.stat().st_size})
    return {
        "schema_version": 1,
        "generated_at": _now(),
        "artifacts": artifacts,
    }




def _extract_run_dir_from_result(result: Any) -> Path | None:
    if isinstance(result, str):
        marker = "Saved outputs to:"
        idx = result.rfind(marker)
        if idx >= 0:
            candidate = Path(result[idx + len(marker):].strip().splitlines()[0].strip()).expanduser()
            if candidate.exists() and candidate.is_dir():
                return candidate.resolve()
    if isinstance(result, dict):
        artifact_path = str(result.get("artifact_path", "")).strip()
        if artifact_path:
            candidate = Path(artifact_path).expanduser()
            if candidate.exists():
                return candidate.resolve().parent
    return None


def _workflow_summary_kind(job_type: str) -> str:
    normalized = str(job_type).strip().lower()
    if "regime" in normalized:
        return "regime_training"
    if "analysis" in normalized:
        return "analysis"
    return "backtesting"

def _dispatch(job_type: str, payload: dict[str, Any], cancellation_token: CancellationToken) -> Any:
    normalized = str(job_type).strip().lower()

    if normalized in {"optimizer", JOB_BACKTEST_OPTIMIZATION}:
        return run_strategy_optimization(cancellation_token=cancellation_token, **payload)
    if normalized in {"walk_forward", JOB_BACKTEST_WALK_FORWARD}:
        return run_walk_forward_backtest(cancellation_token=cancellation_token, **payload)
    if normalized in {"backtest", JOB_BACKTEST_MULTI_SIGNAL}:
        return run_multi_signal_backtest(cancellation_token=cancellation_token, **payload)
    if normalized == JOB_BACKTEST_TRAINED_REGIME:
        return run_trained_regime_backtest(**payload)
    if normalized == JOB_BACKTEST_TIME_SERIES:
        return run_time_series_momentum_backtest(**payload)
    if normalized == JOB_BACKTEST_CACHE:
        return run_backtest_cache(**payload)
    if normalized in {"regime_training", JOB_REGIME_TRAINING}:
        return execute_regime_training_pipeline(**payload)
    if normalized in {"general_analysis", JOB_ANALYSIS_CALLABLE}:
        fn = payload.get("callable")
        kwargs = payload.get("kwargs", {})
        if not callable(fn):
            raise ValueError("general_analysis payload must include a callable reference")
        return fn(**kwargs)
    raise ValueError(f"Unsupported job_type: {job_type}")


def _run_worker(job_file: Path) -> int:
    envelope = json.loads(job_file.read_text(encoding="utf-8"))
    ensure_schema_compatible(int(envelope.get("schema_version", 1)), source="remote worker envelope")
    job_dir = job_file.parent
    status_path = job_dir / "status.json"
    logs_path = job_dir / "logs.txt"
    summary_path = job_dir / "summary.json"
    artifacts_path = job_dir / "artifacts.json"
    result_path = job_dir / "result.json"
    cancel_path = job_dir / "cancel.requested"
    cancellation_token = CancellationToken()

    started_at = _now()
    _append_log(logs_path, f"Worker starting for {envelope.get('job_type')} ({envelope.get('job_id')})")
    _write_json(status_path, _status_payload(envelope=envelope, status="running", started_at=started_at))
    _write_json(summary_path, _summary_payload(envelope=envelope, status="running", started_at=started_at))
    _append_registry_event(job_dir, {"event": "running", "job_id": envelope.get("job_id") or envelope.get("run_id"), "job_type": envelope.get("job_type")})

    stop_flag = threading.Event()

    def _cancel_watcher() -> None:
        while not stop_flag.is_set():
            if cancel_path.exists():
                cancellation_token.cancel("Cancel marker detected")
                _append_log(logs_path, "Cancellation marker observed; cancellation requested.")
                return
            stop_flag.wait(1.0)

    watcher = threading.Thread(target=_cancel_watcher, daemon=True)
    watcher.start()

    try:
        raw_payload = envelope.get("params", {})
        payload = deserialize_from_json(raw_payload)
        result = _dispatch(str(envelope.get("job_type", "")), payload, cancellation_token)
        run_dir = _extract_run_dir_from_result(result)
        summary_files: list[str] = []
        if run_dir is not None:
            summary_files = build_workflow_summary(run_dir=run_dir, workflow=_workflow_summary_kind(str(envelope.get("job_type", ""))))
        result_payload = {"result": serialize_for_json(result), "summary_files": summary_files, "server_run_dir": str(run_dir) if run_dir is not None else None}
        _write_json(result_path, result_payload)

        completed_at = _now()
        _append_log(logs_path, "Job completed successfully.")
        _write_json(
            status_path,
            _status_payload(
                envelope=envelope,
                status="succeeded",
                started_at=started_at,
                completed_at=completed_at,
            ),
        )
        _write_json(summary_path, _summary_payload(envelope=envelope, status="succeeded", started_at=started_at, completed_at=completed_at))
        _append_registry_event(job_dir, {"event": "completed", "job_id": envelope.get("job_id") or envelope.get("run_id"), "job_type": envelope.get("job_type"), "status": "succeeded"})
        _write_json(artifacts_path, _artifact_manifest(job_dir))
        return 0
    except TaskCancellationError as exc:
        completed_at = _now()
        _append_log(logs_path, f"Job canceled cooperatively: {exc}")
        _write_json(
            status_path,
            _status_payload(
                envelope=envelope,
                status="canceled",
                started_at=started_at,
                completed_at=completed_at,
                error=str(exc),
            ),
        )
        _write_json(summary_path, _summary_payload(envelope=envelope, status="canceled", started_at=started_at, completed_at=completed_at, error=str(exc)))
        _append_registry_event(job_dir, {"event": "completed", "job_id": envelope.get("job_id") or envelope.get("run_id"), "job_type": envelope.get("job_type"), "status": "canceled", "error": str(exc)})
        _write_json(artifacts_path, _artifact_manifest(job_dir))
        return 2
    except Exception as exc:  # noqa: BLE001
        completed_at = _now()
        _append_log(logs_path, f"Job failed: {exc}")
        _append_log(logs_path, traceback.format_exc())
        _write_json(
            status_path,
            _status_payload(
                envelope=envelope,
                status="failed",
                started_at=started_at,
                completed_at=completed_at,
                error=str(exc),
            ),
        )
        _write_json(summary_path, _summary_payload(envelope=envelope, status="failed", started_at=started_at, completed_at=completed_at, error=str(exc)))
        _append_registry_event(job_dir, {"event": "completed", "job_id": envelope.get("job_id") or envelope.get("run_id"), "job_type": envelope.get("job_type"), "status": "failed", "error": str(exc)})
        _write_json(artifacts_path, _artifact_manifest(job_dir))
        return 1
    finally:
        stop_flag.set()
        try:
            tick_scheduler(job_file.parent.parent)
        except Exception:  # noqa: BLE001
            _append_log(logs_path, "Scheduler tick failed during worker finalization.")


def main() -> int:
    parser = argparse.ArgumentParser(description="Execute one stoptions remote job envelope")
    parser.add_argument("--job-file", required=True, help="Path to the job envelope JSON file")
    args = parser.parse_args()
    job_file = Path(args.job_file).expanduser().resolve()
    if not job_file.exists():
        raise FileNotFoundError(f"Job file not found: {job_file}")
    return _run_worker(job_file)


if __name__ == "__main__":
    raise SystemExit(main())

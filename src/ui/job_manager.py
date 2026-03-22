from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
import socket
import threading
import time
from typing import Any, Callable


TERMINAL_JOB_STATES = {"succeeded", "failed", "canceled", "completed"}
TRANSIENT_ERROR_HINTS = (
    "ssh",
    "network",
    "timed out",
    "timeout",
    "connection reset",
    "connection refused",
    "broken pipe",
    "temporar",
    "unreachable",
)


@dataclass
class JobRunResult:
    job_id: str
    status: str
    result: Any = None
    logs: list[str] | None = None
    error_kind: str | None = None
    error_message: str | None = None


class JobManager:
    def __init__(self, *, controller: Any, poll_interval_seconds: float = 0.8, max_retries: int = 4) -> None:
        self.controller = controller
        self._poll_interval_seconds = max(0.2, float(poll_interval_seconds))
        self._max_retries = max(0, int(max_retries))
        self._lock = threading.Lock()
        self._recover_active_jobs()

    @property
    def _active_jobs(self) -> dict[str, dict[str, object]]:
        jobs = getattr(self.controller.state, "active_jobs", None)
        if isinstance(jobs, dict):
            return jobs
        self.controller.state.active_jobs = {}
        return self.controller.state.active_jobs

    def run_job_and_wait(
        self,
        *,
        job_type: str,
        payload: dict[str, Any],
        source_page: str,
        on_update: Callable[[dict[str, object]], None] | None = None,
    ) -> JobRunResult:
        backend = self.controller.execution_backend
        job_id: str | None = None
        submit_attempt = 0
        server_host = self._server_hostname()

        while submit_attempt <= self._max_retries:
            try:
                job_id = backend.submit_job(job_type, payload)
                break
            except Exception as exc:  # noqa: BLE001
                if not self._is_transient_error(exc) or submit_attempt >= self._max_retries:
                    return JobRunResult(
                        job_id=job_id or "",
                        status="failed",
                        error_kind="submission_failure",
                        error_message=self._build_error_message("submission_failure", exc),
                    )
                submit_attempt += 1
                time.sleep(self._backoff_seconds(submit_attempt))

        assert job_id is not None
        metadata = {
            "job_id": job_id,
            "job_type": job_type,
            "source_page": source_page,
            "status": "queued",
            "submitted_at": self._iso_now(),
            "started_at": None,
            "ended_at": None,
            "server_hostname": server_host,
            "artifact_sync_status": "not_started",
            "poll_interval_seconds": self._poll_interval_seconds,
            "transport_retries": 0,
            "max_transport_retries": self._max_retries,
            "retryable_transport_failure": False,
            "error_kind": None,
            "error_message": None,
        }
        self._store_metadata(job_id, metadata)
        self._emit_update(on_update, metadata)

        status = "queued"
        transport_retries = 0
        while True:
            try:
                status = str(backend.get_status(job_id)).strip().lower() or "queued"
            except Exception as exc:  # noqa: BLE001
                if self._is_transient_error(exc) and transport_retries < self._max_retries:
                    transport_retries += 1
                    metadata["transport_retries"] = transport_retries
                    metadata["retryable_transport_failure"] = True
                    metadata["error_kind"] = "transport_failure"
                    metadata["error_message"] = self._build_error_message("transport_failure", exc)
                    self._store_metadata(job_id, metadata)
                    self._emit_update(on_update, metadata)
                    time.sleep(self._backoff_seconds(transport_retries))
                    continue
                metadata["status"] = "failed"
                metadata["ended_at"] = self._iso_now()
                metadata["error_kind"] = "transport_failure"
                metadata["error_message"] = self._build_error_message("transport_failure", exc)
                metadata["retryable_transport_failure"] = True
                self._store_metadata(job_id, metadata)
                self._emit_update(on_update, metadata)
                return JobRunResult(
                    job_id=job_id,
                    status="failed",
                    error_kind="transport_failure",
                    error_message=str(metadata["error_message"]),
                )

            metadata["status"] = status
            if status in {"running"} and metadata.get("started_at") is None:
                metadata["started_at"] = self._iso_now()
            if status in TERMINAL_JOB_STATES:
                metadata["ended_at"] = self._iso_now()
            self._store_metadata(job_id, metadata)
            self._emit_update(on_update, metadata)

            if status in TERMINAL_JOB_STATES:
                break
            time.sleep(self._poll_interval_seconds)

        logs = backend.stream_logs(job_id) if hasattr(backend, "stream_logs") else []
        if status in {"failed", "canceled"}:
            error_message = logs[-1] if logs else f"{job_type} failed"
            metadata["error_kind"] = "remote_runtime_failure"
            metadata["error_message"] = self._build_error_message("remote_runtime_failure", RuntimeError(error_message))
            self._store_metadata(job_id, metadata)
            self._emit_update(on_update, metadata)
            return JobRunResult(
                job_id=job_id,
                status=status,
                logs=logs,
                error_kind="remote_runtime_failure",
                error_message=str(metadata["error_message"]),
            )

        result = backend.get_result(job_id) if hasattr(backend, "get_result") else None
        return JobRunResult(job_id=job_id, status=status, result=result, logs=logs)

    def mark_artifact_sync(self, job_id: str, *, status: str, error: Exception | None = None) -> None:
        with self._lock:
            metadata = dict(self._active_jobs.get(job_id, {}))
            if not metadata:
                return
            metadata["artifact_sync_status"] = status
            if error is not None:
                metadata["error_kind"] = "artifact_sync_failure"
                metadata["error_message"] = self._build_error_message("artifact_sync_failure", error)
            self._active_jobs[job_id] = metadata
        self.controller.persist_state()

    def retry_transport_failure(self, job_id: str) -> bool:
        backend = self.controller.execution_backend
        with self._lock:
            metadata = dict(self._active_jobs.get(job_id, {}))
        if not metadata or not metadata.get("retryable_transport_failure"):
            return False
        if str(metadata.get("status", "")).lower() in TERMINAL_JOB_STATES:
            # Recover terminal state metadata without rerunning the job itself.
            try:
                status = str(backend.get_status(job_id)).strip().lower()
            except Exception:
                return False
            metadata["status"] = status or str(metadata.get("status", "failed"))
            metadata["retryable_transport_failure"] = False
            metadata["transport_retries"] = 0
            with self._lock:
                self._active_jobs[job_id] = metadata
            self.controller.persist_state()
            return True
        return False

    def _recover_active_jobs(self) -> None:
        backend = self.controller.execution_backend
        for job_id, payload in list(self._active_jobs.items()):
            if not isinstance(payload, dict):
                continue
            status = str(payload.get("status", "")).lower()
            if status in TERMINAL_JOB_STATES:
                continue
            if hasattr(backend, "register_existing_job"):
                try:
                    backend.register_existing_job(job_id=job_id, job_type=str(payload.get("job_type", "unknown")))
                except Exception:
                    payload["retryable_transport_failure"] = True
                    payload["error_kind"] = "transport_failure"
                    payload["error_message"] = "Unable to recover remote job handle; use retry transport action."
            self._active_jobs[job_id] = payload
        self.controller.persist_state()

    def _store_metadata(self, job_id: str, metadata: dict[str, object]) -> None:
        with self._lock:
            self._active_jobs[job_id] = dict(metadata)
        self.controller.persist_state()

    def _emit_update(self, on_update: Callable[[dict[str, object]], None] | None, metadata: dict[str, object]) -> None:
        if on_update is None:
            return
        on_update(dict(metadata))

    def _server_hostname(self) -> str:
        settings = getattr(self.controller.state, "remote_execution_settings", {})
        mode = str(settings.get("mode", "local")).strip().lower() if isinstance(settings, dict) else "local"
        if mode == "remote":
            return str(settings.get("ssh_host", "remote-host")).strip() or "remote-host"
        return socket.gethostname()

    @staticmethod
    def _iso_now() -> str:
        return datetime.now(timezone.utc).isoformat()

    def _backoff_seconds(self, retry_index: int) -> float:
        base = self._poll_interval_seconds
        return min(12.0, base * (2 ** max(0, retry_index - 1)))

    @staticmethod
    def _is_transient_error(exc: Exception) -> bool:
        message = str(exc).lower()
        return any(hint in message for hint in TRANSIENT_ERROR_HINTS)

    @staticmethod
    def _build_error_message(error_kind: str, exc: Exception) -> str:
        labels = {
            "submission_failure": "Submission failure",
            "transport_failure": "Transport failure",
            "remote_runtime_failure": "Remote runtime failure",
            "artifact_sync_failure": "Artifact sync failure",
        }
        return f"{labels.get(error_kind, 'Job failure')}: {exc}"

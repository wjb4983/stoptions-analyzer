from __future__ import annotations

from dataclasses import dataclass
import socket
import threading
import time
from typing import Any, Callable

from execution.contracts import (
    CancelJobRequest,
    CancelJobResponse,
    JobState,
    JobStatusResponse,
    JobSummaryResponse,
    SubmitJobRequest,
    SubmitJobResponse,
    ensure_schema_compatible,
    normalize_job_state,
    now_utc_iso,
)


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
    summary: JobSummaryResponse
    result: Any = None
    logs: list[str] | None = None
    error_kind: str | None = None
    error_message: str | None = None

    @property
    def job_id(self) -> str:
        return self.summary.job_id

    @property
    def status(self) -> str:
        return self.summary.state.value


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
        request: SubmitJobRequest,
        source_page: str,
        on_update: Callable[[dict[str, object]], None] | None = None,
    ) -> JobRunResult:
        backend = self.controller.execution_backend
        ensure_schema_compatible(request.schema_version, source="client request")
        job_id: str | None = None
        submit_attempt = 0
        server_host = self._server_hostname()
        submitted_at = now_utc_iso()
        submit_response: SubmitJobResponse | None = None

        while submit_attempt <= self._max_retries:
            try:
                job_id = backend.submit_job(request.job_type, request.payload)
                submit_response = SubmitJobResponse(
                    job_id=job_id,
                    job_type=request.job_type,
                    state=JobState.QUEUED,
                    submitted_at=submitted_at,
                    schema_version=request.schema_version,
                )
                break
            except Exception as exc:  # noqa: BLE001
                if not self._is_transient_error(exc) or submit_attempt >= self._max_retries:
                    return JobRunResult(
                        summary=JobSummaryResponse(
                            job_id=job_id or "",
                            job_type=request.job_type,
                            state=JobState.FAILED,
                            submitted_at=submitted_at,
                            started_at=None,
                            ended_at=now_utc_iso(),
                            server_run_dir=None,
                            summary_payload={},
                            summary_paths={},
                            schema_version=request.schema_version,
                        ),
                        error_kind="submission_failure",
                        error_message=self._build_error_message("submission_failure", exc),
                    )
                submit_attempt += 1
                time.sleep(self._backoff_seconds(submit_attempt))

        assert job_id is not None
        metadata = {
            "job_id": job_id,
            "job_type": request.job_type,
            "source_page": source_page,
            "status": submit_response.state.value if submit_response else "queued",
            "submitted_at": submit_response.submitted_at if submit_response else submitted_at,
            "started_at": None,
            "ended_at": None,
            "server_hostname": server_host,
            "server_run_dir": None,
            "summary_paths": {},
            "summary_payload": {},
            "schema_version": request.schema_version,
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

        status = JobState.QUEUED.value
        transport_retries = 0
        while True:
            try:
                backend_status = str(backend.get_status(job_id)).strip().lower() or JobState.QUEUED.value
                status = normalize_job_state(backend_status).value
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
                metadata["ended_at"] = now_utc_iso()
                metadata["error_kind"] = "transport_failure"
                metadata["error_message"] = self._build_error_message("transport_failure", exc)
                metadata["retryable_transport_failure"] = True
                self._store_metadata(job_id, metadata)
                self._emit_update(on_update, metadata)
                return JobRunResult(
                    summary=self._build_summary(metadata),
                    error_kind="transport_failure",
                    error_message=str(metadata["error_message"]),
                )

            metadata["status"] = status
            if status == JobState.RUNNING.value and metadata.get("started_at") is None:
                metadata["started_at"] = now_utc_iso()
            if status in TERMINAL_JOB_STATES:
                metadata["ended_at"] = now_utc_iso()
            self._store_metadata(job_id, metadata)
            self._emit_update(on_update, metadata)

            if status in TERMINAL_JOB_STATES:
                break
            time.sleep(self._poll_interval_seconds)

        logs = backend.stream_logs(job_id) if hasattr(backend, "stream_logs") else []
        metadata["summary_payload"] = {
            "log_line_count": len(logs),
            "error_kind": metadata.get("error_kind"),
        }
        self._store_metadata(job_id, metadata)
        if status in {"failed", "canceled"}:
            error_message = logs[-1] if logs else f"{request.job_type} failed"
            metadata["error_kind"] = "remote_runtime_failure"
            metadata["error_message"] = self._build_error_message("remote_runtime_failure", RuntimeError(error_message))
            self._store_metadata(job_id, metadata)
            self._emit_update(on_update, metadata)
            return JobRunResult(
                summary=self._build_summary(metadata),
                logs=logs,
                error_kind="remote_runtime_failure",
                error_message=str(metadata["error_message"]),
            )

        result = backend.get_result(job_id) if hasattr(backend, "get_result") else None
        if isinstance(result, dict):
            metadata["summary_payload"] = {**metadata.get("summary_payload", {}), **result}
        self._store_metadata(job_id, metadata)
        return JobRunResult(summary=self._build_summary(metadata), result=result, logs=logs)

    def cancel_job(self, request: CancelJobRequest) -> CancelJobResponse:
        ensure_schema_compatible(request.schema_version, source="cancel request")
        backend = self.controller.execution_backend
        backend.cancel_job(request.job_id)
        return CancelJobResponse(
            job_id=request.job_id,
            state=JobState.CANCELED,
            canceled_at=now_utc_iso(),
            schema_version=request.schema_version,
        )

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
    def _build_summary(metadata: dict[str, object]) -> JobSummaryResponse:
        ensure_schema_compatible(int(metadata.get("schema_version", 1)), source="job metadata")
        status = str(metadata.get("status", JobState.QUEUED.value))
        status_response = JobStatusResponse(
            job_id=str(metadata.get("job_id", "")),
            job_type=str(metadata.get("job_type", "")),
            state=normalize_job_state(status),
            submitted_at=str(metadata.get("submitted_at")) if metadata.get("submitted_at") is not None else None,
            started_at=str(metadata.get("started_at")) if metadata.get("started_at") is not None else None,
            ended_at=str(metadata.get("ended_at")) if metadata.get("ended_at") is not None else None,
            server_run_dir=str(metadata.get("server_run_dir")) if metadata.get("server_run_dir") else None,
            schema_version=int(metadata.get("schema_version", 1)),
        )
        return JobSummaryResponse(
            job_id=status_response.job_id,
            job_type=status_response.job_type,
            state=status_response.state,
            submitted_at=status_response.submitted_at,
            started_at=status_response.started_at,
            ended_at=status_response.ended_at,
            server_run_dir=status_response.server_run_dir,
            summary_paths=dict(metadata.get("summary_paths", {})) if isinstance(metadata.get("summary_paths"), dict) else {},
            summary_payload=dict(metadata.get("summary_payload", {})) if isinstance(metadata.get("summary_payload"), dict) else {},
            schema_version=status_response.schema_version,
        )

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

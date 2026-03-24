from __future__ import annotations

from dataclasses import dataclass, field
import json
from pathlib import Path
import threading
import time
from typing import Any, Callable
from uuid import uuid4

from backtesting.cache_runner import (
    run_backtest_cache,
    run_multi_signal_backtest,
    run_strategy_optimization,
    run_time_series_momentum_backtest,
    run_trained_regime_backtest,
    run_walk_forward_backtest,
)
from backtesting.regime_training_pipeline import execute_regime_training_pipeline

JOB_BACKTEST_OPTIMIZATION = "backtesting.optimization"
JOB_BACKTEST_WALK_FORWARD = "backtesting.walk_forward"
JOB_BACKTEST_MULTI_SIGNAL = "backtesting.multi_signal"
JOB_BACKTEST_TRAINED_REGIME = "backtesting.trained_regime"
JOB_BACKTEST_TIME_SERIES = "backtesting.time_series_momentum"
JOB_BACKTEST_CACHE = "backtesting.cache"
JOB_REGIME_TRAINING = "regime.training"
JOB_ANALYSIS_CALLABLE = "analysis.callable"

# Remote envelope-friendly aliases
JOB_BACKTEST = "backtest"
JOB_WALK_FORWARD = "walk_forward"
JOB_OPTIMIZER = "optimizer"
JOB_REGIME_TRAINING_ALIAS = "regime_training"
JOB_GENERAL_ANALYSIS = "general_analysis"


class ExecutionBackend:
    def submit_job(self, job_type: str, payload: dict[str, Any]) -> str:
        raise NotImplementedError

    def get_status(self, job_id: str) -> str:
        raise NotImplementedError

    def stream_logs(self, job_id: str) -> list[str]:
        raise NotImplementedError

    def fetch_artifacts(
        self,
        job_id: str,
        target_dir: str | Path,
        *,
        fetch_mode: str = "summary_only",
        selected_files: list[str] | None = None,
        allow_full_artifacts: bool = False,
    ) -> Path:
        raise NotImplementedError

    def cancel_job(self, job_id: str) -> None:
        raise NotImplementedError


@dataclass
class _JobRecord:
    job_id: str
    job_type: str
    payload: dict[str, Any]
    status: str = "queued"
    logs: list[str] = field(default_factory=list)
    result: Any = None
    error: str | None = None
    started_at: float | None = None
    completed_at: float | None = None


class LocalExecutionBackend(ExecutionBackend):
    def __init__(self) -> None:
        self._jobs: dict[str, _JobRecord] = {}
        self._lock = threading.Lock()

    def submit_job(self, job_type: str, payload: dict[str, Any]) -> str:
        job_id = uuid4().hex
        record = _JobRecord(job_id=job_id, job_type=job_type, payload=dict(payload))
        with self._lock:
            self._jobs[job_id] = record

        thread = threading.Thread(target=self._run_job, args=(job_id,), daemon=True)
        thread.start()
        return job_id

    def get_status(self, job_id: str) -> str:
        return self._get_job(job_id).status

    def stream_logs(self, job_id: str) -> list[str]:
        return list(self._get_job(job_id).logs)

    def fetch_artifacts(
        self,
        job_id: str,
        target_dir: str | Path,
        *,
        fetch_mode: str = "summary_only",
        selected_files: list[str] | None = None,
        allow_full_artifacts: bool = False,
    ) -> Path:
        record = self._get_job(job_id)
        artifact_dir = Path(target_dir).expanduser() / job_id
        artifact_dir.mkdir(parents=True, exist_ok=True)

        (artifact_dir / "logs.txt").write_text("\n".join(record.logs), encoding="utf-8")
        summary = {
            "job_id": record.job_id,
            "job_type": record.job_type,
            "status": record.status,
            "error": record.error,
            "started_at": record.started_at,
            "completed_at": record.completed_at,
            "fetch_mode": fetch_mode,
            "selected_files": list(selected_files or []),
            "allow_full_artifacts": bool(allow_full_artifacts),
        }
        (artifact_dir / "summary.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")

        if isinstance(record.result, str):
            (artifact_dir / "result.txt").write_text(record.result, encoding="utf-8")
        elif record.result is not None:
            (artifact_dir / "result.json").write_text(json.dumps(record.result, indent=2, default=str), encoding="utf-8")
        return artifact_dir

    def get_result(self, job_id: str) -> Any:
        return self._get_job(job_id).result

    def get_error(self, job_id: str) -> str | None:
        return self._get_job(job_id).error

    def cancel_job(self, job_id: str) -> None:
        record = self._get_job(job_id)
        if record.status in {"succeeded", "failed", "canceled"}:
            return
        record.status = "canceled"
        record.error = "Canceled by user"
        record.logs.append("Cancellation requested")
        record.completed_at = time.time()

    def _get_job(self, job_id: str) -> _JobRecord:
        with self._lock:
            record = self._jobs.get(job_id)
        if record is None:
            raise KeyError(f"Unknown job_id: {job_id}")
        return record

    def _run_job(self, job_id: str) -> None:
        record = self._get_job(job_id)
        if record.status == "canceled":
            return
        record.status = "running"
        record.started_at = time.time()
        record.logs.append(f"Started {record.job_type}")

        try:
            record.result = self._dispatch(record.job_type, record.payload)
            record.status = "succeeded"
            record.logs.append("Completed successfully")
        except Exception as exc:
            record.status = "failed"
            record.error = str(exc)
            record.logs.append(f"Failed: {exc}")
        finally:
            record.completed_at = time.time()

    def _dispatch(self, job_type: str, payload: dict[str, Any]) -> Any:
        if job_type == JOB_BACKTEST_OPTIMIZATION:
            return run_strategy_optimization(**payload)
        if job_type == JOB_BACKTEST_WALK_FORWARD:
            return run_walk_forward_backtest(**payload)
        if job_type == JOB_BACKTEST_MULTI_SIGNAL:
            return run_multi_signal_backtest(**payload)
        if job_type == JOB_BACKTEST_TRAINED_REGIME:
            return run_trained_regime_backtest(**payload)
        if job_type == JOB_BACKTEST_TIME_SERIES:
            return run_time_series_momentum_backtest(**payload)
        if job_type == JOB_BACKTEST_CACHE:
            return run_backtest_cache(**payload)
        if job_type == JOB_REGIME_TRAINING:
            return execute_regime_training_pipeline(**payload)
        if job_type == JOB_ANALYSIS_CALLABLE:
            fn = payload.get("callable")
            kwargs = payload.get("kwargs", {})
            if not callable(fn):
                raise ValueError("analysis.callable payload must include a callable")
            return fn(**kwargs)
        raise ValueError(f"Unsupported job_type: {job_type}")


def build_execution_backend(*, mode: str = "local", remote_settings: dict[str, object] | None = None) -> ExecutionBackend:
    normalized = str(mode).strip().lower()
    if normalized in {"", "local"}:
        return LocalExecutionBackend()
    if normalized in {"remote", "ssh", "remote_ssh", "ssh_remote"}:
        from .remote_ssh_backend import build_remote_backend_from_settings

        return build_remote_backend_from_settings(remote_settings or {})
    raise ValueError(f"Unsupported execution backend mode: {mode}")

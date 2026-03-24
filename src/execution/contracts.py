from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from typing import Any

try:
    from enum import StrEnum
except ImportError:  # pragma: no cover - Python < 3.11 compatibility
    class StrEnum(str, Enum):
        pass

SCHEMA_VERSION = 1
MIN_SUPPORTED_SCHEMA_VERSION = 1
MAX_SUPPORTED_SCHEMA_VERSION = 1


class JobType(StrEnum):
    BACKTEST = "backtest"
    GPU_TRAINING = "gpu_training"
    GENERAL_ANALYSIS = "general_analysis"
    REGIME_TRAINING = "regime_training"


class JobState(StrEnum):
    QUEUED = "queued"
    RUNNING = "running"
    COMPLETED = "completed"
    FAILED = "failed"
    CANCELED = "canceled"


_BACKEND_STATUS_MAP: dict[str, JobState] = {
    "queued": JobState.QUEUED,
    "running": JobState.RUNNING,
    "succeeded": JobState.COMPLETED,
    "completed": JobState.COMPLETED,
    "failed": JobState.FAILED,
    "canceled": JobState.CANCELED,
    "canceling": JobState.RUNNING,
}


@dataclass(frozen=True)
class SubmitJobRequest:
    job_type: str
    payload: dict[str, Any]
    schema_version: int = SCHEMA_VERSION


@dataclass(frozen=True)
class SubmitJobResponse:
    job_id: str
    job_type: str
    state: JobState
    submitted_at: str
    schema_version: int = SCHEMA_VERSION


@dataclass(frozen=True)
class JobStatusResponse:
    job_id: str
    job_type: str
    state: JobState
    submitted_at: str | None = None
    started_at: str | None = None
    ended_at: str | None = None
    server_run_dir: str | None = None
    schema_version: int = SCHEMA_VERSION


@dataclass(frozen=True)
class JobSummaryResponse:
    job_id: str
    job_type: str
    state: JobState
    submitted_at: str | None
    started_at: str | None
    ended_at: str | None
    server_run_dir: str | None
    summary_paths: dict[str, str] = field(default_factory=dict)
    summary_payload: dict[str, Any] = field(default_factory=dict)
    schema_version: int = SCHEMA_VERSION


@dataclass(frozen=True)
class CancelJobRequest:
    job_id: str
    schema_version: int = SCHEMA_VERSION


@dataclass(frozen=True)
class CancelJobResponse:
    job_id: str
    state: JobState
    canceled_at: str
    schema_version: int = SCHEMA_VERSION


class ContractVersionError(ValueError):
    pass


def ensure_schema_compatible(schema_version: int, *, source: str) -> None:
    version = int(schema_version)
    if MIN_SUPPORTED_SCHEMA_VERSION <= version <= MAX_SUPPORTED_SCHEMA_VERSION:
        return
    raise ContractVersionError(
        f"Schema version mismatch ({source}): got {version}, "
        f"supported={MIN_SUPPORTED_SCHEMA_VERSION}-{MAX_SUPPORTED_SCHEMA_VERSION}"
    )


def now_utc_iso() -> str:
    return datetime.utcnow().isoformat() + "Z"


def normalize_job_state(status: str) -> JobState:
    normalized = str(status).strip().lower()
    if normalized in _BACKEND_STATUS_MAP:
        return _BACKEND_STATUS_MAP[normalized]
    raise ValueError(f"Unknown job state: {status}")

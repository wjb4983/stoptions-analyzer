from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import shlex
import tempfile
import time
from typing import Any
from uuid import uuid4

from .backend import ExecutionBackend
from .contracts import SCHEMA_VERSION, ensure_schema_compatible, normalize_job_state
from .remote_payloads import deserialize_from_json, serialize_for_json
from .ssh_transport import SSHTransport, SSHTransportConfig


DEFAULT_REMOTE_ROOT = "~/stoptions_jobs"
REMOTE_REGISTRY_FILENAME = "registry.jsonl"
_STATUS_POLL_SECONDS = 1.5


@dataclass
class _RemoteJobRecord:
    job_id: str
    job_type: str
    local_dir: Path
    remote_dir: str
    status: str = "queued"
    last_status_at: float = 0.0
    status_payload: dict[str, Any] = field(default_factory=dict)


class RemoteSSHExecutionBackend(ExecutionBackend):
    def __init__(
        self,
        *,
        host: str,
        user: str | None = None,
        port: int | None = None,
        python_bin: str = "python",
        remote_root: str = DEFAULT_REMOTE_ROOT,
        ssh_options: str = "",
        ssh_identity_file: str | None = None,
        ssh_known_hosts_file: str | None = None,
        strict_host_key_checking: bool = True,
        poll_interval_seconds: float = _STATUS_POLL_SECONDS,
    ) -> None:
        self._host = host.strip()
        if not self._host:
            raise ValueError("RemoteSSHExecutionBackend requires a host")
        self._user = (user or "").strip() or None
        self._port = int(port) if port else None
        self._python_bin = str(python_bin).strip() or "python"
        self._remote_root = str(remote_root).strip() or DEFAULT_REMOTE_ROOT
        self._ssh_options = str(ssh_options).strip()
        self._poll_interval_seconds = max(0.2, float(poll_interval_seconds))
        self._jobs: dict[str, _RemoteJobRecord] = {}
        self._local_root = Path(tempfile.gettempdir()) / "stoptions_remote_jobs"
        self._local_root.mkdir(parents=True, exist_ok=True)
        self._transport = SSHTransport(
            SSHTransportConfig(
                host=self._host,
                user=self._user,
                port=self._port,
                identity_file=(ssh_identity_file or "").strip() or None,
                known_hosts_file=(ssh_known_hosts_file or "").strip() or None,
                strict_host_key_checking=bool(strict_host_key_checking),
                extra_options=self._ssh_options,
            )
        )

    def validate_connection(self) -> tuple[bool, str]:
        try:
            output = self._transport.run(f"{shlex.quote(self._python_bin)} --version")
        except Exception as exc:  # noqa: BLE001
            return False, str(exc)
        return True, output.strip() or "Remote python is reachable."

    def submit_job(self, job_type: str, payload: dict[str, Any]) -> str:
        job_id = uuid4().hex
        created_at = datetime.now(timezone.utc).isoformat()
        remote_dir = f"{self._remote_root.rstrip('/')}/{job_id}"
        local_dir = self._local_root / job_id
        local_dir.mkdir(parents=True, exist_ok=True)

        envelope = {
            "schema_version": SCHEMA_VERSION,
            "run_id": job_id,
            "job_id": job_id,
            "job_type": str(job_type),
            "params": serialize_for_json(payload),
            "requested_outputs": ["status.json", "logs.txt", "summary.json", "result.json"],
            "timestamps": {
                "created_at": created_at,
                "submitted_at": created_at,
            },
        }
        envelope_text = json.dumps(envelope, indent=2)
        (local_dir / "job_request.json").write_text(envelope_text, encoding="utf-8")

        escaped_remote_dir = shlex.quote(remote_dir)
        self._transport.run(f"mkdir -p {escaped_remote_dir}")
        self._transport.write_text(f"{remote_dir}/job_request.json", envelope_text)
        self._transport.write_text(
            f"{remote_dir}/status.json",
            json.dumps(
                {
                    "schema_version": SCHEMA_VERSION,
                    "job_id": job_id,
                    "job_type": str(job_type),
                    "status": "queued",
                    "timestamps": {"created_at": created_at, "submitted_at": created_at, "started_at": None, "completed_at": None},
                    "error": None,
                },
                indent=2,
            ),
        )
        self._transport.write_text(f"{remote_dir}/logs.txt", "")
        self._transport.write_text(
            f"{remote_dir}/summary.json",
            json.dumps(
                {
                    "schema_version": SCHEMA_VERSION,
                    "job_id": job_id,
                    "job_type": str(job_type),
                    "status": "queued",
                    "timestamps": {"created_at": created_at, "submitted_at": created_at, "started_at": None, "completed_at": None},
                },
                indent=2,
            ),
        )

        launched_at = datetime.now(timezone.utc).isoformat()
        launch_output = self._transport.run(
            f"cd {escaped_remote_dir} && nohup {shlex.quote(self._python_bin)} -m remote.worker "
            f"--job-file {escaped_remote_dir}/job_request.json "
            ">> logs.txt 2>&1 & echo $!"
        ).strip()
        child_pid = int(launch_output.splitlines()[-1]) if launch_output else None
        launch_metadata = {
            "schema_version": SCHEMA_VERSION,
            "job_id": job_id,
            "job_type": str(job_type),
            "started_at": launched_at,
            "child_pid": child_pid,
            "remote_dir": remote_dir,
        }
        self._transport.write_text(f"{remote_dir}/launch_metadata.json", json.dumps(launch_metadata, indent=2))
        self._append_registry_entry(
            {
                "event": "submitted",
                "job_id": job_id,
                "job_type": str(job_type),
                "remote_dir": remote_dir,
                "submitted_at": created_at,
                "started_at": launched_at,
                "child_pid": child_pid,
            }
        )

        self._jobs[job_id] = _RemoteJobRecord(
            job_id=job_id,
            job_type=str(job_type),
            local_dir=local_dir,
            remote_dir=remote_dir,
            status="queued",
        )
        return job_id

    def get_status(self, job_id: str) -> str:
        record = self._get_record(job_id)
        now = time.monotonic()
        if record.status in {"succeeded", "failed", "canceled"}:
            return record.status
        if now - record.last_status_at < self._poll_interval_seconds:
            return record.status

        status_json = self._transport.read_text(f"{record.remote_dir}/status.json")
        record.last_status_at = now
        if status_json is None:
            return record.status
        try:
            payload = json.loads(status_json)
        except json.JSONDecodeError:
            return record.status
        record.status_payload = payload
        ensure_schema_compatible(int(payload.get("schema_version", 1)), source="remote status payload")
        record.status = normalize_job_state(str(payload.get("status", record.status))).value
        return record.status

    def stream_logs(self, job_id: str) -> list[str]:
        record = self._get_record(job_id)
        text = self._transport.read_text(f"{record.remote_dir}/logs.txt")
        if text is None:
            return []
        return text.splitlines()

    def fetch_artifacts(self, job_id: str, target_dir: str | Path) -> Path:
        record = self._get_record(job_id)
        target_root = Path(target_dir).expanduser() / job_id
        target_root.mkdir(parents=True, exist_ok=True)

        for fixed_name in ("job_request.json", "status.json", "logs.txt", "summary.json", "artifacts.json", "result.json", "launch_metadata.json"):
            self._transport.download_file(f"{record.remote_dir}/{fixed_name}", target_root / fixed_name)

        manifest_path = target_root / "artifacts.json"
        if manifest_path.exists():
            try:
                manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                manifest = {}
            for artifact in manifest.get("artifacts", []):
                rel_path = str(artifact.get("path", "")).strip()
                if not rel_path:
                    continue
                destination = target_root / rel_path
                destination.parent.mkdir(parents=True, exist_ok=True)
                self._transport.download_file(f"{record.remote_dir}/{rel_path}", destination)
        return target_root

    def cancel_job(self, job_id: str) -> None:
        record = self._get_record(job_id)
        self._transport.run(f"mkdir -p {shlex.quote(record.remote_dir)} && touch {shlex.quote(record.remote_dir + '/cancel.requested')}")
        record.status = "canceling"

    def register_existing_job(self, *, job_id: str, job_type: str = "unknown") -> None:
        cleaned = str(job_id).strip()
        if not cleaned or cleaned in self._jobs:
            return
        remote_dir = f"{self._remote_root.rstrip('/')}/{cleaned}"
        local_dir = self._local_root / cleaned
        local_dir.mkdir(parents=True, exist_ok=True)
        self._jobs[cleaned] = _RemoteJobRecord(
            job_id=cleaned,
            job_type=str(job_type or "unknown"),
            local_dir=local_dir,
            remote_dir=remote_dir,
            status="queued",
        )

    def get_result(self, job_id: str) -> Any:
        record = self._get_record(job_id)
        result_text = self._transport.read_text(f"{record.remote_dir}/result.json")
        if not result_text:
            return None
        payload = json.loads(result_text)
        return deserialize_from_json(payload.get("result"))

    def _get_record(self, job_id: str) -> _RemoteJobRecord:
        record = self._jobs.get(job_id)
        if record is None:
            raise KeyError(f"Unknown job_id: {job_id}")
        return record

    def _append_registry_entry(self, payload: dict[str, Any]) -> None:
        entry = json.dumps({"timestamp": datetime.now(timezone.utc).isoformat(), **payload}, default=str)
        remote_registry = f"{self._remote_root.rstrip('/')}/{REMOTE_REGISTRY_FILENAME}"
        command = (
            f"mkdir -p {shlex.quote(self._remote_root)} && "
            f"cat >> {shlex.quote(remote_registry)} <<'JSON'\n{entry}\nJSON"
        )
        self._transport.run(command)


def build_remote_backend_from_settings(settings: dict[str, object]) -> RemoteSSHExecutionBackend:
    host = str(settings.get("ssh_host", "")).strip()
    if not host:
        host = os.getenv("STOPTIONS_REMOTE_HOST", "").strip()
    if not host:
        raise ValueError("Remote host is required for remote backend mode")
    user = str(settings.get("ssh_user", "")).strip() or os.getenv("STOPTIONS_REMOTE_USER", "").strip() or None
    port_raw = str(settings.get("ssh_port", "")).strip() or os.getenv("STOPTIONS_REMOTE_PORT", "").strip()
    port = int(port_raw) if port_raw else None
    remote_venv_path = str(settings.get("remote_venv_path", "")).strip()
    python_from_venv = f"{remote_venv_path.rstrip('/')}/bin/python" if remote_venv_path else ""
    python_bin = (
        str(settings.get("remote_python_command", "")).strip()
        or python_from_venv
        or os.getenv("STOPTIONS_REMOTE_PYTHON", "python")
    )
    remote_root = str(settings.get("remote_project_path", "")).strip() or os.getenv("STOPTIONS_REMOTE_ROOT", DEFAULT_REMOTE_ROOT)
    ssh_options = str(settings.get("ssh_options", "")).strip() or os.getenv("STOPTIONS_REMOTE_SSH_OPTIONS", "")
    identity_file = str(settings.get("ssh_identity_file", "")).strip() or os.getenv("STOPTIONS_REMOTE_IDENTITY_FILE", "").strip() or None
    known_hosts_file = str(settings.get("ssh_known_hosts_file", "")).strip() or os.getenv("STOPTIONS_REMOTE_KNOWN_HOSTS_FILE", "").strip() or None
    strict_host_key_checking = str(settings.get("ssh_strict_host_key_checking", "true")).strip().lower() not in {"0", "false", "no", "off"}
    poll_raw = str(settings.get("scheduler_poll_seconds", "")).strip() or os.getenv("STOPTIONS_REMOTE_POLL_SECONDS", str(_STATUS_POLL_SECONDS))
    poll_interval = float(poll_raw)
    return RemoteSSHExecutionBackend(
        host=host,
        user=user,
        port=port,
        python_bin=python_bin,
        remote_root=remote_root,
        ssh_options=ssh_options,
        ssh_identity_file=identity_file,
        ssh_known_hosts_file=known_hosts_file,
        strict_host_key_checking=strict_host_key_checking,
        poll_interval_seconds=poll_interval,
    )

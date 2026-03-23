from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import shlex
import subprocess
import tempfile
import time
from typing import Any
from uuid import uuid4

from .backend import ExecutionBackend
from .contracts import SCHEMA_VERSION, ensure_schema_compatible, normalize_job_state
from .remote_payloads import deserialize_from_json, serialize_for_json


DEFAULT_REMOTE_ROOT = "~/stoptions_jobs"
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

    def validate_connection(self) -> tuple[bool, str]:
        try:
            output = self._run_ssh(f"{shlex.quote(self._python_bin)} --version")
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
            "requested_outputs": ["status.json", "logs.txt", "artifacts.json", "result.json"],
            "timestamps": {
                "created_at": created_at,
                "submitted_at": created_at,
            },
        }
        envelope_text = json.dumps(envelope, indent=2)
        (local_dir / "job.json").write_text(envelope_text, encoding="utf-8")

        escaped_remote_dir = shlex.quote(remote_dir)
        launch_cmd = (
            f"mkdir -p {escaped_remote_dir} && "
            f"cat > {escaped_remote_dir}/job.json <<'JSON'\n{envelope_text}\nJSON\n"
            f"nohup {shlex.quote(self._python_bin)} -m remote.worker --job-file {escaped_remote_dir}/job.json "
            f">> {escaped_remote_dir}/launcher.log 2>&1 & echo $! > {escaped_remote_dir}/worker.pid"
        )
        self._run_ssh(launch_cmd)

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

        status_json = self._read_remote_text(record.remote_dir, "status.json")
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
        text = self._read_remote_text(record.remote_dir, "logs.txt")
        if text is None:
            return []
        return text.splitlines()

    def fetch_artifacts(self, job_id: str, target_dir: str | Path) -> Path:
        record = self._get_record(job_id)
        target_root = Path(target_dir).expanduser() / job_id
        target_root.mkdir(parents=True, exist_ok=True)

        for fixed_name in ("job.json", "status.json", "logs.txt", "artifacts.json", "result.json"):
            self._download_remote_file(record.remote_dir, fixed_name, target_root / fixed_name)

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
                self._download_remote_file(record.remote_dir, rel_path, destination)
        return target_root

    def cancel_job(self, job_id: str) -> None:
        record = self._get_record(job_id)
        self._run_ssh(f"mkdir -p {shlex.quote(record.remote_dir)} && touch {shlex.quote(record.remote_dir + '/cancel.requested')}")
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
        result_text = self._read_remote_text(record.remote_dir, "result.json")
        if not result_text:
            return None
        payload = json.loads(result_text)
        return deserialize_from_json(payload.get("result"))

    def _get_record(self, job_id: str) -> _RemoteJobRecord:
        record = self._jobs.get(job_id)
        if record is None:
            raise KeyError(f"Unknown job_id: {job_id}")
        return record

    def _build_ssh_target(self) -> str:
        if self._user:
            return f"{self._user}@{self._host}"
        return self._host

    def _build_ssh_args(self, remote_command: str) -> list[str]:
        args = ["ssh"]
        if self._port:
            args.extend(["-p", str(self._port)])
        if self._ssh_options:
            args.extend(shlex.split(self._ssh_options))
        args.append(self._build_ssh_target())
        args.append(remote_command)
        return args

    def _run_ssh(self, remote_command: str, *, allow_failure: bool = False) -> str:
        cmd = self._build_ssh_args(remote_command)
        proc = subprocess.run(cmd, capture_output=True, text=True, check=False, timeout=60)
        if proc.returncode != 0 and not allow_failure:
            stderr = proc.stderr.strip()
            stdout = proc.stdout.strip()
            detail = stderr or stdout or f"exit={proc.returncode}"
            raise RuntimeError(f"SSH command failed: {detail}")
        return proc.stdout

    def _read_remote_text(self, remote_dir: str, rel_path: str) -> str | None:
        remote_file = f"{remote_dir.rstrip('/')}/{rel_path.lstrip('/')}"
        command = (
            "python - <<'PY'\n"
            "from pathlib import Path\n"
            f"p = Path({remote_file!r}).expanduser()\n"
            "if p.exists():\n"
            "    print(p.read_text(encoding='utf-8'))\n"
            "PY"
        )
        output = self._run_ssh(command, allow_failure=True)
        if not output.strip():
            return None
        return output

    def _download_remote_file(self, remote_dir: str, rel_path: str, local_path: Path) -> None:
        remote_file = f"{remote_dir.rstrip('/')}/{rel_path.lstrip('/')}"
        command = (
            "python - <<'PY'\n"
            "import base64\n"
            "from pathlib import Path\n"
            f"p = Path({remote_file!r}).expanduser()\n"
            "if p.exists() and p.is_file():\n"
            "    print(base64.b64encode(p.read_bytes()).decode('ascii'))\n"
            "PY"
        )
        output = self._run_ssh(command, allow_failure=True).strip()
        if not output:
            return
        import base64

        local_path.write_bytes(base64.b64decode(output))


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
    identity_file = str(settings.get("ssh_identity_file", "")).strip()
    if identity_file:
        ssh_options = f"{ssh_options} -i {shlex.quote(identity_file)}".strip()
    poll_raw = str(settings.get("scheduler_poll_seconds", "")).strip() or os.getenv("STOPTIONS_REMOTE_POLL_SECONDS", str(_STATUS_POLL_SECONDS))
    poll_interval = float(poll_raw)
    return RemoteSSHExecutionBackend(
        host=host,
        user=user,
        port=port,
        python_bin=python_bin,
        remote_root=remote_root,
        ssh_options=ssh_options,
        poll_interval_seconds=poll_interval,
    )

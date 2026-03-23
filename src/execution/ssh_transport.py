from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import shlex
import subprocess
import tempfile


_DEFAULT_TIMEOUT_SECONDS = 60


@dataclass(frozen=True)
class SSHTransportConfig:
    host: str
    user: str | None = None
    port: int | None = None
    identity_file: str | None = None
    known_hosts_file: str | None = None
    strict_host_key_checking: bool = True
    extra_options: str = ""


class SSHTransport:
    """Dispatches commands via ssh and transfers files via sftp.

    Security assumptions:
    - key-based SSH auth is pre-provisioned on the client host.
    - strict host-key validation remains enabled (no TOFU/accept-new bypass).
    """

    def __init__(self, config: SSHTransportConfig) -> None:
        host = config.host.strip()
        if not host:
            raise ValueError("SSHTransport requires a host")
        self._config = config
        self._host = host

    def run(self, command: str, *, allow_failure: bool = False, timeout_seconds: int = _DEFAULT_TIMEOUT_SECONDS) -> str:
        proc = subprocess.run(
            self._build_ssh_args(command),
            capture_output=True,
            text=True,
            check=False,
            timeout=timeout_seconds,
        )
        if proc.returncode != 0 and not allow_failure:
            detail = proc.stderr.strip() or proc.stdout.strip() or f"exit={proc.returncode}"
            raise RuntimeError(f"SSH command failed: {detail}")
        return proc.stdout

    def write_text(self, remote_path: str, text: str) -> None:
        target = self._expand_remote_path(remote_path)
        self.run(f"mkdir -p {shlex.quote(target.rsplit('/', 1)[0])}")
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", delete=False) as handle:
            handle.write(text)
            local_path = Path(handle.name)
        try:
            self._run_sftp_batch([
                f"put {shlex.quote(str(local_path))} {shlex.quote(target)}",
            ])
        finally:
            local_path.unlink(missing_ok=True)

    def read_text(self, remote_path: str) -> str | None:
        target = self._expand_remote_path(remote_path)
        with tempfile.NamedTemporaryFile(mode="wb", delete=False) as handle:
            local_path = Path(handle.name)
        try:
            self._run_sftp_batch([f"get {shlex.quote(target)} {shlex.quote(str(local_path))}"], allow_failure=True)
            payload = local_path.read_bytes()
            if not payload:
                return None
            return payload.decode("utf-8")
        finally:
            local_path.unlink(missing_ok=True)

    def download_file(self, remote_path: str, local_path: Path) -> None:
        target = self._expand_remote_path(remote_path)
        local_path.parent.mkdir(parents=True, exist_ok=True)
        self._run_sftp_batch([f"get {shlex.quote(target)} {shlex.quote(str(local_path))}"], allow_failure=True)

    def _build_target(self) -> str:
        if self._config.user:
            return f"{self._config.user}@{self._host}"
        return self._host

    def _base_connection_args(self, *, for_sftp: bool) -> list[str]:
        args: list[str] = []
        if self._config.port:
            args.extend(["-P" if for_sftp else "-p", str(self._config.port)])
        if self._config.identity_file:
            args.extend(["-i", self._config.identity_file])
        strict_value = "yes" if self._config.strict_host_key_checking else "no"
        args.extend(["-o", f"StrictHostKeyChecking={strict_value}"])
        args.extend(["-o", "BatchMode=yes"])
        if self._config.known_hosts_file:
            args.extend(["-o", f"UserKnownHostsFile={self._config.known_hosts_file}"])
        if self._config.extra_options:
            args.extend(shlex.split(self._config.extra_options))
        return args

    def _build_ssh_args(self, command: str) -> list[str]:
        args = ["ssh", *self._base_connection_args(for_sftp=False)]
        args.extend([self._build_target(), command])
        return args

    def _run_sftp_batch(self, lines: list[str], *, allow_failure: bool = False) -> None:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", delete=False) as batch_file:
            batch_file.write("\n".join(lines))
            batch_file.write("\n")
            batch_path = Path(batch_file.name)
        try:
            proc = subprocess.run(
                ["sftp", *self._base_connection_args(for_sftp=True), "-b", str(batch_path), self._build_target()],
                capture_output=True,
                text=True,
                check=False,
                timeout=_DEFAULT_TIMEOUT_SECONDS,
            )
            if proc.returncode != 0 and not allow_failure:
                detail = proc.stderr.strip() or proc.stdout.strip() or f"exit={proc.returncode}"
                raise RuntimeError(f"SFTP command failed: {detail}")
        finally:
            batch_path.unlink(missing_ok=True)

    @staticmethod
    def _expand_remote_path(path: str) -> str:
        cleaned = str(path).strip()
        if cleaned.startswith("~/"):
            return cleaned
        return cleaned

from __future__ import annotations

from execution.remote_ssh_backend import build_remote_backend_from_settings


def test_builder_prefers_remote_venv_python_when_present() -> None:
    backend = build_remote_backend_from_settings(
        {
            "mode": "remote",
            "ssh_host": "example-host",
            "remote_project_path": "~/stoptions_jobs",
            "remote_python_command": "python",
            "remote_venv_path": "/opt/stoptions/.venv",
        }
    )
    assert backend._python_bin == "/opt/stoptions/.venv/bin/python"

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


def test_shell_value_expands_tilde_to_home_expression() -> None:
    assert (
        build_remote_backend_from_settings(
            {
                "mode": "remote",
                "ssh_host": "example-host",
                "remote_project_path": "~/stoptions_jobs",
                "remote_python_command": "python",
            }
        )._shell_value_with_home_expansion("~/venvs/stoptions/bin/python")
        == "$HOME/venvs/stoptions/bin/python"
    )


def test_builder_accepts_direct_python_executable_path_in_remote_venv_field() -> None:
    backend = build_remote_backend_from_settings(
        {
            "mode": "remote",
            "ssh_host": "example-host",
            "remote_project_path": "~/stoptions_jobs",
            "remote_python_command": "python",
            "remote_venv_path": "/opt/stoptions/.venv/bin/python",
        }
    )
    assert backend._python_bin == "/opt/stoptions/.venv/bin/python"

from __future__ import annotations

from config import validate_remote_execution_settings


def test_remote_settings_validation_requires_host_in_remote_mode() -> None:
    errors = validate_remote_execution_settings({"mode": "remote", "ssh_host": ""})
    assert any("SSH host" in item for item in errors)


def test_remote_settings_validation_accepts_local_mode() -> None:
    assert validate_remote_execution_settings({"mode": "local"}) == []


def test_remote_settings_rejects_forward_per_job_policy() -> None:
    errors = validate_remote_execution_settings(
        {
            "mode": "remote",
            "ssh_host": "host",
            "ssh_port": "22",
            "remote_project_path": "~/jobs",
            "remote_python_command": "python",
            "ssh_identity_file": "~/.ssh/id_ed25519",
            "ssh_known_hosts_file": "~/.ssh/known_hosts",
            "ssh_strict_host_key_checking": True,
            "api_key_policy": "forward_per_job",
        }
    )
    assert any("not enabled" in item for item in errors)


def test_remote_settings_require_key_auth_and_strict_host_validation() -> None:
    errors = validate_remote_execution_settings(
        {
            "mode": "remote",
            "ssh_host": "host",
            "ssh_port": "22",
            "remote_project_path": "~/jobs",
            "remote_python_command": "python",
            "ssh_identity_file": "",
            "ssh_known_hosts_file": "",
            "ssh_strict_host_key_checking": False,
        }
    )
    assert any("identity file is required" in item.lower() for item in errors)
    assert any("strict host-key checking" in item.lower() for item in errors)
    assert any("known hosts file" in item.lower() for item in errors)

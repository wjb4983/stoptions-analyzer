from __future__ import annotations

from pathlib import Path

from ui import helpers


def test_load_api_key_prefers_env(monkeypatch, tmp_path):
    monkeypatch.setattr(helpers, "API_KEY_PATH", tmp_path / "api_key.txt")
    monkeypatch.setenv("MASSIVE_API_KEY", " from-env ")
    assert helpers.load_api_key() == "from-env"


def test_load_api_key_reads_file_when_env_missing(monkeypatch, tmp_path):
    key_path = tmp_path / "api_key.txt"
    key_path.write_text("from-file\n", encoding="utf-8")
    monkeypatch.setattr(helpers, "API_KEY_PATH", key_path)
    monkeypatch.delenv("MASSIVE_API_KEY", raising=False)
    assert helpers.load_api_key() == "from-file"


def test_save_api_key_trims_and_writes(monkeypatch, tmp_path):
    config_dir = tmp_path / "cfg"
    key_path = config_dir / "key.txt"
    monkeypatch.setattr(helpers, "CONFIG_DIR", config_dir)
    monkeypatch.setattr(helpers, "API_KEY_PATH", key_path)

    helpers.save_api_key("  abc123  ")
    assert key_path.read_text(encoding="utf-8") == "abc123"


def test_safe_dir_name_guardrails():
    assert helpers._safe_dir_name("CON") == "CON_"
    assert helpers._safe_dir_name("") == "TICKER"
    assert helpers._safe_dir_name(" AAPL ") == "_AAPL_"

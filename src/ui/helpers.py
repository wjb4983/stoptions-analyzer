from __future__ import annotations

import os
import json

from config import API_KEY_PATH, CONFIG_DIR, REMOTE_SECRETS_PATH
from data_access.cache import _safe_ticker_name


def load_api_key() -> str:
    env_key = os.getenv("MASSIVE_API_KEY", "").strip()
    if env_key:
        return env_key
    if API_KEY_PATH.exists():
        return API_KEY_PATH.read_text().strip()
    return ""


def save_api_key(key: str) -> None:
    _ensure_config_dir()
    API_KEY_PATH.write_text(key.strip())
    try:
        API_KEY_PATH.chmod(0o600)
    except OSError:
        pass


def load_remote_secrets() -> dict[str, str]:
    if not REMOTE_SECRETS_PATH.exists():
        return {}
    try:
        payload = json.loads(REMOTE_SECRETS_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(payload, dict):
        return {}
    return {str(key): str(value) for key, value in payload.items()}


def save_remote_secrets(payload: dict[str, str]) -> None:
    _ensure_config_dir()
    safe_payload = {str(key): str(value) for key, value in payload.items() if str(value).strip()}
    REMOTE_SECRETS_PATH.write_text(json.dumps(safe_payload, indent=2), encoding="utf-8")
    try:
        REMOTE_SECRETS_PATH.chmod(0o600)
    except OSError:
        pass


def _ensure_config_dir() -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    try:
        CONFIG_DIR.chmod(0o700)
    except OSError:
        pass


def _safe_dir_name(name: str) -> str:
    safe = _safe_ticker_name(name).rstrip(" .")
    if not safe:
        safe = "TICKER"
    reserved = {
        "CON",
        "PRN",
        "AUX",
        "NUL",
        *(f"COM{index}" for index in range(1, 10)),
        *(f"LPT{index}" for index in range(1, 10)),
    }
    if safe.upper() in reserved:
        return f"{safe}_"
    return safe

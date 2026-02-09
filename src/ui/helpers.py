from __future__ import annotations

import os

from config import API_KEY_PATH, CONFIG_DIR
from data_access.cache import _safe_ticker_name


def load_api_key() -> str:
    env_key = os.getenv("MASSIVE_API_KEY", "").strip()
    if env_key:
        return env_key
    if API_KEY_PATH.exists():
        return API_KEY_PATH.read_text().strip()
    return ""


def save_api_key(key: str) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    API_KEY_PATH.write_text(key.strip())
    try:
        API_KEY_PATH.chmod(0o600)
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

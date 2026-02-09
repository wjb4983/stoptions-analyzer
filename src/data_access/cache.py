import json
from pathlib import Path

from config import DATA_DIR


def _safe_ticker_name(ticker: str) -> str:
    return "".join(char if char.isalnum() else "_" for char in ticker.upper())


def _cache_path(ticker: str) -> Path:
    return DATA_DIR / f"{_safe_ticker_name(ticker)}.json"


def load_cached_market_data(ticker: str) -> dict | None:
    path = _cache_path(ticker)
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text())
    except json.JSONDecodeError:
        return None


def save_cached_market_data(ticker: str, payload: dict) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    path = _cache_path(ticker)
    path.write_text(json.dumps(payload, indent=2))

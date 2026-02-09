from __future__ import annotations

import json
from datetime import date, datetime, timedelta
from html.parser import HTMLParser
from pathlib import Path
from urllib.error import HTTPError
from zoneinfo import ZoneInfo

import numpy as np

from config import BACKTEST_CACHE_DIR


def effective_market_date() -> date:
    now = datetime.now(ZoneInfo("America/New_York"))
    market_close = now.replace(hour=16, minute=0, second=0, microsecond=0)
    if now < market_close:
        return (now - timedelta(days=1)).date()
    return now.date()


def normalize_contract_type(value: str | None) -> str | None:
    if not value:
        return None
    return str(value).strip().upper()


def format_strike(value: float | int | str | None) -> str | None:
    if value is None:
        return None
    try:
        numeric = float(value)
    except (TypeError, ValueError):
        return str(value)
    if numeric.is_integer():
        return str(int(numeric))
    return f"{numeric:.2f}".rstrip("0").rstrip(".")


def normalize_option_records(records: list[dict]) -> list[dict]:
    normalized: list[dict] = []
    for record in records:
        if not isinstance(record, dict):
            continue
        details = record.get("details") or {}
        greeks = record.get("greeks") or {}
        day = record.get("day") or {}
        last_trade = record.get("last_trade") or {}
        last_quote = record.get("last_quote") or {}
        if not isinstance(greeks, dict):
            greeks = {}
        implied_vol = greeks.get("iv")
        if implied_vol is None:
            implied_vol = record.get("implied_volatility") or record.get("implied_vol")
        volume = record.get("volume")
        if volume is None:
            volume = day.get("volume") or day.get("v")
        open_interest = record.get("open_interest") or details.get("open_interest")
        normalized.append(
            {
                "ticker": record.get("ticker") or details.get("ticker"),
                "expiration_date": record.get("expiration_date") or details.get("expiration_date"),
                "contract_type": record.get("contract_type") or details.get("contract_type"),
                "strike_price": record.get("strike_price") or details.get("strike_price"),
                "implied_volatility": implied_vol,
                "volume": volume,
                "open_interest": open_interest,
                "day_close": record.get("close")
                or day.get("close")
                or day.get("c")
                or record.get("day_close"),
                "bid": record.get("bid")
                or last_quote.get("bid")
                or last_quote.get("bid_price")
                or last_quote.get("bp"),
                "ask": record.get("ask")
                or last_quote.get("ask")
                or last_quote.get("ask_price")
                or last_quote.get("ap"),
                "last": record.get("last")
                or last_trade.get("price")
                or last_trade.get("p"),
                "greeks": {
                    "delta": greeks.get("delta"),
                    "gamma": greeks.get("gamma"),
                    "theta": greeks.get("theta"),
                    "vega": greeks.get("vega"),
                    "rho": greeks.get("rho"),
                    "iv": implied_vol,
                },
            }
        )
    return normalized


def extract_greeks(contract: dict) -> dict:
    greeks = contract.get("greeks") or {}
    if not isinstance(greeks, dict):
        greeks = {}
    implied_vol = greeks.get("iv")
    if implied_vol is None:
        implied_vol = contract.get("implied_volatility") or contract.get("implied_vol")
    return {
        "delta": greeks.get("delta"),
        "gamma": greeks.get("gamma"),
        "theta": greeks.get("theta"),
        "vega": greeks.get("vega"),
        "rho": greeks.get("rho"),
        "iv": implied_vol,
    }


def option_mid_price(contract: dict) -> float | None:
    bid = contract.get("bid")
    ask = contract.get("ask")
    last = contract.get("last")
    day_close = contract.get("day_close")
    if isinstance(bid, (int, float)) and isinstance(ask, (int, float)):
        return (bid + ask) / 2
    if isinstance(last, (int, float)):
        return float(last)
    if isinstance(day_close, (int, float)):
        return float(day_close)
    if isinstance(bid, (int, float)):
        return float(bid)
    if isinstance(ask, (int, float)):
        return float(ask)
    return None


def option_likelihood(contract: dict) -> float | None:
    greeks = extract_greeks(contract)
    delta = greeks.get("delta")
    if delta is None:
        return None
    try:
        likelihood = abs(float(delta))
    except (TypeError, ValueError):
        return None
    return max(0.0, min(1.0, likelihood))


def combine_greeks(long_leg: dict, short_leg: dict) -> dict:
    long_greeks = extract_greeks(long_leg)
    short_greeks = extract_greeks(short_leg)
    combined: dict[str, float | None] = {}
    for key in ("delta", "gamma", "theta", "vega", "rho"):
        long_value = long_greeks.get(key)
        short_value = short_greeks.get(key)
        if isinstance(long_value, (int, float)) or isinstance(short_value, (int, float)):
            combined[key] = (long_value or 0) - (short_value or 0)
        else:
            combined[key] = None
    iv_long = long_greeks.get("iv")
    iv_short = short_greeks.get("iv")
    if isinstance(iv_long, (int, float)) and isinstance(iv_short, (int, float)):
        combined["iv"] = (iv_long + iv_short) / 2
    else:
        combined["iv"] = iv_long if iv_long is not None else iv_short
    return combined


def parse_float(value: str) -> float | None:
    text = value.strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def parse_int(value: str) -> int | None:
    text = value.strip()
    if not text:
        return None
    try:
        return int(text)
    except ValueError:
        return None


def parse_date(value: str) -> date | None:
    text = value.strip()
    if not text:
        return None
    try:
        return datetime.strptime(text, "%Y-%m-%d").date()
    except ValueError:
        return None


def normalize_cache_root(value: str) -> Path:
    text = value.strip()
    if not text:
        return BACKTEST_CACHE_DIR
    return Path(text).expanduser()


def chunk_results_by_year(results: list[dict]) -> dict[int, list[dict]]:
    buckets: dict[int, list[dict]] = {}
    for entry in results:
        timestamp = entry.get("t")
        if timestamp is None:
            continue
        try:
            stamp = datetime.fromtimestamp(timestamp / 1000, tz=ZoneInfo("America/New_York"))
        except (TypeError, ValueError, OSError):
            continue
        buckets.setdefault(stamp.year, []).append(entry)
    return buckets


def build_npz_payload(entries: list[dict]) -> dict[str, np.ndarray]:
    def _extract(key: str, default: float = float("nan")) -> np.ndarray:
        return np.array([entry.get(key, default) for entry in entries], dtype=float)

    payload = {
        "t": np.array([entry.get("t", 0) for entry in entries], dtype=np.int64),
        "o": _extract("o"),
        "h": _extract("h"),
        "l": _extract("l"),
        "c": _extract("c"),
        "v": _extract("v"),
        "n": _extract("n"),
    }
    return payload


def normalize_likelihood_threshold(value: str) -> float | None:
    raw = parse_float(value)
    if raw is None:
        return None
    if raw > 1:
        raw = raw / 100
    return max(0.0, min(1.0, raw))


def _parse_iso_date(value: str | None) -> date | None:
    if not value:
        return None
    try:
        normalized = value.replace("Z", "+00:00")
        return datetime.fromisoformat(normalized).date()
    except ValueError:
        return None


def _coerce_number(value: object) -> float | None:
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, dict):
        nested = value.get("value")
        if isinstance(nested, (int, float)):
            return float(nested)
    return None


def _get_nested_value(payload: dict, paths: list[tuple[str, ...]]) -> float | None:
    for path in paths:
        current: object = payload
        for key in path:
            if not isinstance(current, dict):
                current = None
                break
            current = current.get(key)
        value = _coerce_number(current)
        if value is not None:
            return value
    return None


def _has_fundamentals_data(payload: dict) -> bool:
    for value in payload.values():
        if value is not None:
            return True
    return False


def strip_html(text: str) -> str:
    class _HTMLStripper(HTMLParser):
        def __init__(self) -> None:
            super().__init__()
            self.parts: list[str] = []

        def handle_data(self, data: str) -> None:
            self.parts.append(data)

    stripper = _HTMLStripper()
    stripper.feed(text)
    return " ".join(stripper.parts)


def format_http_error_detail(exc: HTTPError) -> str:
    try:
        body = exc.read().decode("utf-8").strip()
    except Exception:
        return ""
    if not body:
        return ""
    if "<html" in body.lower():
        body = strip_html(body)
    try:
        payload = json.loads(body)
    except json.JSONDecodeError:
        return body
    return payload.get("message") or payload.get("error") or payload.get("msg") or body

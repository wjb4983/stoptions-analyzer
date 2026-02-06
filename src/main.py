import json
import math
import os
import random
import socket
import time
import threading
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from html.parser import HTMLParser
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import urlopen
from zoneinfo import ZoneInfo

from analysis.cross_sectional import (
    CrossSectionalSettings,
    MomentumSettings,
    STRATEGY_REGISTRY,
    compute_cross_sectional_momentum,
)
from analysis.reporting import format_cross_sectional_report, format_time_series_report
from analysis.time_series import (
    TIME_SERIES_STRATEGY_REGISTRY,
    TimeSeriesMomentumSettings,
    TimeSeriesSettings,
    compute_time_series_momentum,
)


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


def load_option_records(api_client: "MassiveApiClient", ticker: str) -> list[dict]:
    cache_payload = load_cached_market_data(ticker) or {}
    cache_date = cache_payload.get("last_updated")
    today_label = effective_market_date().isoformat()
    cached_options = cache_payload.get("options")
    if cached_options is not None and cache_date == today_label:
        return normalize_option_records(cached_options or [])
    option_data = api_client.fetch_option_snapshots(ticker)
    option_records = normalize_option_records(option_data)
    cache_payload.update(
        {
            "last_updated": today_label,
            "options": option_records,
        }
    )
    save_cached_market_data(ticker, cache_payload)
    return option_records


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
STATE_PATH = Path(__file__).resolve().parent / "app_state.txt"
CONFIG_DIR = Path.home() / ".stoptions_analyzer"
API_KEY_PATH = CONFIG_DIR / "api_key.txt"
DATA_DIR = Path(__file__).resolve().parent / "data"
API_BASE_URL = os.getenv("MASSIVE_BASE_URL", "https://api.polygon.io")
HORIZON_CONFIGS = [
    ("Day", 1, 10, "10m"),
    ("3 Day", 3, 30, "30m"),
    ("Week", 7, 60, "1h"),
    ("Month", 30, 120, "2h"),
    ("3M", 90, 360, "6h"),
    ("6M", 180, 720, "12h"),
    ("12M", 365, 1440, "1d"),
    ("3Y", 1095, 4320, "3d"),
    ("5Y", 1825, 7200, "5d"),
    ("10Y", 3650, 10080, "7d"),
]
ANALYSIS_OUTPUT_DIR = DATA_DIR / "analysis_outputs"
BACKTEST_CACHE_DIR = DATA_DIR / "backtest_cache"
BACKTEST_OUTPUT_DIR = DATA_DIR / "backtest_outputs"
DEFAULT_GENERAL_ANALYSIS_SETTINGS = {
    "analysis_type": "Cross-Sectional",
    "cross_sectional_strategy": "Momentum",
    "time_series_strategy": "Momentum",
    "lookback_days": 90,
    "skip_days": 5,
    "top_quantile": 0.2,
    "bottom_quantile": 0.2,
    "momentum_use_volatility_scaling": False,
    "momentum_use_residual": False,
    "momentum_use_multi_horizon": False,
    "time_series_use_volatility_scaling": False,
    "time_series_use_residual": False,
    "time_series_use_multi_horizon": False,
    "time_series_use_zscore": False,
    "time_series_winsorize_sigma": None,
    "output_dir": str(ANALYSIS_OUTPUT_DIR),
}

DEFAULT_BACKTEST_SETTINGS = {
    "strategy_type": "Rule-based",
    "ai_model": "Baseline (No ML)",
    "stat_method": "Momentum",
    "data_source": "Massive API",
    "data_granularity": "Daily (EOD)",
    "start_date": "",
    "end_date": "",
    "training_split": "70/30",
    "slippage_bps": "5",
    "commission_per_contract": "0.65",
    "fill_probability": "0.9",
    "use_bid_ask": True,
    "model_walk_forward": False,
    "max_position_pct": "10",
    "notes": "",
}


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


class MassiveApiClient:
    def __init__(self, api_key: str, base_url: str = API_BASE_URL) -> None:
        self.api_key = api_key
        self.base_url = base_url.rstrip("/")

    def _request(self, path: str, params: dict[str, str]) -> dict:
        params = {**params, "apiKey": self.api_key}
        url = f"{self.base_url}{path}?{urlencode(params)}"
        with urlopen(url, timeout=10) as response:
            payload = response.read().decode("utf-8")
        return json.loads(payload)

    def fetch_ticker_details(self, ticker: str) -> dict:
        data = self._request(f"/v3/reference/tickers/{ticker}", {})
        result = data.get("results") or {}
        return {
            "market_cap": result.get("market_cap"),
            "share_class_shares_outstanding": result.get("share_class_shares_outstanding"),
            "weighted_shares_outstanding": result.get("weighted_shares_outstanding"),
        }

    def fetch_financials(self, ticker: str, period: str = "annual") -> list[dict]:
        data = self._request(
            "/v3/reference/financials",
            {"ticker": ticker, "period": period, "limit": "4"},
        )
        return data.get("results", [])

    def fetch_dividends(self, ticker: str) -> list[dict]:
        data = self._request("/v3/reference/dividends", {"ticker": ticker, "limit": "100"})
        return data.get("results", [])

    def fetch_earnings(self, ticker: str) -> list[dict]:
        data = self._request("/v3/reference/earnings", {"ticker": ticker, "limit": "8"})
        return data.get("results", [])

    def fetch_previous_close(self, ticker: str) -> dict:
        data = self._request(f"/v2/aggs/ticker/{ticker}/prev", {"adjusted": "true"})
        result = (data.get("results") or [{}])[0]
        return {
            "close": result.get("c"),
            "open": result.get("o"),
            "high": result.get("h"),
            "low": result.get("l"),
            "volume": result.get("v"),
        }

    def _request_url(self, url: str) -> dict:
        with urlopen(url, timeout=10) as response:
            payload = response.read().decode("utf-8")
        return json.loads(payload)

    def fetch_option_contracts(self, ticker: str, limit: int = 1000) -> list[dict]:
        results: list[dict] = []
        params = {"underlying_ticker": ticker, "limit": str(limit)}
        data = self._request("/v3/reference/options/contracts", params)
        results.extend(data.get("results", []))
        next_url = data.get("next_url")
        while next_url:
            if "apiKey=" not in next_url:
                joiner = "&" if "?" in next_url else "?"
                next_url = f"{next_url}{joiner}apiKey={self.api_key}"
            data = self._request_url(next_url)
            results.extend(data.get("results", []))
            next_url = data.get("next_url")
        return results

    def fetch_option_snapshots(self, ticker: str, limit: int = 250) -> list[dict]:
        results: list[dict] = []
        params = {"limit": str(limit)}
        data = self._request(f"/v3/snapshot/options/{ticker}", params)
        results.extend(self._normalize_option_snapshots(data.get("results", [])))
        next_url = data.get("next_url")
        while next_url:
            if "apiKey=" not in next_url:
                joiner = "&" if "?" in next_url else "?"
                next_url = f"{next_url}{joiner}apiKey={self.api_key}"
            data = self._request_url(next_url)
            results.extend(self._normalize_option_snapshots(data.get("results", [])))
            next_url = data.get("next_url")
        return results

    def _normalize_option_snapshots(self, snapshots: list[dict]) -> list[dict]:
        normalized: list[dict] = []
        for snapshot in snapshots:
            details = snapshot.get("details", {}) or {}
            greeks = snapshot.get("greeks", {}) or {}
            day = snapshot.get("day", {}) or {}
            last_trade = snapshot.get("last_trade", {}) or {}
            last_quote = snapshot.get("last_quote", {}) or {}
            implied_vol = snapshot.get("implied_volatility")
            if implied_vol is not None and "iv" not in greeks:
                greeks = {**greeks, "iv": implied_vol}
            volume = snapshot.get("volume")
            if volume is None:
                volume = day.get("volume") or day.get("v")
            open_interest = snapshot.get("open_interest")
            if open_interest is None:
                open_interest = details.get("open_interest")
            normalized.append(
                {
                    "ticker": details.get("ticker") or snapshot.get("ticker"),
                    "expiration_date": details.get("expiration_date"),
                    "contract_type": details.get("contract_type"),
                    "strike_price": details.get("strike_price"),
                    "greeks": greeks,
                    "implied_volatility": implied_vol,
                    "volume": volume,
                    "open_interest": open_interest,
                    "day_close": snapshot.get("close") or day.get("close") or day.get("c"),
                    "bid": last_quote.get("bid")
                    or last_quote.get("bid_price")
                    or last_quote.get("bp"),
                    "ask": last_quote.get("ask")
                    or last_quote.get("ask_price")
                    or last_quote.get("ap"),
                    "last": last_trade.get("price") or last_trade.get("p"),
                }
            )
        return normalized

    def fetch_aggregates(self, ticker: str, days_back: int, minutes_per_bar: int) -> list[dict]:
        if days_back == 1:
            now = datetime.now(ZoneInfo("America/New_York"))
            market_close = now.replace(hour=16, minute=0, second=0, microsecond=0)
            end_date = (now - timedelta(days=1)).date() if now < market_close else now.date()
            start_date = end_date
        else:
            end_date = date.today()
            start_date = end_date - timedelta(days=days_back)
        data = self._request(
            f"/v2/aggs/ticker/{ticker}/range/{minutes_per_bar}/minute/{start_date}/{end_date}",
            {"adjusted": "true", "sort": "asc", "limit": "5000"},
        )
        return data.get("results", [])

    def fetch_daily_aggregates(self, ticker: str, days_back: int) -> list[dict]:
        if days_back == 1:
            now = datetime.now(ZoneInfo("America/New_York"))
            market_close = now.replace(hour=16, minute=0, second=0, microsecond=0)
            end_date = (now - timedelta(days=1)).date() if now < market_close else now.date()
            start_date = end_date
        else:
            end_date = date.today()
            start_date = end_date - timedelta(days=days_back)
        data = self._request(
            f"/v2/aggs/ticker/{ticker}/range/1/day/{start_date}/{end_date}",
            {"adjusted": "true", "sort": "asc", "limit": "5000"},
        )
        return data.get("results", [])

    def fetch_grouped_daily_aggregates(self, on_date: date) -> list[dict]:
        payload = self._request(
            f"/v2/aggs/grouped/locale/us/market/stocks/{on_date}",
            {"adjusted": "true"},
        )
        return payload.get("results", [])

    def fetch_aggregates_range(
        self, ticker: str, start_date: date, end_date: date, minutes_per_bar: int = 1
    ) -> list[dict]:
        params = {"adjusted": "true", "sort": "asc", "limit": "50000"}
        data = self._request(
            f"/v2/aggs/ticker/{ticker}/range/{minutes_per_bar}/minute/{start_date}/{end_date}",
            params,
        )
        results: list[dict] = []
        results.extend(data.get("results", []))
        next_url = data.get("next_url")
        while next_url:
            if "apiKey=" not in next_url:
                joiner = "&" if "?" in next_url else "?"
                next_url = f"{next_url}{joiner}apiKey={self.api_key}"
            payload = self._request_url(next_url)
            results.extend(payload.get("results", []))
            next_url = payload.get("next_url")
        return results


@dataclass
class AppState:
    tickers: list[str] = field(default_factory=list)
    selected_ticker: str | None = None
    analysis_mode: str = "Stock Analysis"
    option_strategy: str = "Naked Call"
    general_analysis_settings: dict[str, object] = field(
        default_factory=lambda: dict(DEFAULT_GENERAL_ANALYSIS_SETTINGS)
    )
    backtest_settings: dict[str, object] = field(
        default_factory=lambda: dict(DEFAULT_BACKTEST_SETTINGS)
    )

    def save(self) -> None:
        payload = {
            "tickers": self.tickers,
            "selected_ticker": self.selected_ticker,
            "analysis_mode": self.analysis_mode,
            "option_strategy": self.option_strategy,
            "general_analysis_settings": self.general_analysis_settings,
            "backtest_settings": self.backtest_settings,
        }
        STATE_PATH.write_text(json.dumps(payload, indent=2))

    @classmethod
    def load(cls) -> "AppState":
        if not STATE_PATH.exists():
            return cls()
        try:
            payload = json.loads(STATE_PATH.read_text())
        except json.JSONDecodeError:
            return cls()
        return cls(
            tickers=payload.get("tickers", []),
            selected_ticker=payload.get("selected_ticker"),
            analysis_mode=payload.get("analysis_mode", payload.get("analysis_type", "Stock Analysis")),
            option_strategy=payload.get("option_strategy", "Naked Call"),
            general_analysis_settings=payload.get(
                "general_analysis_settings", dict(DEFAULT_GENERAL_ANALYSIS_SETTINGS)
            ),
            backtest_settings=payload.get(
                "backtest_settings", dict(DEFAULT_BACKTEST_SETTINGS)
            ),
        )


class StoptionsApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Stoptions Analyzer")
        self.geometry("1200x800")
        self._maximize_window()
        self.state = AppState.load()
        self.api_key = load_api_key()

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        self.frames: dict[str, ttk.Frame] = {}
        for frame_cls in (
            MainMenu,
            BacktestingPage,
            TickerEntryPage,
            TickerSelectPage,
            AnalysisPage,
            GeneralAnalysisPage,
            CallPutAnalysisPage,
            SpreadAnalysisPage,
        ):
            frame = frame_cls(container, self)
            self.frames[frame_cls.__name__] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("MainMenu")

    def show_frame(self, name: str) -> None:
        frame = self.frames[name]
        if hasattr(frame, "refresh"):
            frame.refresh()
        frame.tkraise()

    def persist_state(self) -> None:
        self.state.save()

    def _maximize_window(self) -> None:
        self.update_idletasks()
        try:
            self.state("zoomed")
        except tk.TclError:
            self.attributes("-fullscreen", True)


class MainMenu(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller

        title = ttk.Label(self, text="Stoptions Analyzer", font=("Arial", 24, "bold"))
        title.pack(pady=20)

        description = ttk.Label(
            self,
            text="Manage tickers, select a stock, and explore option strategy analysis.",
            wraplength=600,
            justify="center",
        )
        description.pack(pady=10)

        api_frame = ttk.LabelFrame(self, text="Massive API Key")
        api_frame.pack(pady=15, padx=40, fill="x")
        api_frame.columnconfigure(1, weight=1)

        ttk.Label(api_frame, text="API Key").grid(row=0, column=0, padx=10, pady=8, sticky="w")
        self.api_key_var = tk.StringVar(value=self.controller.api_key)
        self.api_key_entry = ttk.Entry(api_frame, textvariable=self.api_key_var, show="*")
        self.api_key_entry.grid(row=0, column=1, padx=10, pady=8, sticky="ew")
        ttk.Button(api_frame, text="Save Key", command=self.save_api_key).grid(
            row=0, column=2, padx=10, pady=8
        )

        button_frame = ttk.Frame(self)
        button_frame.pack(pady=40)

        ttk.Button(
            button_frame,
            text="Enter Stock Tickers",
            command=lambda: controller.show_frame("TickerEntryPage"),
            width=30,
        ).grid(row=0, column=0, pady=10)

        ttk.Button(
            button_frame,
            text="Select Stock",
            command=lambda: controller.show_frame("TickerSelectPage"),
            width=30,
        ).grid(row=1, column=0, pady=10)

        ttk.Button(
            button_frame,
            text="Analysis",
            command=lambda: controller.show_frame("AnalysisPage"),
            width=30,
        ).grid(row=2, column=0, pady=10)

        ttk.Button(
            button_frame,
            text="General Analysis",
            command=lambda: controller.show_frame("GeneralAnalysisPage"),
            width=30,
        ).grid(row=3, column=0, pady=10)

        ttk.Button(
            button_frame,
            text="Backtesting",
            command=lambda: controller.show_frame("BacktestingPage"),
            width=30,
        ).grid(row=4, column=0, pady=10)

    def refresh(self) -> None:
        self.api_key_var.set(self.controller.api_key)

    def save_api_key(self) -> None:
        key = self.api_key_var.get().strip()
        if not key:
            messagebox.showinfo("Missing key", "Enter a Massive API key first.")
            return
        save_api_key(key)
        self.controller.api_key = key
        messagebox.showinfo(
            "Saved", f"API key saved to {API_KEY_PATH} (not tracked in git)."
        )



class BacktestingPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller

        ttk.Label(self, text="Backtesting Parameters", font=("Arial", 18, "bold")).pack(
            pady=10
        )

        intro = (
            "Configure strategy selection, data, and execution realism. "
            "Use intraday data for realistic options backtests; EOD is cheaper and "
            "better for swing-style research."
        )
        ttk.Label(self, text=intro, wraplength=900, justify="center").pack(pady=5)

        content = ttk.Frame(self)
        content.pack(fill="both", expand=True, padx=30, pady=10)
        content.columnconfigure(0, weight=1)
        content.columnconfigure(1, weight=1)

        strategy_frame = ttk.LabelFrame(content, text="Strategy / Model")
        strategy_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        strategy_frame.columnconfigure(1, weight=1)

        ttk.Label(strategy_frame, text="Strategy Type").grid(
            row=0, column=0, sticky="w", padx=8, pady=6
        )
        self.strategy_type_var = tk.StringVar()
        self.strategy_type_combo = ttk.Combobox(
            strategy_frame,
            textvariable=self.strategy_type_var,
            values=["Rule-based", "ML signals", "Hybrid"],
            state="readonly",
        )
        self.strategy_type_combo.grid(row=0, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(strategy_frame, text="AI Model").grid(
            row=1, column=0, sticky="w", padx=8, pady=6
        )
        self.ai_model_var = tk.StringVar()
        self.ai_model_combo = ttk.Combobox(
            strategy_frame,
            textvariable=self.ai_model_var,
            values=[
                "Baseline (No ML)",
                "Random Forest",
                "XGBoost",
                "LSTM",
                "Custom",
            ],
            state="readonly",
        )
        self.ai_model_combo.grid(row=1, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(strategy_frame, text="Stat Method").grid(
            row=2, column=0, sticky="w", padx=8, pady=6
        )
        self.stat_method_var = tk.StringVar()
        self.stat_method_combo = ttk.Combobox(
            strategy_frame,
            textvariable=self.stat_method_var,
            values=[
                "Momentum",
                "Mean Reversion",
                "Volatility Breakout",
                "Pairs Trading",
                "Custom",
            ],
            state="readonly",
        )
        self.stat_method_combo.grid(row=2, column=1, sticky="ew", padx=8, pady=6)

        data_frame = ttk.LabelFrame(content, text="Data")
        data_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        data_frame.columnconfigure(1, weight=1)

        ttk.Label(data_frame, text="Data Source").grid(
            row=0, column=0, sticky="w", padx=8, pady=6
        )
        self.data_source_var = tk.StringVar()
        self.data_source_combo = ttk.Combobox(
            data_frame,
            textvariable=self.data_source_var,
            values=["Massive API", "Discount Option Data", "Custom CSV"],
            state="readonly",
        )
        self.data_source_combo.grid(row=0, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(data_frame, text="Granularity").grid(
            row=1, column=0, sticky="w", padx=8, pady=6
        )
        self.data_granularity_var = tk.StringVar()
        self.data_granularity_combo = ttk.Combobox(
            data_frame,
            textvariable=self.data_granularity_var,
            values=["Daily (EOD)", "60m", "30m", "15m", "5m", "1m"],
            state="readonly",
        )
        self.data_granularity_combo.grid(row=1, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(data_frame, text="Start Date (YYYY-MM-DD)").grid(
            row=2, column=0, sticky="w", padx=8, pady=6
        )
        self.start_date_var = tk.StringVar()
        ttk.Entry(data_frame, textvariable=self.start_date_var).grid(
            row=2, column=1, sticky="ew", padx=8, pady=6
        )

        ttk.Label(data_frame, text="End Date (YYYY-MM-DD)").grid(
            row=3, column=0, sticky="w", padx=8, pady=6
        )
        self.end_date_var = tk.StringVar()
        ttk.Entry(data_frame, textvariable=self.end_date_var).grid(
            row=3, column=1, sticky="ew", padx=8, pady=6
        )

        ttk.Label(data_frame, text="Training Split").grid(
            row=4, column=0, sticky="w", padx=8, pady=6
        )
        self.training_split_var = tk.StringVar()
        ttk.Entry(data_frame, textvariable=self.training_split_var).grid(
            row=4, column=1, sticky="ew", padx=8, pady=6
        )

        realism_frame = ttk.LabelFrame(content, text="Execution Realism")
        realism_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        realism_frame.columnconfigure(1, weight=1)

        ttk.Label(realism_frame, text="Slippage (bps)").grid(
            row=0, column=0, sticky="w", padx=8, pady=6
        )
        self.slippage_bps_var = tk.StringVar()
        ttk.Entry(realism_frame, textvariable=self.slippage_bps_var).grid(
            row=0, column=1, sticky="ew", padx=8, pady=6
        )

        ttk.Label(realism_frame, text="Commission per Contract ($)").grid(
            row=1, column=0, sticky="w", padx=8, pady=6
        )
        self.commission_var = tk.StringVar()
        ttk.Entry(realism_frame, textvariable=self.commission_var).grid(
            row=1, column=1, sticky="ew", padx=8, pady=6
        )

        ttk.Label(realism_frame, text="Fill Probability (0-1)").grid(
            row=2, column=0, sticky="w", padx=8, pady=6
        )
        self.fill_probability_var = tk.StringVar()
        ttk.Entry(realism_frame, textvariable=self.fill_probability_var).grid(
            row=2, column=1, sticky="ew", padx=8, pady=6
        )

        self.use_bid_ask_var = tk.BooleanVar()
        ttk.Checkbutton(
            realism_frame, text="Use bid/ask spread", variable=self.use_bid_ask_var
        ).grid(row=3, column=0, columnspan=2, sticky="w", padx=8, pady=4)

        self.walk_forward_var = tk.BooleanVar()
        ttk.Checkbutton(
            realism_frame, text="Walk-forward retraining", variable=self.walk_forward_var
        ).grid(row=4, column=0, columnspan=2, sticky="w", padx=8, pady=4)

        ttk.Label(realism_frame, text="Max Position Size (%)").grid(
            row=5, column=0, sticky="w", padx=8, pady=6
        )
        self.max_position_pct_var = tk.StringVar()
        ttk.Entry(realism_frame, textvariable=self.max_position_pct_var).grid(
            row=5, column=1, sticky="ew", padx=8, pady=6
        )

        notes_frame = ttk.LabelFrame(content, text="Notes / Metrics")
        notes_frame.grid(row=1, column=1, sticky="nsew", padx=10, pady=10)
        notes_frame.columnconfigure(0, weight=1)
        self.notes_text = tk.Text(notes_frame, height=8)
        self.notes_text.grid(row=0, column=0, sticky="nsew", padx=8, pady=6)

        button_row = ttk.Frame(self)
        button_row.pack(pady=10)

        ttk.Button(button_row, text="Save Parameters", command=self.save_settings).grid(
            row=0, column=0, padx=10
        )
        self.run_button = ttk.Button(
            button_row,
            text="Run Backtest (Cache Data)",
            command=self.run_backtest_cache,
        )
        self.run_button.grid(row=0, column=1, padx=10)
        ttk.Button(
            button_row,
            text="Back to Main Menu",
            command=lambda: controller.show_frame("MainMenu"),
        ).grid(row=0, column=2, padx=10)

    def refresh(self) -> None:
        settings = dict(DEFAULT_BACKTEST_SETTINGS)
        settings.update(self.controller.state.backtest_settings)
        self.strategy_type_var.set(settings.get("strategy_type", "Rule-based"))
        self.ai_model_var.set(settings.get("ai_model", "Baseline (No ML)"))
        self.stat_method_var.set(settings.get("stat_method", "Momentum"))
        self.data_source_var.set(settings.get("data_source", "Massive API"))
        self.data_granularity_var.set(settings.get("data_granularity", "Daily (EOD)"))
        self.start_date_var.set(settings.get("start_date", ""))
        self.end_date_var.set(settings.get("end_date", ""))
        self.training_split_var.set(settings.get("training_split", "70/30"))
        self.slippage_bps_var.set(settings.get("slippage_bps", "5"))
        self.commission_var.set(settings.get("commission_per_contract", "0.65"))
        self.fill_probability_var.set(settings.get("fill_probability", "0.9"))
        self.use_bid_ask_var.set(bool(settings.get("use_bid_ask", True)))
        self.walk_forward_var.set(bool(settings.get("model_walk_forward", False)))
        self.max_position_pct_var.set(settings.get("max_position_pct", "10"))
        self.notes_text.delete("1.0", tk.END)
        self.notes_text.insert("1.0", settings.get("notes", ""))

    def save_settings(self) -> None:
        slippage_bps = parse_float(self.slippage_bps_var.get())
        commission = parse_float(self.commission_var.get())
        fill_probability = parse_float(self.fill_probability_var.get())
        max_position_pct = parse_float(self.max_position_pct_var.get())

        if slippage_bps is None or slippage_bps < 0:
            messagebox.showinfo("Invalid input", "Slippage must be zero or positive.")
            return
        if commission is None or commission < 0:
            messagebox.showinfo("Invalid input", "Commission must be zero or positive.")
            return
        if fill_probability is None or not (0 <= fill_probability <= 1):
            messagebox.showinfo(
                "Invalid input", "Fill probability must be between 0 and 1."
            )
            return
        if max_position_pct is None or not (0 < max_position_pct <= 100):
            messagebox.showinfo(
                "Invalid input", "Max position size must be between 0 and 100."
            )
            return

        self.controller.state.backtest_settings = {
            "strategy_type": self.strategy_type_var.get(),
            "ai_model": self.ai_model_var.get(),
            "stat_method": self.stat_method_var.get(),
            "data_source": self.data_source_var.get(),
            "data_granularity": self.data_granularity_var.get(),
            "start_date": self.start_date_var.get().strip(),
            "end_date": self.end_date_var.get().strip(),
            "training_split": self.training_split_var.get().strip(),
            "slippage_bps": str(slippage_bps),
            "commission_per_contract": str(commission),
            "fill_probability": str(fill_probability),
            "use_bid_ask": self.use_bid_ask_var.get(),
            "model_walk_forward": self.walk_forward_var.get(),
            "max_position_pct": str(max_position_pct),
            "notes": self.notes_text.get("1.0", tk.END).strip(),
        }
        self.controller.persist_state()
        messagebox.showinfo("Saved", "Backtesting parameters saved.")

    def run_backtest_cache(self) -> None:
        if not self.controller.api_key:
            messagebox.showinfo("Missing key", "Enter a Massive API key first.")
            return
        tickers = list(self.controller.state.tickers)
        if not tickers:
            messagebox.showinfo("No tickers", "Add tickers before running a backtest.")
            return

        start_date = parse_date(self.start_date_var.get())
        end_date = parse_date(self.end_date_var.get())
        if end_date is None:
            end_date = date.today()
        if start_date is None:
            start_date = end_date - timedelta(days=365 * 5)
        if start_date >= end_date:
            messagebox.showinfo("Invalid dates", "Start date must be before end date.")
            return

        self.run_button.config(state="disabled")
        self.notes_text.delete("1.0", tk.END)
        self.notes_text.insert(
            "1.0",
            "Running backtest data cache...\nThis may take a while for 1-minute data.\n",
        )

        thread = threading.Thread(
            target=self._run_backtest_worker,
            args=(tickers, start_date, end_date),
            daemon=True,
        )
        thread.start()

    def _run_backtest_worker(
        self, tickers: list[str], start_date: date, end_date: date
    ) -> None:
        api_client = MassiveApiClient(self.controller.api_key)
        BACKTEST_CACHE_DIR.mkdir(parents=True, exist_ok=True)
        BACKTEST_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        lines: list[str] = []
        for ticker in tickers:
            safe_ticker = _safe_ticker_name(ticker)
            cache_path = (
                BACKTEST_CACHE_DIR
                / f"{safe_ticker}_1m_{start_date.isoformat()}_{end_date.isoformat()}.json"
            )
            try:
                results = api_client.fetch_aggregates_range(
                    ticker, start_date, end_date, minutes_per_bar=1
                )
                cache_path.write_text(
                    json.dumps(
                        {
                            "ticker": ticker,
                            "start_date": start_date.isoformat(),
                            "end_date": end_date.isoformat(),
                            "results": results,
                        },
                        indent=2,
                    )
                )
                if results:
                    sample = random.choice(results)
                    sample_text = (
                        f"{ticker}: sample close={sample.get('c')} "
                        f"timestamp={sample.get('t')}"
                    )
                else:
                    sample_text = f"{ticker}: no data returned"
            except Exception as exc:
                sample_text = f"{ticker}: error fetching data ({exc})"
            lines.append(sample_text)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = BACKTEST_OUTPUT_DIR / f"backtest_cache_{timestamp}.txt"
        output_path.write_text("\n".join(lines))
        output_text = "\n".join(lines) + f"\n\nSaved summary to: {output_path}"
        self.after(0, lambda: self._finish_backtest_run(output_text))

    def _finish_backtest_run(self, output_text: str) -> None:
        self.notes_text.delete("1.0", tk.END)
        self.notes_text.insert("1.0", output_text)
        self.run_button.config(state="normal")



class TickerEntryPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller

        ttk.Label(self, text="Enter Stock Tickers", font=("Arial", 18, "bold")).pack(pady=10)

        instructions = ttk.Label(
            self,
            text="Enter one ticker per line. Click Save to store them locally.",
        )
        instructions.pack(pady=5)

        self.text_box = tk.Text(self, height=18, width=40)
        self.text_box.pack(pady=10)

        button_row = ttk.Frame(self)
        button_row.pack(pady=10)

        ttk.Button(button_row, text="Save", command=self.save_tickers).grid(row=0, column=0, padx=10)
        ttk.Button(
            button_row,
            text="Back to Main Menu",
            command=lambda: controller.show_frame("MainMenu"),
        ).grid(row=0, column=1, padx=10)

    def refresh(self) -> None:
        self.text_box.delete("1.0", tk.END)
        self.text_box.insert("1.0", "\n".join(self.controller.state.tickers))

    def save_tickers(self) -> None:
        raw = self.text_box.get("1.0", tk.END)
        tickers = [line.strip().upper() for line in raw.splitlines() if line.strip()]
        if not tickers:
            messagebox.showinfo("No tickers", "Please enter at least one ticker.")
            return
        self.controller.state.tickers = tickers
        if self.controller.state.selected_ticker not in tickers:
            self.controller.state.selected_ticker = tickers[0]
        self.controller.persist_state()
        messagebox.showinfo("Saved", "Tickers saved successfully.")


class TickerSelectPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller

        ttk.Label(self, text="Select a Stock", font=("Arial", 18, "bold")).pack(pady=10)

        list_frame = ttk.Frame(self)
        list_frame.pack(pady=10, fill="both", expand=True)

        self.ticker_list = tk.Listbox(list_frame, height=18)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.ticker_list.yview)
        self.ticker_list.configure(yscrollcommand=scrollbar.set)

        self.ticker_list.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        button_row = ttk.Frame(self)
        button_row.pack(pady=10)

        ttk.Button(button_row, text="Use Selected", command=self.use_selected).grid(
            row=0, column=0, padx=10
        )
        ttk.Button(
            button_row,
            text="Back to Main Menu",
            command=lambda: controller.show_frame("MainMenu"),
        ).grid(row=0, column=1, padx=10)

    def refresh(self) -> None:
        self.ticker_list.delete(0, tk.END)
        for ticker in self.controller.state.tickers:
            self.ticker_list.insert(tk.END, ticker)
        if self.controller.state.selected_ticker in self.controller.state.tickers:
            index = self.controller.state.tickers.index(self.controller.state.selected_ticker)
            self.ticker_list.selection_set(index)
            self.ticker_list.see(index)

    def use_selected(self) -> None:
        selection = self.ticker_list.curselection()
        if not selection:
            messagebox.showinfo("Select a ticker", "Please select a ticker from the list.")
            return
        ticker = self.ticker_list.get(selection[0])
        self.controller.state.selected_ticker = ticker
        self.controller.persist_state()
        self.controller.show_frame("AnalysisPage")


class AnalysisPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller
        self.api_client: MassiveApiClient | None = None
        self.option_contract: dict | None = None
        self.scroll_canvas = tk.Canvas(self, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.scroll_canvas.yview)
        self.scroll_canvas.configure(yscrollcommand=self.scrollbar.set)
        self.scroll_canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.content_frame = ttk.Frame(self.scroll_canvas)
        self.scroll_window = self.scroll_canvas.create_window(
            (0, 0), window=self.content_frame, anchor="nw"
        )
        self.content_frame.bind("<Configure>", self._on_content_configure)
        self.scroll_canvas.bind("<Configure>", self._on_canvas_configure)
        self.scroll_canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        stock_frame = ttk.LabelFrame(self.content_frame, text="Stock Analysis")
        stock_frame.pack(pady=10, fill="both", expand=True, padx=40)

        chart_header = ttk.Label(
            stock_frame,
            text="Current (or previous trading day) chart",
            font=("Arial", 12, "bold"),
        )
        chart_header.pack(pady=(10, 5))

        chart_frame = ttk.Frame(stock_frame)
        chart_frame.pack(pady=5, fill="both", expand=True)

        self.chart_canvas = tk.Canvas(chart_frame, height=180, bg="#f0f0f0")
        self.chart_canvas.pack(fill="both", expand=True, padx=10, pady=10)
        self.chart_canvas.create_text(
            220,
            110,
            text="Daily chart preview will render here.",
            fill="#666",
        )

        slider_frame = ttk.Frame(stock_frame)
        slider_frame.pack(fill="x", padx=20, pady=(5, 10))

        ttk.Label(slider_frame, text="Time Horizon").grid(
            row=0, column=0, columnspan=len(HORIZON_CONFIGS), sticky="w"
        )
        self.horizon_var = tk.IntVar(value=1)
        self.horizon_slider = tk.Scale(
            slider_frame,
            from_=0,
            to=len(HORIZON_CONFIGS) - 1,
            orient="horizontal",
            variable=self.horizon_var,
            resolution=1,
            showvalue=False,
            command=self._snap_horizon,
            length=600,
        )
        self.horizon_slider.grid(
            row=1, column=0, columnspan=len(HORIZON_CONFIGS), sticky="ew", pady=5
        )
        for index in range(len(HORIZON_CONFIGS)):
            slider_frame.columnconfigure(index, weight=1)

        labels_frame = ttk.Frame(slider_frame)
        labels_frame.grid(row=2, column=0, columnspan=len(HORIZON_CONFIGS), sticky="ew")
        for index, (label, _days, _minutes, cadence_label) in enumerate(HORIZON_CONFIGS):
            ttk.Label(labels_frame, text=f"{label}\n({cadence_label})").grid(
                row=0, column=index, padx=4
            )
            labels_frame.columnconfigure(index, weight=1)

        self.stock_info_frame = ttk.LabelFrame(stock_frame, text="Stock Snapshot")
        self.stock_info_frame.pack(padx=20, pady=(5, 15), fill="x")
        self.stock_values: dict[str, ttk.Label] = {}
        self._build_info_grid(
            self.stock_info_frame,
            [
                ("Price", "price"),
                ("Previous Close", "prev_close"),
                ("Open", "open"),
                ("High", "high"),
                ("Low", "low"),
                ("Volume", "volume"),
                ("Market Cap", "market_cap"),
                ("52 Week Range", "range_52w"),
            ],
            self.stock_values,
            columns=4,
        )

        self.option_info_frame = ttk.LabelFrame(stock_frame, text="Option Snapshot")
        self.option_info_frame.pack(padx=20, pady=(5, 15), fill="x")
        self.option_values: dict[str, ttk.Label] = {}
        self._build_info_grid(
            self.option_info_frame,
            [
                ("Contract", "contract"),
                ("Expiration", "expiration"),
                ("Type", "type"),
                ("Strike", "strike"),
                ("Option Price", "price"),
            ],
            self.option_values,
            columns=4,
        )

        self.options_frame = ttk.LabelFrame(stock_frame, text="Option Contracts")
        self.options_frame.pack(padx=20, pady=(5, 15), fill="x")
        self.options_frame.columnconfigure(0, weight=1)
        self.options_frame.columnconfigure(1, weight=0)
        self.options_frame.rowconfigure(0, weight=1)

        self.option_records: list[dict] = []
        self.all_option_records: list[dict] = []
        list_frame = ttk.Frame(self.options_frame)
        list_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=8)
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        filter_frame = ttk.Frame(self.options_frame)
        filter_frame.grid(row=0, column=1, sticky="nsew", padx=(0, 10), pady=8)
        filter_frame.columnconfigure(1, weight=1)

        ttk.Label(filter_frame, text="Expiration").grid(
            row=0, column=0, padx=5, pady=2, sticky="w"
        )
        self.expiration_var = tk.StringVar(value="All")
        self.expiration_dropdown = ttk.Combobox(
            filter_frame, textvariable=self.expiration_var, state="readonly", width=18
        )
        self.expiration_dropdown.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        self.expiration_dropdown.bind("<<ComboboxSelected>>", self.on_option_filter_change)

        ttk.Label(filter_frame, text="Strike").grid(
            row=1, column=0, padx=5, pady=2, sticky="w"
        )
        self.strike_var = tk.StringVar(value="All")
        self.strike_dropdown = ttk.Combobox(
            filter_frame, textvariable=self.strike_var, state="readonly", width=12
        )
        self.strike_dropdown.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        self.strike_dropdown.bind("<<ComboboxSelected>>", self.on_option_filter_change)

        ttk.Label(filter_frame, text="Type").grid(row=2, column=0, padx=5, pady=2, sticky="w")
        self.type_var = tk.StringVar(value="All")
        self.type_dropdown = ttk.Combobox(
            filter_frame, textvariable=self.type_var, state="readonly", width=10
        )
        self.type_dropdown.grid(row=2, column=1, padx=5, pady=2, sticky="ew")
        self.type_dropdown.bind("<<ComboboxSelected>>", self.on_option_filter_change)

        self.options_list = tk.Listbox(list_frame, height=8, width=42)
        options_scroll = ttk.Scrollbar(
            list_frame, orient="vertical", command=self.options_list.yview
        )
        self.options_list.configure(yscrollcommand=options_scroll.set)
        self.options_list.grid(row=0, column=0, sticky="nsew")
        options_scroll.grid(row=0, column=1, sticky="ns")
        self.options_list.bind("<<ListboxSelect>>", self.on_option_select)

        self.greeks_frame = ttk.LabelFrame(stock_frame, text="Option Greeks")
        self.greeks_frame.pack(padx=20, pady=(5, 15), fill="x")
        self.greeks_values: dict[str, ttk.Label] = {}
        self._build_info_grid(
            self.greeks_frame,
            [
                ("Delta", "delta"),
                ("Gamma", "gamma"),
                ("Theta", "theta"),
                ("Vega", "vega"),
                ("Rho", "rho"),
                ("IV", "iv"),
            ],
            self.greeks_values,
            columns=3,
        )

        self.strategy_var = tk.StringVar(value=self.controller.state.option_strategy)
        self.strategy_dropdown = ttk.Combobox(
            filter_frame,
            textvariable=self.strategy_var,
            values=[
                "Naked Call",
                "Naked Put",
                "Vertical Spread",
                "Calendar Spread",
            ],
            state="readonly",
            width=20,
        )
        self.strategy_dropdown.bind("<<ComboboxSelected>>", self.on_strategy_change)

        ttk.Label(filter_frame, text="Option Strategy:").grid(
            row=3, column=0, padx=5, pady=(8, 2), sticky="w"
        )
        self.strategy_dropdown.grid(row=3, column=1, padx=5, pady=(8, 2), sticky="ew")

        ttk.Button(filter_frame, text="Go", command=self.go_to_strategy).grid(
            row=4, column=0, columnspan=2, padx=5, pady=(8, 2), sticky="ew"
        )

        button_row = ttk.Frame(self.content_frame)
        button_row.pack(pady=10)

        ttk.Button(button_row, text="Save Analysis", command=self.save_analysis).grid(
            row=0, column=0, padx=10
        )
        ttk.Button(
            button_row, text="Load Data", command=self.load_market_data
        ).grid(row=0, column=1, padx=10)
        ttk.Button(
            button_row,
            text="Select Stock",
            command=lambda: controller.show_frame("TickerSelectPage"),
        ).grid(row=0, column=2, padx=10)
        ttk.Button(
            button_row,
            text="Back to Main Menu",
            command=lambda: controller.show_frame("MainMenu"),
        ).grid(row=0, column=3, padx=10)
        ttk.Label(button_row, text="Analysis Mode:").grid(row=0, column=4, padx=(20, 5))
        self.analysis_mode_var = tk.StringVar(value=self.controller.state.analysis_mode)
        self.analysis_mode_dropdown = ttk.Combobox(
            button_row,
            textvariable=self.analysis_mode_var,
            values=["Stock Analysis", "Option Analysis"],
            state="readonly",
            width=20,
        )
        self.analysis_mode_dropdown.grid(row=0, column=5, padx=5)
        self.analysis_mode_dropdown.bind("<<ComboboxSelected>>", self.on_analysis_mode_change)

    def _snap_horizon(self, value: str) -> None:
        snapped = int(round(float(value)))
        self.horizon_var.set(snapped)
        self.horizon_slider.set(snapped)

    def _build_info_grid(
        self,
        parent: ttk.Frame,
        rows: list[tuple[str, str]],
        target: dict[str, ttk.Label],
        columns: int = 1,
    ) -> None:
        for item_index, (label, key) in enumerate(rows):
            row_index = item_index // columns
            column_index = (item_index % columns) * 2
            ttk.Label(parent, text=label).grid(
                row=row_index, column=column_index, padx=10, pady=4, sticky="w"
            )
            value_label = ttk.Label(parent, text="--", foreground="#b00020")
            value_label.grid(
                row=row_index, column=column_index + 1, padx=10, pady=4, sticky="w"
            )
            target[key] = value_label
        for index in range(columns * 2):
            parent.columnconfigure(index, weight=1)

    def _format_float(self, value: float) -> str:
        decimals = 2 if abs(value) >= 1 else 4
        multiplier = 10**decimals
        truncated = math.trunc(value * multiplier) / multiplier
        return f"{truncated:.{decimals}f}".rstrip("0").rstrip(".")

    def _set_value(self, label: ttk.Label, value: str | int | float | None) -> None:
        if value in (None, "", "--"):
            label.config(text="--", foreground="#b00020")
        else:
            if isinstance(value, float):
                text = self._format_float(value)
            elif isinstance(value, int):
                text = str(value)
            else:
                text = str(value)
            label.config(text=text, foreground="#0a7a2f")

    def _render_chart(self, aggregates: list[dict]) -> None:
        self.chart_canvas.delete("all")
        points_raw: list[tuple[float, int]] = []
        for item in aggregates:
            try:
                close_value = float(item.get("c"))
                timestamp = int(item.get("t"))
            except (TypeError, ValueError):
                continue
            points_raw.append((close_value, timestamp))
        if not points_raw:
            self.chart_canvas.create_text(
                220,
                110,
                text="No chart data available for this range.",
                fill="#666",
            )
            return
        if len(points_raw) < 2:
            self.chart_canvas.update_idletasks()
            width = max(self.chart_canvas.winfo_width(), 1)
            height = max(self.chart_canvas.winfo_height(), 1)
            padding = 20
            x = width / 2
            y = height / 2
            self.chart_canvas.create_oval(
                x - 4,
                y - 4,
                x + 4,
                y + 4,
                fill="#1f77b4",
                outline="",
            )
            self.chart_canvas.create_text(
                padding,
                padding / 2,
                anchor="w",
                text=f"{points_raw[0][0]:.2f}",
                fill="#1f77b4",
            )
            return
        self.chart_canvas.update_idletasks()
        width = max(self.chart_canvas.winfo_width(), 1)
        height = max(self.chart_canvas.winfo_height(), 1)
        padding_left = 60
        padding_right = 20
        padding_top = 20
        padding_bottom = 30
        min_price = min(price for price, _ts in points_raw)
        max_price = max(price for price, _ts in points_raw)
        price_span = max(max_price - min_price, 1e-6)
        x_span = max(len(points_raw) - 1, 1)

        points = []
        for idx, (price, _ts) in enumerate(points_raw):
            x = padding_left + (width - padding_left - padding_right) * (idx / x_span)
            y = height - padding_bottom - (
                height - padding_top - padding_bottom
            ) * ((price - min_price) / price_span)
            points.extend([x, y])

        if len(points) < 4:
            self.chart_canvas.create_text(
                220,
                110,
                text="Not enough chart data to render a line.",
                fill="#666",
            )
            return

        try:
            if len(points) < 4:
                raise tk.TclError("Insufficient points for line rendering.")
            self.chart_canvas.create_line(*points, fill="#1f77b4", width=2, smooth=True)
        except tk.TclError:
            self.chart_canvas.create_text(
                220,
                110,
                text="Unable to render chart line for this data.",
                fill="#666",
            )
            return
        grid_color = "#d9d9d9"
        axis_color = "#444"
        for step in range(5):
            fraction = step / 4
            y = height - padding_bottom - (
                height - padding_top - padding_bottom
            ) * fraction
            self.chart_canvas.create_line(
                padding_left, y, width - padding_right, y, fill=grid_color
            )
            value = min_price + (price_span * fraction)
            self.chart_canvas.create_text(
                padding_left - 8,
                y,
                anchor="e",
                text=f"{value:.2f}",
                fill=axis_color,
            )

        self.chart_canvas.create_line(
            padding_left, padding_top, padding_left, height - padding_bottom, fill=axis_color
        )
        self.chart_canvas.create_line(
            padding_left,
            height - padding_bottom,
            width - padding_right,
            height - padding_bottom,
            fill=axis_color,
        )

        total_points = len(points_raw)
        tick_count = min(5, total_points)
        for tick_index in range(tick_count):
            idx = int(round(tick_index * (total_points - 1) / max(tick_count - 1, 1)))
            _price, ts = points_raw[idx]
            x = padding_left + (width - padding_left - padding_right) * (idx / x_span)
            dt = datetime.fromtimestamp(ts / 1000)
            label = dt.strftime("%m/%d")
            self.chart_canvas.create_text(
                x,
                height - padding_bottom + 12,
                anchor="n",
                text=label,
                fill=axis_color,
            )

    def _format_http_error_detail(self, exc: HTTPError) -> str:
        return format_http_error_detail(exc)

    def _show_api_error(self, exc: HTTPError, service: str, hint: str | None = None) -> None:
        detail = self._format_http_error_detail(exc)
        detail_msg = f"\nDetails: {detail}" if detail else ""
        hint_msg = f"\n{hint}" if hint else ""
        self._show_error_dialog(
            "API Error",
            f"{service} API returned an error: {exc.code} {exc.reason}.{detail_msg}{hint_msg}",
        )

    def _show_error_dialog(self, title: str, message: str) -> None:
        dialog = tk.Toplevel(self)
        dialog.title(title)
        dialog.geometry("620x320")
        dialog.transient(self)

        dialog.rowconfigure(0, weight=1)
        dialog.columnconfigure(0, weight=1)

        text_frame = ttk.Frame(dialog)
        text_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

        text_widget = tk.Text(text_frame, wrap="word")
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        text_widget.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        text_widget.insert("1.0", message)
        text_widget.focus_set()

        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=1, column=0, pady=(0, 10))

        def copy_to_clipboard() -> None:
            dialog.clipboard_clear()
            dialog.clipboard_append(text_widget.get("1.0", "end-1c"))
            dialog.update_idletasks()

        ttk.Button(button_frame, text="Copy", command=copy_to_clipboard).grid(
            row=0, column=0, padx=5
        )
        ttk.Button(button_frame, text="Close", command=dialog.destroy).grid(
            row=0, column=1, padx=5
        )

    def _on_content_configure(self, _event: tk.Event) -> None:
        self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all"))

    def _on_canvas_configure(self, event: tk.Event) -> None:
        self.scroll_canvas.itemconfigure(self.scroll_window, width=event.width)

    def _on_mousewheel(self, event: tk.Event) -> None:
        if self.scroll_canvas.winfo_height() < self.content_frame.winfo_height():
            self.scroll_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all"))

    def _sync_option_snapshot(self) -> None:
        contract = self.option_contract or {}
        self._set_value(self.option_values["contract"], contract.get("ticker"))
        self._set_value(self.option_values["expiration"], contract.get("expiration_date"))
        contract_type = contract.get("contract_type")
        display_type = contract_type.upper() if contract_type else None
        self._set_value(self.option_values["type"], display_type)
        self._set_value(self.option_values["strike"], contract.get("strike_price"))
        self._set_value(self.option_values["price"], option_mid_price(contract))

    def _toggle_info_panels(self) -> None:
        is_stock = self.analysis_mode_var.get() == "Stock Analysis"

        if not self.stock_info_frame.winfo_ismapped():
            self.stock_info_frame.pack(padx=20, pady=(5, 15), fill="x")

        if is_stock:
            self.option_info_frame.pack_forget()
            self.options_frame.pack_forget()
            self.greeks_frame.pack_forget()
        else:
            if not self.option_info_frame.winfo_ismapped():
                self.option_info_frame.pack(padx=20, pady=(5, 15), fill="x")
            if not self.options_frame.winfo_ismapped():
                self.options_frame.pack(padx=20, pady=(5, 15), fill="x")
            if not self.greeks_frame.winfo_ismapped():
                self.greeks_frame.pack(padx=20, pady=(5, 15), fill="x")
        self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all"))

    def _get_filter_value(self, var: tk.StringVar) -> str | None:
        value = var.get()
        return None if value == "All" else value

    def _record_matches_filters(self, record: dict, filters: dict[str, str | None]) -> bool:
        expiration = record.get("expiration_date")
        strike = format_strike(record.get("strike_price"))
        contract_type = normalize_contract_type(record.get("contract_type"))
        if filters.get("expiration") and filters["expiration"] != expiration:
            return False
        if filters.get("strike") and filters["strike"] != strike:
            return False
        if filters.get("type") and filters["type"] != contract_type:
            return False
        return True

    def _compute_filter_options(
        self, records: list[dict], current: dict[str, str | None]
    ) -> dict[str, list[str]]:
        options: dict[str, set[str]] = {"expiration": set(), "strike": set(), "type": set()}
        for record in records:
            expiration = record.get("expiration_date")
            strike = format_strike(record.get("strike_price"))
            contract_type = normalize_contract_type(record.get("contract_type"))
            if self._record_matches_filters(record, {**current, "expiration": None}):
                if expiration:
                    options["expiration"].add(expiration)
            if self._record_matches_filters(record, {**current, "strike": None}):
                if strike:
                    options["strike"].add(strike)
            if self._record_matches_filters(record, {**current, "type": None}):
                if contract_type:
                    options["type"].add(contract_type)
        return {
            "expiration": sorted(options["expiration"]),
            "strike": sorted(
                options["strike"], key=lambda value: float(value) if value.replace(".", "", 1).isdigit() else value
            ),
            "type": sorted(options["type"]),
        }

    def _refresh_option_filters(self, reset: bool = False) -> None:
        if reset:
            self.expiration_var.set("All")
            self.strike_var.set("All")
            self.type_var.set("All")
        filters = {
            "expiration": self._get_filter_value(self.expiration_var),
            "strike": self._get_filter_value(self.strike_var),
            "type": self._get_filter_value(self.type_var),
        }
        options = self._compute_filter_options(self.all_option_records, filters)
        for key, dropdown, var in (
            ("expiration", self.expiration_dropdown, self.expiration_var),
            ("strike", self.strike_dropdown, self.strike_var),
            ("type", self.type_dropdown, self.type_var),
        ):
            values = ["All"] + options[key]
            dropdown["values"] = values
            if var.get() not in values:
                var.set("All")
        self._apply_option_filters()

    def _apply_option_filters(self) -> None:
        filters = {
            "expiration": self._get_filter_value(self.expiration_var),
            "strike": self._get_filter_value(self.strike_var),
            "type": self._get_filter_value(self.type_var),
        }
        self.option_records = [
            record
            for record in self.all_option_records
            if self._record_matches_filters(record, filters)
        ]
        self.options_list.delete(0, tk.END)
        if not self.option_records:
            self.options_list.insert(tk.END, "No option contracts returned.")
            self.option_contract = None
        else:
            for contract in self.option_records:
                self.options_list.insert(
                    tk.END,
                    "{ticker} {expiration} {type} {strike}".format(
                        ticker=contract.get("ticker", "--"),
                        expiration=contract.get("expiration_date", "--"),
                        type=str(contract.get("contract_type", "--")).upper(),
                        strike=contract.get("strike_price", "--"),
                    ),
                )
            self.options_list.selection_set(0)
            self.options_list.see(0)
            self.option_contract = self.option_records[0]
        self._sync_option_snapshot()
        self._sync_greeks()

    def refresh(self) -> None:
        self.controller.state.analysis_mode = "Option Analysis"
        self.analysis_mode_var.set("Option Analysis")
        self.strategy_var.set(self.controller.state.option_strategy)
        api_key = load_api_key()
        self.api_client = MassiveApiClient(api_key) if api_key else None
        self._toggle_info_panels()
        self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all"))
        self.after(0, self.load_market_data)

    def on_analysis_mode_change(self, _event: object) -> None:
        self.controller.state.analysis_mode = self.analysis_mode_var.get()
        self.controller.persist_state()
        self._toggle_info_panels()
        self._sync_option_snapshot()

    def on_strategy_change(self, _event: object) -> None:
        self.controller.state.option_strategy = self.strategy_var.get()
        self.controller.persist_state()

    def go_to_strategy(self) -> None:
        strategy = self.strategy_var.get()
        self.controller.state.option_strategy = strategy
        self.controller.persist_state()
        if strategy in ("Naked Call", "Naked Put"):
            self.controller.show_frame("CallPutAnalysisPage")
        else:
            self.controller.show_frame("SpreadAnalysisPage")

    def load_market_data(self) -> None:
        if not self.api_client:
            messagebox.showinfo(
                "Missing key", "Enter or set a Massive API key to load stock data."
            )
            return
        ticker = self.controller.state.selected_ticker
        if not ticker:
            messagebox.showinfo("Missing ticker", "Select a ticker first.")
            return
        horizon_index = int(round(self.horizon_var.get()))
        horizon_index = min(max(horizon_index, 0), len(HORIZON_CONFIGS) - 1)
        _label, days_back, minutes_per_bar, _cadence_label = HORIZON_CONFIGS[horizon_index]
        cache_payload = load_cached_market_data(ticker) or {}
        cache_date = cache_payload.get("last_updated")
        today_label = effective_market_date().isoformat()
        aggregates_map = cache_payload.get("aggregates", {})
        cached_stock = cache_payload.get("stock")
        cached_options = cache_payload.get("options")
        cached_aggregates = aggregates_map.get(str(horizon_index))
        should_fetch = cache_date != today_label
        if cached_stock is None or cached_options is None or cached_aggregates is None:
            should_fetch = True

        if should_fetch:
            try:
                stock_data = self.api_client.fetch_previous_close(ticker)
            except HTTPError as exc:
                self._show_api_error(exc, "Massive", "Verify your Massive API key.")
                return
            except URLError as exc:
                self._show_error_dialog(
                    "Connection Error",
                    f"Could not reach Massive API endpoint: {exc.reason}",
                )
                return
            try:
                option_data = self.api_client.fetch_option_snapshots(ticker)
            except HTTPError as exc:
                self._show_api_error(
                    exc,
                    "Massive",
                    "Verify your Massive API key and options data entitlements.",
                )
                return
            except URLError as exc:
                self._show_error_dialog(
                    "Connection Error",
                    f"Could not reach Massive API endpoint: {exc.reason}",
                )
                return
            try:
                aggregates = self.api_client.fetch_aggregates(
                    ticker, days_back, minutes_per_bar
                )
            except HTTPError as exc:
                self._show_api_error(exc, "Massive", "Verify your Massive API key.")
                return
            except URLError as exc:
                self._show_error_dialog(
                    "Connection Error",
                    f"Could not reach Massive API endpoint: {exc.reason}",
                )
                return
            aggregates_map[str(horizon_index)] = aggregates
            option_records = normalize_option_records(option_data)
            cache_payload.update(
                {
                    "last_updated": today_label,
                    "stock": stock_data,
                    "options": option_records,
                    "aggregates": aggregates_map,
                }
            )
            save_cached_market_data(ticker, cache_payload)
        else:
            stock_data = cached_stock or {}
            option_records = normalize_option_records(cached_options or [])
            aggregates = cached_aggregates or []

        self._set_value(self.stock_values["price"], stock_data.get("close"))
        self._set_value(self.stock_values["prev_close"], stock_data.get("close"))
        self._set_value(self.stock_values["open"], stock_data.get("open"))
        self._set_value(self.stock_values["high"], stock_data.get("high"))
        self._set_value(self.stock_values["low"], stock_data.get("low"))
        self._set_value(self.stock_values["volume"], stock_data.get("volume"))
        self._set_value(self.stock_values["market_cap"], "--")
        self._set_value(self.stock_values["range_52w"], "--")
        self.option_contract = option_records[0] if option_records else None
        self._sync_option_snapshot()

        self._render_chart(aggregates)

        self.all_option_records = option_records
        self._refresh_option_filters(reset=True)

    def save_analysis(self) -> None:
        self.controller.state.analysis_mode = self.analysis_mode_var.get()
        self.controller.state.option_strategy = self.strategy_var.get()
        self.controller.persist_state()
        self._sync_option_snapshot()
        messagebox.showinfo("Saved", "Analysis settings saved locally.")

    def on_option_select(self, _event: object) -> None:
        selection = self.options_list.curselection()
        if not selection:
            return
        index = selection[0]
        if index >= len(self.option_records):
            return
        self.option_contract = self.option_records[index]
        self._sync_option_snapshot()
        self._sync_greeks()

    def on_option_filter_change(self, _event: object) -> None:
        self._refresh_option_filters()

    def _sync_greeks(self) -> None:
        greeks = extract_greeks(self.option_contract or {})
        self._set_value(self.greeks_values["delta"], greeks.get("delta"))
        self._set_value(self.greeks_values["gamma"], greeks.get("gamma"))
        self._set_value(self.greeks_values["theta"], greeks.get("theta"))
        self._set_value(self.greeks_values["vega"], greeks.get("vega"))
        self._set_value(self.greeks_values["rho"], greeks.get("rho"))
        self._set_value(self.greeks_values["iv"], greeks.get("iv"))


class GeneralAnalysisPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller
        self.api_client: MassiveApiClient | None = None
        self._rate_lock = threading.Lock()
        self._last_request_time = 0.0
        self._min_request_interval = 0.05
        self._grouped_ticker_pattern = re.compile(r"^[A-Z0-9]+$")

        header = ttk.Label(self, text="General Analysis", font=("Arial", 18, "bold"))
        header.pack(pady=10)

        description = ttk.Label(
            self,
            text=(
                "Run cross-sectional, time-series, and factor analysis across your current "
                "stock universe. Tune parameters and export results to a text file."
            ),
            wraplength=700,
            justify="center",
        )
        description.pack(pady=(0, 15))

        form_frame = ttk.LabelFrame(self, text="Analysis Settings")
        form_frame.pack(padx=40, pady=10, fill="x")
        form_frame.columnconfigure(1, weight=1)

        ttk.Label(form_frame, text="Analysis Type").grid(
            row=0, column=0, padx=10, pady=6, sticky="w"
        )
        self.analysis_type_var = tk.StringVar()
        self.analysis_type_dropdown = ttk.Combobox(
            form_frame,
            textvariable=self.analysis_type_var,
            values=[
                "Cross-Sectional",
                "Time-Series",
                "Cross-Sectional + Time-Series",
            ],
            state="readonly",
            width=30,
        )
        self.analysis_type_dropdown.grid(row=0, column=1, padx=10, pady=6, sticky="ew")

        ttk.Label(form_frame, text="Cross-Sectional Strategy").grid(
            row=1, column=0, padx=10, pady=6, sticky="w"
        )
        self.cross_sectional_var = tk.StringVar()
        working_strategies = ["Momentum", "Low Volatility", "Liquidity", "Size"]
        non_working = [
            "Value",
            "Quality",
            "Investment",
            "Earnings Momentum",
            "Carry / Yield",
        ]
        self.strategy_label_to_key: dict[str, str] = {
            name: name for name in working_strategies
        }
        self.strategy_label_to_key.update(
            {f"{name} (no work)": name for name in non_working}
        )
        self.strategy_key_to_label = {
            key: label for label, key in self.strategy_label_to_key.items()
        }
        dropdown_values = list(self.strategy_label_to_key.keys())
        self.cross_sectional_dropdown = ttk.Combobox(
            form_frame,
            textvariable=self.cross_sectional_var,
            values=dropdown_values,
            state="readonly",
            width=30,
        )
        self.cross_sectional_dropdown.grid(row=1, column=1, padx=10, pady=6, sticky="ew")
        self.cross_sectional_dropdown.bind("<<ComboboxSelected>>", self._on_strategy_change)

        self.strategy_detail_var = tk.StringVar(value="Required data: prices")
        self.strategy_detail_label = ttk.Label(form_frame, textvariable=self.strategy_detail_var)
        self.strategy_detail_label.grid(row=2, column=0, columnspan=2, padx=10, pady=2, sticky="w")

        ttk.Label(form_frame, text="Lookback (days)").grid(
            row=3, column=0, padx=10, pady=6, sticky="w"
        )
        self.lookback_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.lookback_var).grid(
            row=3, column=1, padx=10, pady=6, sticky="ew"
        )

        ttk.Label(form_frame, text="Skip (days)").grid(
            row=4, column=0, padx=10, pady=6, sticky="w"
        )
        self.skip_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.skip_var).grid(
            row=4, column=1, padx=10, pady=6, sticky="ew"
        )

        ttk.Label(form_frame, text="Top Quantile (0-1)").grid(
            row=5, column=0, padx=10, pady=6, sticky="w"
        )
        self.top_quantile_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.top_quantile_var).grid(
            row=5, column=1, padx=10, pady=6, sticky="ew"
        )

        ttk.Label(form_frame, text="Bottom Quantile (0-1)").grid(
            row=6, column=0, padx=10, pady=6, sticky="w"
        )
        self.bottom_quantile_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.bottom_quantile_var).grid(
            row=6, column=1, padx=10, pady=6, sticky="ew"
        )

        self.momentum_label = ttk.Label(form_frame, text="Momentum Toggles")
        self.momentum_label.grid(row=7, column=0, padx=10, pady=6, sticky="w")
        self.momentum_toggles_frame = ttk.Frame(form_frame)
        self.momentum_toggles_frame.grid(row=7, column=1, padx=10, pady=6, sticky="w")
        self.momentum_volatility_var = tk.BooleanVar()
        self.momentum_residual_var = tk.BooleanVar()
        self.momentum_multi_horizon_var = tk.BooleanVar()
        self.momentum_volatility_check = ttk.Checkbutton(
            self.momentum_toggles_frame,
            text="Volatility Scaling",
            variable=self.momentum_volatility_var,
        )
        self.momentum_residual_check = ttk.Checkbutton(
            self.momentum_toggles_frame,
            text="Residual Momentum",
            variable=self.momentum_residual_var,
        )
        self.momentum_multi_horizon_check = ttk.Checkbutton(
            self.momentum_toggles_frame,
            text="Multi-Horizon",
            variable=self.momentum_multi_horizon_var,
        )
        self.momentum_volatility_check.grid(row=0, column=0, padx=5, sticky="w")
        self.momentum_residual_check.grid(row=0, column=1, padx=5, sticky="w")
        self.momentum_multi_horizon_check.grid(row=0, column=2, padx=5, sticky="w")

        ttk.Label(form_frame, text="Output Directory").grid(
            row=8, column=0, padx=10, pady=6, sticky="w"
        )
        self.output_dir_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.output_dir_var).grid(
            row=8, column=1, padx=10, pady=6, sticky="ew"
        )

        button_row = ttk.Frame(self)
        button_row.pack(pady=15)

        ttk.Button(button_row, text="Run Analysis", command=self.run_analysis).grid(
            row=0, column=0, padx=10
        )
        ttk.Button(
            button_row,
            text="Back to Main Menu",
            command=lambda: controller.show_frame("MainMenu"),
        ).grid(row=0, column=1, padx=10)

        output_frame = ttk.LabelFrame(self, text="Latest Analysis Output")
        output_frame.pack(padx=40, pady=(5, 15), fill="both", expand=True)
        output_frame.rowconfigure(0, weight=1)
        output_frame.columnconfigure(0, weight=1)
        self.output_text = tk.Text(output_frame, height=12, wrap="word")
        output_scrollbar = ttk.Scrollbar(
            output_frame, orient="vertical", command=self.output_text.yview
        )
        self.output_text.configure(yscrollcommand=output_scrollbar.set)
        self.output_text.grid(row=0, column=0, sticky="nsew")
        output_scrollbar.grid(row=0, column=1, sticky="ns")

    def refresh(self) -> None:
        settings = dict(DEFAULT_GENERAL_ANALYSIS_SETTINGS)
        settings.update(self.controller.state.general_analysis_settings or {})
        self.analysis_type_var.set(settings.get("analysis_type", "Cross-Sectional"))
        strategy_key = settings.get("cross_sectional_strategy", "Momentum")
        self.cross_sectional_var.set(
            self.strategy_key_to_label.get(strategy_key, "Momentum")
        )
        self._update_strategy_detail()
        self.lookback_var.set(str(settings.get("lookback_days", 90)))
        self.skip_var.set(str(settings.get("skip_days", 5)))
        self.top_quantile_var.set(str(settings.get("top_quantile", 0.2)))
        self.bottom_quantile_var.set(str(settings.get("bottom_quantile", 0.2)))
        self.momentum_volatility_var.set(settings.get("momentum_use_volatility_scaling", False))
        self.momentum_residual_var.set(settings.get("momentum_use_residual", False))
        self.momentum_multi_horizon_var.set(settings.get("momentum_use_multi_horizon", False))
        self.output_dir_var.set(settings.get("output_dir", str(ANALYSIS_OUTPUT_DIR)))

    def _on_strategy_change(self, _event: object) -> None:
        self._update_strategy_detail()

    def _update_strategy_detail(self) -> None:
        strategy = self._selected_strategy_key()
        if strategy == "Momentum":
            self.strategy_detail_var.set("Required data: prices (with optional volume)")
            self._set_momentum_toggles_state(enabled=True)
            return
        spec = STRATEGY_REGISTRY.get(strategy)
        if spec:
            required = ", ".join(spec.required_data)
            self.strategy_detail_var.set(f"Required data: {required}")
        else:
            self.strategy_detail_var.set("Required data: prices")
        self._set_momentum_toggles_state(enabled=False)

    def _set_momentum_toggles_state(self, *, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        if not enabled:
            self.momentum_volatility_var.set(False)
            self.momentum_residual_var.set(False)
            self.momentum_multi_horizon_var.set(False)
            self.momentum_label.grid_remove()
            self.momentum_toggles_frame.grid_remove()
        else:
            self.momentum_label.grid()
            self.momentum_toggles_frame.grid()
        for widget in (
            self.momentum_volatility_check,
            self.momentum_residual_check,
            self.momentum_multi_horizon_check,
        ):
            widget.configure(state=state)

    def run_analysis(self) -> None:
        if not self.controller.api_key:
            messagebox.showinfo("Missing key", "Enter a Massive API key first.")
            return
        if not self.controller.state.tickers:
            messagebox.showinfo("Missing universe", "Please add tickers first.")
            return
        lookback_days = parse_int(self.lookback_var.get())
        skip_days = parse_int(self.skip_var.get())
        top_quantile = parse_float(self.top_quantile_var.get())
        bottom_quantile = parse_float(self.bottom_quantile_var.get())
        if lookback_days is None or lookback_days <= 0:
            messagebox.showinfo("Invalid input", "Lookback days must be a positive integer.")
            return
        if skip_days is None or skip_days < 0:
            messagebox.showinfo("Invalid input", "Skip days must be zero or a positive integer.")
            return
        if skip_days >= lookback_days:
            messagebox.showinfo(
                "Invalid input",
                "Skip days must be less than lookback days to compute momentum.",
            )
            return
        if top_quantile is None or not (0 < top_quantile <= 1):
            messagebox.showinfo("Invalid input", "Top quantile must be between 0 and 1.")
            return
        if bottom_quantile is None or not (0 < bottom_quantile <= 1):
            messagebox.showinfo("Invalid input", "Bottom quantile must be between 0 and 1.")
            return
        output_dir = self.output_dir_var.get().strip() or str(ANALYSIS_OUTPUT_DIR)

        strategy = self._selected_strategy_key()
        settings_payload = {
            "analysis_type": self.analysis_type_var.get(),
            "cross_sectional_strategy": strategy,
            "time_series_strategy": strategy,
            "lookback_days": lookback_days,
            "skip_days": skip_days,
            "top_quantile": top_quantile,
            "bottom_quantile": bottom_quantile,
            "momentum_use_volatility_scaling": self.momentum_volatility_var.get(),
            "momentum_use_residual": self.momentum_residual_var.get(),
            "momentum_use_multi_horizon": self.momentum_multi_horizon_var.get(),
            "time_series_use_volatility_scaling": self.momentum_volatility_var.get(),
            "time_series_use_residual": self.momentum_residual_var.get(),
            "time_series_use_multi_horizon": self.momentum_multi_horizon_var.get(),
            "time_series_use_zscore": False,
            "time_series_winsorize_sigma": None,
            "output_dir": output_dir,
        }
        self.controller.state.general_analysis_settings = settings_payload
        self.controller.persist_state()

        if self.api_client is None:
            self.api_client = MassiveApiClient(self.controller.api_key)

        as_of = effective_market_date().isoformat()
        universe = list(self.controller.state.tickers)
        price_history, fetch_skipped = self._collect_price_history(
            universe, lookback_days, skip_days
        )
        fundamentals_by_ticker, fundamentals_skipped = self._collect_fundamentals(
            universe, as_of
        )

        analysis_type = settings_payload["analysis_type"]
        reports: list[str] = []
        report_labels: list[str] = []
        if analysis_type in {"Cross-Sectional", "Cross-Sectional + Time-Series"}:
            cross_result = self._run_cross_sectional_analysis(
                strategy,
                price_history,
                fundamentals_by_ticker,
                fetch_skipped,
                fundamentals_skipped,
                lookback_days,
                skip_days,
                top_quantile,
                bottom_quantile,
                settings_payload,
                universe,
                as_of,
            )
            if cross_result is None:
                return
            reports.append(cross_result)
            report_labels.append("cross_sectional")
        if analysis_type in {"Time-Series", "Cross-Sectional + Time-Series"}:
            time_series_result = self._run_time_series_analysis(
                strategy,
                price_history,
                fundamentals_by_ticker,
                fetch_skipped,
                fundamentals_skipped,
                lookback_days,
                skip_days,
                top_quantile,
                bottom_quantile,
                settings_payload,
                universe,
                as_of,
            )
            if time_series_result is None:
                return
            reports.append(time_series_result)
            report_labels.append("time_series")

        combined_report = "\n\n".join(reports)
        self.output_text.delete("1.0", tk.END)
        self.output_text.insert(tk.END, combined_report)

        analysis_slug = "combined" if len(report_labels) > 1 else report_labels[0]
        output_path = self._write_report(combined_report, output_dir, strategy, analysis_slug)
        messagebox.showinfo(
            "Analysis complete",
            f"{analysis_type} {strategy.lower()} results written to:\n{output_path}",
        )

    def _selected_strategy_key(self) -> str:
        label = self.cross_sectional_var.get()
        return self.strategy_label_to_key.get(label, "Momentum")

    def _missing_data_guidance(self, missing_requirements: list[str]) -> str:
        suggestions: list[str] = []
        if "fundamentals" in missing_requirements:
            suggestions.append(
                "Fundamentals: import a fundamentals file (CSV/JSON) with fields like book value, earnings, dividends, and market cap."
            )
        if "market_cap" in missing_requirements:
            suggestions.append(
                "Market cap: compute from price * shares outstanding if shares data is available, or ingest a market-cap file."
            )
        if "earnings" in missing_requirements or "analyst_revisions" in missing_requirements:
            suggestions.append(
                "Earnings/revisions: ingest analyst estimates or earnings surprise data from a data vendor or your own CSV export."
            )
        if not suggestions:
            suggestions.append("Provide the missing data in a local file and wire it into the analysis pipeline.")
        return "\n".join(f"- {item}" for item in suggestions)

    def _run_cross_sectional_analysis(
        self,
        strategy: str,
        price_history: dict[str, list[dict]],
        fundamentals_by_ticker: dict[str, dict],
        fetch_skipped: dict[str, str],
        fundamentals_skipped: dict[str, str],
        lookback_days: int,
        skip_days: int,
        top_quantile: float,
        bottom_quantile: float,
        settings_payload: dict[str, object],
        universe: list[str],
        as_of: str,
    ) -> str | None:
        if strategy == "Momentum":
            momentum_settings = MomentumSettings(
                lookback_days=lookback_days,
                skip_days=skip_days,
                top_quantile=top_quantile,
                bottom_quantile=bottom_quantile,
                use_volatility_scaling=settings_payload["momentum_use_volatility_scaling"],
                use_residual=settings_payload["momentum_use_residual"],
                use_multi_horizon=settings_payload["momentum_use_multi_horizon"],
            )
            result = compute_cross_sectional_momentum(
                price_history, fundamentals_by_ticker, momentum_settings
            )
        else:
            spec = STRATEGY_REGISTRY.get(strategy, STRATEGY_REGISTRY["Value"])
            data_availability = {
                "fundamentals": bool(fundamentals_by_ticker),
                "market_cap": any(
                    payload.get("market_cap") is not None
                    for payload in fundamentals_by_ticker.values()
                ),
                "earnings": any(
                    payload.get("earnings_actual") is not None
                    for payload in fundamentals_by_ticker.values()
                ),
                "analyst_revisions": False,
            }
            missing_requirements = [
                requirement
                for requirement in spec.required_data
                if requirement not in {"prices", "volume"}
                and not data_availability.get(requirement, False)
            ]
            if missing_requirements:
                guidance = self._missing_data_guidance(missing_requirements)
                messagebox.showinfo(
                    "Missing data",
                    "This strategy requires additional data sources that are not yet wired: "
                    f"{', '.join(missing_requirements)}.\n\n{guidance}",
                )
                return None
            factor_settings = CrossSectionalSettings(
                top_quantile=top_quantile, bottom_quantile=bottom_quantile
            )
            result = spec.compute(price_history, fundamentals_by_ticker, factor_settings)
        result.skipped.update(fetch_skipped)
        result.skipped.update(fundamentals_skipped)

        report_title = f"Cross-Sectional {strategy} Report"
        return format_cross_sectional_report(
            title=report_title,
            as_of=as_of,
            universe=universe,
            settings=settings_payload,
            result=result,
        )

    def _run_time_series_analysis(
        self,
        strategy: str,
        price_history: dict[str, list[dict]],
        fundamentals_by_ticker: dict[str, dict],
        fetch_skipped: dict[str, str],
        fundamentals_skipped: dict[str, str],
        lookback_days: int,
        skip_days: int,
        top_quantile: float,
        bottom_quantile: float,
        settings_payload: dict[str, object],
        universe: list[str],
        as_of: str,
    ) -> str | None:
        if strategy == "Momentum":
            momentum_settings = TimeSeriesMomentumSettings(
                lookback_days=lookback_days,
                skip_days=skip_days,
                top_quantile=top_quantile,
                bottom_quantile=bottom_quantile,
                use_volatility_scaling=settings_payload["time_series_use_volatility_scaling"],
                use_residual=settings_payload["time_series_use_residual"],
                use_multi_horizon=settings_payload["time_series_use_multi_horizon"],
                use_zscore=settings_payload["time_series_use_zscore"],
                winsorize_sigma=settings_payload["time_series_winsorize_sigma"],
            )
            result = compute_time_series_momentum(
                price_history, fundamentals_by_ticker, momentum_settings
            )
        else:
            spec = TIME_SERIES_STRATEGY_REGISTRY.get(
                strategy, TIME_SERIES_STRATEGY_REGISTRY["Value"]
            )
            data_availability = {
                "fundamentals": bool(fundamentals_by_ticker),
                "market_cap": any(
                    payload.get("market_cap") is not None
                    for payload in fundamentals_by_ticker.values()
                ),
                "earnings": any(
                    payload.get("earnings_actual") is not None
                    for payload in fundamentals_by_ticker.values()
                ),
                "analyst_revisions": False,
            }
            missing_requirements = [
                requirement
                for requirement in spec.required_data
                if requirement not in {"prices", "volume"}
                and not data_availability.get(requirement, False)
            ]
            if missing_requirements:
                guidance = self._missing_data_guidance(missing_requirements)
                messagebox.showinfo(
                    "Missing data",
                    "This strategy requires additional data sources that are not yet wired: "
                    f"{', '.join(missing_requirements)}.\n\n{guidance}",
                )
                return None
            factor_settings = TimeSeriesSettings(
                lookback_days=lookback_days,
                skip_days=skip_days,
                top_quantile=top_quantile,
                bottom_quantile=bottom_quantile,
                use_volatility_scaling=settings_payload["time_series_use_volatility_scaling"],
                use_residual=settings_payload["time_series_use_residual"],
                use_multi_horizon=settings_payload["time_series_use_multi_horizon"],
                use_zscore=settings_payload["time_series_use_zscore"],
                winsorize_sigma=settings_payload["time_series_winsorize_sigma"],
            )
            result = spec.compute(price_history, fundamentals_by_ticker, factor_settings)
        result.skipped.update(fetch_skipped)
        result.skipped.update(fundamentals_skipped)

        report_title = f"Time-Series {strategy} Report"
        return format_time_series_report(
            title=report_title,
            as_of=as_of,
            universe=universe,
            settings=settings_payload,
            result=result,
        )

    def _collect_fundamentals(
        self, tickers: list[str], as_of: str
    ) -> tuple[dict[str, dict], dict[str, str]]:
        fundamentals_by_ticker: dict[str, dict] = {}
        skipped: dict[str, str] = {}
        if not tickers:
            return fundamentals_by_ticker, skipped
        max_workers = min(8, max(1, len(tickers)))
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(self._load_fundamentals, ticker, as_of) for ticker in tickers]
            for future in as_completed(futures):
                ticker, payload, reason = future.result()
                if payload and _has_fundamentals_data(payload):
                    fundamentals_by_ticker[ticker] = payload
                if reason:
                    skipped[ticker] = reason
                elif not payload or not _has_fundamentals_data(payload):
                    skipped[ticker] = "missing_fundamentals"
        return fundamentals_by_ticker, skipped

    def _load_fundamentals(
        self, ticker: str, as_of: str
    ) -> tuple[str, dict | None, str | None]:
        cache_payload = load_cached_market_data(ticker) or {}
        fundamentals_cache = cache_payload.get("fundamentals") or {}
        cached_as_of = fundamentals_cache.get("as_of")
        cached_payload = fundamentals_cache.get("payload")
        if cached_as_of == as_of and isinstance(cached_payload, dict):
            return ticker, cached_payload, None

        if self.api_client is None:
            self.api_client = MassiveApiClient(self.controller.api_key)

        errors: list[str] = []
        ticker_details: dict = {}
        financials: list[dict] = []
        dividends: list[dict] = []
        earnings: list[dict] = []

        try:
            self._throttle_request()
            ticker_details = self.api_client.fetch_ticker_details(ticker)
        except HTTPError as exc:
            errors.append(f"ticker_details_http_error_{exc.code}")
        except (TimeoutError, socket.timeout):
            errors.append("ticker_details_timeout")
        except URLError as exc:
            errors.append(f"ticker_details_url_error_{exc.reason}")

        try:
            self._throttle_request()
            financials = self.api_client.fetch_financials(ticker, period="annual")
        except HTTPError as exc:
            errors.append(f"financials_http_error_{exc.code}")
        except (TimeoutError, socket.timeout):
            errors.append("financials_timeout")
        except URLError as exc:
            errors.append(f"financials_url_error_{exc.reason}")

        try:
            self._throttle_request()
            dividends = self.api_client.fetch_dividends(ticker)
        except HTTPError as exc:
            errors.append(f"dividends_http_error_{exc.code}")
        except (TimeoutError, socket.timeout):
            errors.append("dividends_timeout")
        except URLError as exc:
            errors.append(f"dividends_url_error_{exc.reason}")

        try:
            self._throttle_request()
            earnings = self.api_client.fetch_earnings(ticker)
        except HTTPError as exc:
            if exc.code not in {403, 404}:
                errors.append(f"earnings_http_error_{exc.code}")
        except (TimeoutError, socket.timeout):
            errors.append("earnings_timeout")
        except URLError as exc:
            errors.append(f"earnings_url_error_{exc.reason}")

        payload = self._build_fundamentals_payload(
            ticker_details, financials, dividends, earnings
        )
        cache_payload["fundamentals"] = {
            "as_of": as_of,
            "payload": payload,
        }
        save_cached_market_data(ticker, cache_payload)
        reason = errors[0] if errors else None
        if reason is None and not _has_fundamentals_data(payload):
            reason = "missing_fundamentals"
        return ticker, payload, reason

    def _build_fundamentals_payload(
        self,
        ticker_details: dict,
        financials: list[dict],
        dividends: list[dict],
        earnings: list[dict],
    ) -> dict:
        financials_record = self._select_latest_by_date(
            financials,
            ["fiscal_period_end_date", "reporting_date", "filing_date", "calendar_date"],
        )
        financials_record = financials_record or {}

        book_value = _get_nested_value(
            financials_record,
            [
                ("financials", "balance_sheet", "equity"),
                ("financials", "balance_sheet", "stockholders_equity"),
                ("financials", "balance_sheet", "total_equity"),
                ("financials", "balance_sheet", "shareholders_equity"),
            ],
        )
        total_assets = _get_nested_value(
            financials_record,
            [
                ("financials", "balance_sheet", "assets"),
                ("financials", "balance_sheet", "total_assets"),
            ],
        )
        total_equity = _get_nested_value(
            financials_record,
            [
                ("financials", "balance_sheet", "equity"),
                ("financials", "balance_sheet", "total_equity"),
                ("financials", "balance_sheet", "stockholders_equity"),
            ],
        )
        net_income = _get_nested_value(
            financials_record,
            [
                ("financials", "income_statement", "net_income_loss"),
                ("financials", "income_statement", "net_income"),
            ],
        )
        eps = _get_nested_value(
            financials_record,
            [
                ("financials", "income_statement", "basic_earnings_per_share"),
                ("financials", "income_statement", "diluted_earnings_per_share"),
                ("financials", "income_statement", "eps"),
            ],
        )
        shares_outstanding = (
            _coerce_number(ticker_details.get("weighted_shares_outstanding"))
            or _coerce_number(ticker_details.get("share_class_shares_outstanding"))
            or _get_nested_value(
                financials_record,
                [
                    ("financials", "income_statement", "weighted_average_shares_outstanding"),
                    ("financials", "income_statement", "weighted_average_shares_outstanding_diluted"),
                ],
            )
        )
        market_cap = _coerce_number(ticker_details.get("market_cap"))

        dividends_ttm, last_dividend_date, last_dividend_amount = self._summarize_dividends(
            dividends
        )
        earnings_record = self._select_latest_by_date(
            earnings,
            ["reporting_date", "fiscal_period_end_date", "publish_date", "date"],
        )
        earnings_record = earnings_record or {}
        earnings_actual = _coerce_number(
            earnings_record.get("actual_eps")
            or earnings_record.get("eps_actual")
            or earnings_record.get("eps")
        )
        earnings_estimate = _coerce_number(
            earnings_record.get("estimated_eps") or earnings_record.get("eps_estimate")
        )
        if earnings_actual is None:
            earnings_actual = eps
        earnings_surprise = (
            earnings_actual - earnings_estimate
            if earnings_actual is not None and earnings_estimate is not None
            else None
        )

        return {
            "market_cap": market_cap,
            "book_value": book_value or total_equity,
            "total_assets": total_assets,
            "total_equity": total_equity,
            "net_income": net_income,
            "eps": eps,
            "dividends_ttm": dividends_ttm,
            "shares_outstanding": shares_outstanding,
            "earnings_actual": earnings_actual,
            "earnings_estimate": earnings_estimate,
            "earnings_surprise": earnings_surprise,
            "last_ex_dividend_date": last_dividend_date.isoformat()
            if last_dividend_date
            else None,
            "last_dividend_amount": last_dividend_amount,
        }

    def _select_latest_by_date(
        self, records: list[dict], date_fields: list[str]
    ) -> dict | None:
        best_record: dict | None = None
        best_date: date | None = None
        for record in records:
            record_date: date | None = None
            for field in date_fields:
                record_date = _parse_iso_date(record.get(field))
                if record_date:
                    break
            if record_date and (best_date is None or record_date > best_date):
                best_date = record_date
                best_record = record
        if best_record is None and records:
            return records[0]
        return best_record

    def _summarize_dividends(
        self, dividends: list[dict]
    ) -> tuple[float | None, date | None, float | None]:
        as_of_date = effective_market_date()
        ttm_start = as_of_date - timedelta(days=365)
        total = 0.0
        total_found = False
        last_date: date | None = None
        last_amount: float | None = None
        for record in dividends:
            ex_date = _parse_iso_date(
                record.get("ex_dividend_date")
                or record.get("ex_date")
                or record.get("pay_date")
                or record.get("payment_date")
            )
            amount = _coerce_number(
                record.get("cash_amount") or record.get("dividend") or record.get("amount")
            )
            if ex_date and amount is not None:
                if ex_date >= ttm_start:
                    total += amount
                    total_found = True
                if last_date is None or ex_date > last_date:
                    last_date = ex_date
                    last_amount = amount
        dividends_ttm = total if total_found else None
        return dividends_ttm, last_date, last_amount

    def _collect_price_history(
        self, tickers: list[str], lookback_days: int, skip_days: int
    ) -> tuple[dict[str, list[float]], dict[str, str]]:
        min_points = lookback_days + skip_days + 1
        buffer_days = max(10, int(min_points * 1.5))
        days_back = min_points + buffer_days
        as_of = effective_market_date().isoformat()

        prices_by_ticker: dict[str, list[float]] = {}
        skipped: dict[str, str] = {}

        if len(tickers) >= 200:
            if self.api_client is None:
                self.api_client = MassiveApiClient(self.controller.api_key)
            grouped = [ticker for ticker in tickers if self._is_grouped_eligible(ticker)]
            special = [ticker for ticker in tickers if ticker not in grouped]
            prices_by_ticker, skipped = self._collect_grouped_history(grouped, min_points)
            if special:
                special_prices, special_skipped = self._collect_ticker_history(
                    special, days_back, min_points, as_of
                )
                prices_by_ticker.update(special_prices)
                skipped.update(special_skipped)
            fallback_tickers = [
                ticker
                for ticker, reason in skipped.items()
                if reason == "insufficient_history" and ticker not in special
            ]
            if fallback_tickers:
                fallback_prices, fallback_skipped = self._collect_ticker_history(
                    fallback_tickers, days_back, min_points, as_of
                )
                prices_by_ticker.update(fallback_prices)
                for ticker in fallback_tickers:
                    skipped.pop(ticker, None)
                skipped.update(fallback_skipped)
            return prices_by_ticker, skipped

        return self._collect_ticker_history(tickers, days_back, min_points, as_of)

    def _collect_ticker_history(
        self, tickers: list[str], days_back: int, min_points: int, as_of: str
    ) -> tuple[dict[str, list[float]], dict[str, str]]:
        prices_by_ticker: dict[str, list[float]] = {}
        skipped: dict[str, str] = {}
        max_workers = min(12, max(1, len(tickers)))
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [
                executor.submit(
                    self._load_daily_closes, ticker, days_back, min_points, as_of
                )
                for ticker in tickers
            ]
            for future in as_completed(futures):
                ticker, prices, reason = future.result()
                if prices:
                    prices_by_ticker[ticker] = prices
                else:
                    skipped[ticker] = reason or "no_data"
        return prices_by_ticker, skipped

    def _collect_grouped_history(
        self, tickers: list[str], min_points: int
    ) -> tuple[dict[str, list[float]], dict[str, str]]:
        if self.api_client is None:
            self.api_client = MassiveApiClient(self.controller.api_key)
        universe = set(tickers)
        closes: dict[str, list[float]] = {ticker: [] for ticker in tickers}
        skipped: dict[str, str] = {}
        pending = set(tickers)

        end_date = effective_market_date()
        cache_dir = DATA_DIR / "grouped_cache" / end_date.isoformat()
        cache_dir.mkdir(parents=True, exist_ok=True)
        max_calendar_days = max(30, min_points * 4)
        current_date = end_date
        days_checked = 0

        retry_count = 0
        max_retries = 4
        backoff_seconds = 1.0

        while pending and days_checked < max_calendar_days:
            if current_date.weekday() >= 5:
                current_date -= timedelta(days=1)
                days_checked += 1
                continue

            try:
                day_cache_path = cache_dir / f"{current_date}.json"
                if day_cache_path.exists():
                    day_payload = json.loads(day_cache_path.read_text())
                    if day_payload:
                        sample_value = next(iter(day_payload.values()))
                        if isinstance(sample_value, (int, float)):
                            day_payload = {
                                ticker: {"close": float(value), "volume": None}
                                for ticker, value in day_payload.items()
                            }
                else:
                    self._throttle_request()
                    aggregates = self.api_client.fetch_grouped_daily_aggregates(current_date)
                    day_payload = {}
                    for item in aggregates:
                        ticker = item.get("T")
                        if ticker not in universe:
                            continue
                        close_value = item.get("c")
                        volume_value = item.get("v")
                        if isinstance(close_value, (int, float)) and close_value > 0:
                            day_payload[ticker] = {
                                "close": float(close_value),
                                "volume": float(volume_value)
                                if isinstance(volume_value, (int, float))
                                else None,
                            }
                    day_cache_path.write_text(json.dumps(day_payload))
            except HTTPError as exc:
                if exc.code == 429 and retry_count < max_retries:
                    retry_after = exc.headers.get("Retry-After") if exc.headers else None
                    wait_seconds = backoff_seconds
                    if retry_after:
                        try:
                            wait_seconds = max(wait_seconds, float(retry_after))
                        except ValueError:
                            pass
                    time.sleep(wait_seconds)
                    retry_count += 1
                    backoff_seconds *= 2
                    continue
                for ticker in pending:
                    skipped[ticker] = f"http_error_{exc.code}"
                break
            except (TimeoutError, socket.timeout):
                if retry_count < max_retries:
                    time.sleep(backoff_seconds)
                    retry_count += 1
                    backoff_seconds *= 2
                    continue
                for ticker in pending:
                    skipped[ticker] = "timeout"
                break
            except URLError as exc:
                for ticker in pending:
                    skipped[ticker] = f"url_error_{exc.reason}"
                break

            retry_count = 0
            backoff_seconds = 3.0
            for ticker, payload in day_payload.items():
                if ticker in pending:
                    closes[ticker].append(payload)
                    if len(closes[ticker]) >= min_points:
                        pending.discard(ticker)

            current_date -= timedelta(days=1)
            days_checked += 1

        prices_by_ticker: dict[str, list[float]] = {}
        for ticker, values in closes.items():
            if len(values) >= min_points:
                prices_by_ticker[ticker] = list(reversed(values))
            elif ticker not in skipped:
                skipped[ticker] = "insufficient_history"

        return prices_by_ticker, skipped

    def _is_grouped_eligible(self, ticker: str) -> bool:
        return bool(self._grouped_ticker_pattern.fullmatch(ticker))

    def _load_daily_closes(
        self, ticker: str, days_back: int, min_points: int, as_of: str
    ) -> tuple[str, list[float] | None, str | None]:
        cache_payload = load_cached_market_data(ticker) or {}
        daily_cache = cache_payload.get("daily_closes") or {}
        cached_as_of = daily_cache.get("as_of")
        cached_prices = daily_cache.get("prices")
        if (
            cached_as_of == as_of
            and isinstance(cached_prices, list)
            and len(cached_prices) >= min_points
        ):
            if cached_prices and isinstance(cached_prices[0], dict):
                return ticker, cached_prices, None
            return ticker, [{"close": float(value), "volume": None} for value in cached_prices], None

        if self.api_client is None:
            self.api_client = MassiveApiClient(self.controller.api_key)

        current_days_back = days_back
        max_days_back = max(days_back * 3, min_points * 4)
        closes: list[dict] = []

        retry_count = 0
        max_retries = 4
        backoff_seconds = 1.0
        while current_days_back <= max_days_back:
            self._throttle_request()
            try:
                aggregates = self.api_client.fetch_daily_aggregates(
                    ticker, days_back=current_days_back
                )
            except HTTPError as exc:
                if exc.code == 429 and retry_count < max_retries:
                    retry_after = exc.headers.get("Retry-After") if exc.headers else None
                    wait_seconds = backoff_seconds
                    if retry_after:
                        try:
                            wait_seconds = max(wait_seconds, float(retry_after))
                        except ValueError:
                            pass
                    time.sleep(wait_seconds)
                    retry_count += 1
                    backoff_seconds *= 2
                    continue
                return ticker, None, f"http_error_{exc.code}"
            except (TimeoutError, socket.timeout):
                if retry_count < max_retries:
                    time.sleep(backoff_seconds)
                    retry_count += 1
                    backoff_seconds *= 2
                    continue
                return ticker, None, "timeout"
            except URLError as exc:
                return ticker, None, f"url_error_{exc.reason}"

            closes = []
            for item in sorted(aggregates, key=lambda row: row.get("t", 0)):
                close_value = item.get("c")
                volume_value = item.get("v")
                if isinstance(close_value, (int, float)) and close_value > 0:
                    closes.append(
                        {
                            "close": float(close_value),
                            "volume": float(volume_value)
                            if isinstance(volume_value, (int, float))
                            else None,
                        }
                    )

            if len(closes) >= min_points:
                break
            current_days_back = int(current_days_back * 1.5) + 1

        if closes:
            cache_payload["daily_closes"] = {
                "as_of": as_of,
                "prices": closes,
                "days_back": current_days_back,
            }
            save_cached_market_data(ticker, cache_payload)

        if len(closes) < min_points:
            return ticker, None, "insufficient_history"
        return ticker, closes, None

    def _throttle_request(self) -> None:
        with self._rate_lock:
            now = time.monotonic()
            elapsed = now - self._last_request_time
            if elapsed < self._min_request_interval:
                time.sleep(self._min_request_interval - elapsed)
            self._last_request_time = time.monotonic()

    def _write_report(
        self, report: str, output_dir: str, strategy: str, analysis_slug: str
    ) -> Path:
        directory = Path(output_dir)
        directory.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        strategy_slug = re.sub(r"[^a-z0-9]+", "_", strategy.strip().lower()).strip("_")
        output_path = directory / f"{analysis_slug}_{strategy_slug}_{timestamp}.txt"
        output_path.write_text(report)
        return output_path


class CallPutAnalysisPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller
        self.api_client: MassiveApiClient | None = None
        self.option_contract: dict | None = None
        self.all_option_records: list[dict] = []
        self.option_records: list[dict] = []

        header = ttk.Label(self, text="Call/Put Analysis", font=("Arial", 18, "bold"))
        header.pack(pady=10)

        self.strategy_label = ttk.Label(self, text="Strategy: --", font=("Arial", 12))
        self.strategy_label.pack(pady=(0, 10))

        self.options_frame = ttk.LabelFrame(self, text="Option Candidates")
        self.options_frame.pack(padx=30, pady=10, fill="both", expand=True)
        self.options_frame.columnconfigure(0, weight=1)
        self.options_frame.columnconfigure(1, weight=0)
        self.options_frame.rowconfigure(0, weight=1)

        list_frame = ttk.Frame(self.options_frame)
        list_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=8)
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        filter_frame = ttk.Frame(self.options_frame)
        filter_frame.grid(row=0, column=1, sticky="nsew", padx=(0, 10), pady=8)
        filter_frame.columnconfigure(1, weight=1)

        ttk.Label(filter_frame, text="Max Loss / Contract Price").grid(
            row=0, column=0, padx=5, pady=2, sticky="w"
        )
        self.max_loss_var = tk.StringVar()
        self.max_loss_entry = ttk.Entry(filter_frame, textvariable=self.max_loss_var, width=12)
        self.max_loss_entry.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        self.max_loss_entry.bind("<KeyRelease>", self.on_filter_change)

        ttk.Label(filter_frame, text="Min Likelihood (%)").grid(
            row=1, column=0, padx=5, pady=2, sticky="w"
        )
        self.likelihood_var = tk.StringVar()
        self.likelihood_entry = ttk.Entry(
            filter_frame, textvariable=self.likelihood_var, width=12
        )
        self.likelihood_entry.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        self.likelihood_entry.bind("<KeyRelease>", self.on_filter_change)

        ttk.Label(filter_frame, text="Expiration").grid(
            row=2, column=0, padx=5, pady=2, sticky="w"
        )
        self.expiration_var = tk.StringVar(value="All")
        self.expiration_dropdown = ttk.Combobox(
            filter_frame, textvariable=self.expiration_var, state="readonly", width=18
        )
        self.expiration_dropdown.grid(row=2, column=1, padx=5, pady=2, sticky="ew")
        self.expiration_dropdown.bind("<<ComboboxSelected>>", self.on_filter_change)

        ttk.Label(filter_frame, text="Strike").grid(
            row=3, column=0, padx=5, pady=2, sticky="w"
        )
        self.strike_var = tk.StringVar(value="All")
        self.strike_dropdown = ttk.Combobox(
            filter_frame, textvariable=self.strike_var, state="readonly", width=12
        )
        self.strike_dropdown.grid(row=3, column=1, padx=5, pady=2, sticky="ew")
        self.strike_dropdown.bind("<<ComboboxSelected>>", self.on_filter_change)

        ttk.Label(filter_frame, text="Type").grid(row=4, column=0, padx=5, pady=2, sticky="w")
        self.type_var = tk.StringVar(value="All")
        self.type_dropdown = ttk.Combobox(
            filter_frame, textvariable=self.type_var, state="readonly", width=10
        )
        self.type_dropdown.grid(row=4, column=1, padx=5, pady=2, sticky="ew")
        self.type_dropdown.bind("<<ComboboxSelected>>", self.on_filter_change)

        self.options_list = tk.Listbox(list_frame, height=12, width=48)
        options_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.options_list.yview)
        self.options_list.configure(yscrollcommand=options_scroll.set)
        self.options_list.grid(row=0, column=0, sticky="nsew")
        options_scroll.grid(row=0, column=1, sticky="ns")
        self.options_list.bind("<<ListboxSelect>>", self.on_option_select)

        self.option_info_frame = ttk.LabelFrame(self, text="Selected Option")
        self.option_info_frame.pack(padx=30, pady=(5, 10), fill="x")
        self.option_values: dict[str, ttk.Label] = {}
        self._build_info_grid(
            self.option_info_frame,
            [
                ("Contract", "contract"),
                ("Expiration", "expiration"),
                ("Type", "type"),
                ("Strike", "strike"),
                ("Contract Price", "premium"),
                ("Likelihood", "likelihood"),
            ],
            self.option_values,
            columns=3,
        )

        self.greeks_frame = ttk.LabelFrame(self, text="Option Greeks")
        self.greeks_frame.pack(padx=30, pady=(5, 10), fill="x")
        self.greeks_values: dict[str, ttk.Label] = {}
        self._build_info_grid(
            self.greeks_frame,
            [
                ("Delta", "delta"),
                ("Gamma", "gamma"),
                ("Theta", "theta"),
                ("Vega", "vega"),
                ("Rho", "rho"),
                ("IV", "iv"),
            ],
            self.greeks_values,
            columns=3,
        )

        button_row = ttk.Frame(self)
        button_row.pack(pady=10)
        ttk.Button(button_row, text="Refresh", command=self.load_market_data).grid(
            row=0, column=0, padx=10
        )
        ttk.Button(
            button_row,
            text="Select Stock",
            command=lambda: controller.show_frame("TickerSelectPage"),
        ).grid(row=0, column=1, padx=10)
        ttk.Button(
            button_row,
            text="Back to Analysis",
            command=lambda: controller.show_frame("AnalysisPage"),
        ).grid(row=0, column=2, padx=10)
        ttk.Button(
            button_row,
            text="Back to Main Menu",
            command=lambda: controller.show_frame("MainMenu"),
        ).grid(row=0, column=3, padx=10)

    def _build_info_grid(
        self,
        parent: ttk.Frame,
        rows: list[tuple[str, str]],
        target: dict[str, ttk.Label],
        columns: int = 1,
    ) -> None:
        for item_index, (label, key) in enumerate(rows):
            row_index = item_index // columns
            column_index = (item_index % columns) * 2
            ttk.Label(parent, text=label).grid(
                row=row_index, column=column_index, padx=10, pady=4, sticky="w"
            )
            value_label = ttk.Label(parent, text="--", foreground="#b00020")
            value_label.grid(
                row=row_index, column=column_index + 1, padx=10, pady=4, sticky="w"
            )
            target[key] = value_label
        for index in range(columns * 2):
            parent.columnconfigure(index, weight=1)

    def _format_float(self, value: float) -> str:
        decimals = 2 if abs(value) >= 1 else 4
        multiplier = 10**decimals
        truncated = math.trunc(value * multiplier) / multiplier
        return f"{truncated:.{decimals}f}".rstrip("0").rstrip(".")

    def _set_value(self, label: ttk.Label, value: str | int | float | None) -> None:
        if value in (None, "", "--"):
            label.config(text="--", foreground="#b00020")
        else:
            if isinstance(value, float):
                text = self._format_float(value)
            elif isinstance(value, int):
                text = str(value)
            else:
                text = str(value)
            label.config(text=text, foreground="#0a7a2f")

    def refresh(self) -> None:
        self.api_client = MassiveApiClient(load_api_key()) if load_api_key() else None
        self.load_market_data()

    def load_market_data(self) -> None:
        if not self.api_client:
            messagebox.showinfo(
                "Missing key", "Enter or set a Massive API key to load options data."
            )
            return
        ticker = self.controller.state.selected_ticker
        if not ticker:
            messagebox.showinfo("Missing ticker", "Select a ticker first.")
            return
        strategy = self.controller.state.option_strategy
        self.strategy_label.config(text=f"Strategy: {strategy}")
        try:
            option_records = load_option_records(self.api_client, ticker)
        except HTTPError as exc:
            self._show_api_error(exc, "Massive", "Verify your Massive API key.")
            return
        except URLError as exc:
            self._show_error_dialog(
                "Connection Error",
                f"Could not reach Massive API endpoint: {exc.reason}",
            )
            return
        self.all_option_records = [
            {
                **record,
                "premium": option_mid_price(record),
                "likelihood": option_likelihood(record),
            }
            for record in option_records
        ]
        if strategy == "Naked Call":
            self.type_var.set("CALL")
        elif strategy == "Naked Put":
            self.type_var.set("PUT")
        else:
            self.type_var.set("All")
        self._refresh_option_filters(reset=True)

    def _get_filter_value(self, var: tk.StringVar) -> str | None:
        value = var.get()
        return None if value == "All" else value

    def _record_matches_filters(self, record: dict, filters: dict[str, str | None]) -> bool:
        expiration = record.get("expiration_date")
        strike = format_strike(record.get("strike_price"))
        contract_type = normalize_contract_type(record.get("contract_type"))
        if filters.get("expiration") and filters["expiration"] != expiration:
            return False
        if filters.get("strike") and filters["strike"] != strike:
            return False
        if filters.get("type") and filters["type"] != contract_type:
            return False
        return True

    def _record_matches_constraints(self, record: dict) -> bool:
        max_loss = parse_float(self.max_loss_var.get())
        min_likelihood = normalize_likelihood_threshold(self.likelihood_var.get())
        premium = record.get("premium")
        likelihood = record.get("likelihood")
        if max_loss is not None:
            if not isinstance(premium, (int, float)) or premium > max_loss:
                return False
        if min_likelihood is not None:
            if not isinstance(likelihood, (int, float)) or likelihood < min_likelihood:
                return False
        return True

    def _compute_filter_options(
        self, records: list[dict], current: dict[str, str | None]
    ) -> dict[str, list[str]]:
        options: dict[str, set[str]] = {"expiration": set(), "strike": set(), "type": set()}
        for record in records:
            expiration = record.get("expiration_date")
            strike = format_strike(record.get("strike_price"))
            contract_type = normalize_contract_type(record.get("contract_type"))
            if self._record_matches_filters(record, {**current, "expiration": None}):
                if expiration:
                    options["expiration"].add(expiration)
            if self._record_matches_filters(record, {**current, "strike": None}):
                if strike:
                    options["strike"].add(strike)
            if self._record_matches_filters(record, {**current, "type": None}):
                if contract_type:
                    options["type"].add(contract_type)
        return {
            "expiration": sorted(options["expiration"]),
            "strike": sorted(
                options["strike"],
                key=lambda value: float(value) if value.replace(".", "", 1).isdigit() else value,
            ),
            "type": sorted(options["type"]),
        }

    def _refresh_option_filters(self, reset: bool = False) -> None:
        if reset:
            self.expiration_var.set("All")
            self.strike_var.set("All")
            if self.type_var.get() not in ("CALL", "PUT"):
                self.type_var.set("All")
        eligible_records = [
            record for record in self.all_option_records if self._record_matches_constraints(record)
        ]
        filters = {
            "expiration": self._get_filter_value(self.expiration_var),
            "strike": self._get_filter_value(self.strike_var),
            "type": self._get_filter_value(self.type_var),
        }
        options = self._compute_filter_options(eligible_records, filters)
        for key, dropdown, var in (
            ("expiration", self.expiration_dropdown, self.expiration_var),
            ("strike", self.strike_dropdown, self.strike_var),
            ("type", self.type_dropdown, self.type_var),
        ):
            values = ["All"] + options[key]
            dropdown["values"] = values
            if var.get() not in values:
                var.set("All")
        self._apply_option_filters(eligible_records)

    def _apply_option_filters(self, eligible_records: list[dict]) -> None:
        filters = {
            "expiration": self._get_filter_value(self.expiration_var),
            "strike": self._get_filter_value(self.strike_var),
            "type": self._get_filter_value(self.type_var),
        }
        self.option_records = [
            record
            for record in eligible_records
            if self._record_matches_filters(record, filters)
        ]
        self.options_list.delete(0, tk.END)
        if not self.option_records:
            self.options_list.insert(tk.END, "No option contracts returned.")
            self.option_contract = None
        else:
            for contract in self.option_records:
                likelihood = contract.get("likelihood")
                likelihood_label = (
                    f"{likelihood * 100:.0f}%" if isinstance(likelihood, (int, float)) else "--"
                )
                premium = contract.get("premium")
                premium_label = f"{premium:.2f}" if isinstance(premium, (int, float)) else "--"
                self.options_list.insert(
                    tk.END,
                    "{ticker} {expiration} {type} {strike} | Loss {loss} | Likely {likelihood}".format(
                        ticker=contract.get("ticker", "--"),
                        expiration=contract.get("expiration_date", "--"),
                        type=str(contract.get("contract_type", "--")).upper(),
                        strike=contract.get("strike_price", "--"),
                        loss=premium_label,
                        likelihood=likelihood_label,
                    ),
                )
            self.options_list.selection_set(0)
            self.options_list.see(0)
            self.option_contract = self.option_records[0]
        self._sync_option_snapshot()
        self._sync_greeks()

    def _sync_option_snapshot(self) -> None:
        contract = self.option_contract or {}
        self._set_value(self.option_values["contract"], contract.get("ticker"))
        self._set_value(self.option_values["expiration"], contract.get("expiration_date"))
        contract_type = normalize_contract_type(contract.get("contract_type"))
        self._set_value(self.option_values["type"], contract_type)
        self._set_value(self.option_values["strike"], contract.get("strike_price"))
        premium = contract.get("premium")
        self._set_value(self.option_values["premium"], premium)
        likelihood = contract.get("likelihood")
        likelihood_label = (
            f"{likelihood * 100:.1f}%" if isinstance(likelihood, (int, float)) else None
        )
        self._set_value(self.option_values["likelihood"], likelihood_label)

    def _sync_greeks(self) -> None:
        greeks = extract_greeks(self.option_contract or {})
        self._set_value(self.greeks_values["delta"], greeks.get("delta"))
        self._set_value(self.greeks_values["gamma"], greeks.get("gamma"))
        self._set_value(self.greeks_values["theta"], greeks.get("theta"))
        self._set_value(self.greeks_values["vega"], greeks.get("vega"))
        self._set_value(self.greeks_values["rho"], greeks.get("rho"))
        self._set_value(self.greeks_values["iv"], greeks.get("iv"))

    def on_option_select(self, _event: object) -> None:
        selection = self.options_list.curselection()
        if not selection:
            return
        index = selection[0]
        if index >= len(self.option_records):
            return
        self.option_contract = self.option_records[index]
        self._sync_option_snapshot()
        self._sync_greeks()

    def on_filter_change(self, _event: object) -> None:
        self._refresh_option_filters()

    def _show_api_error(self, exc: HTTPError, service: str, hint: str | None = None) -> None:
        detail = format_http_error_detail(exc)
        detail_msg = f"\nDetails: {detail}" if detail else ""
        hint_msg = f"\n{hint}" if hint else ""
        self._show_error_dialog(
            "API Error",
            f"{service} API returned an error: {exc.code} {exc.reason}.{detail_msg}{hint_msg}",
        )

    def _show_error_dialog(self, title: str, message: str) -> None:
        dialog = tk.Toplevel(self)
        dialog.title(title)
        dialog.geometry("620x320")
        dialog.transient(self)

        dialog.rowconfigure(0, weight=1)
        dialog.columnconfigure(0, weight=1)

        text_frame = ttk.Frame(dialog)
        text_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

        text_widget = tk.Text(text_frame, wrap="word")
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        text_widget.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        text_widget.insert("1.0", message)
        text_widget.focus_set()

        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=1, column=0, pady=(0, 10))

        def copy_to_clipboard() -> None:
            dialog.clipboard_clear()
            dialog.clipboard_append(text_widget.get("1.0", "end-1c"))
            dialog.update_idletasks()

        ttk.Button(button_frame, text="Copy", command=copy_to_clipboard).grid(
            row=0, column=0, padx=5
        )
        ttk.Button(button_frame, text="Close", command=dialog.destroy).grid(
            row=0, column=1, padx=5
        )


class SpreadAnalysisPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller
        self.api_client: MassiveApiClient | None = None
        self.spread_records: list[dict] = []
        self.all_spread_records: list[dict] = []
        self.selected_spread: dict | None = None

        header = ttk.Label(self, text="Spread Analysis", font=("Arial", 18, "bold"))
        header.pack(pady=10)

        self.strategy_label = ttk.Label(self, text="Strategy: --", font=("Arial", 12))
        self.strategy_label.pack(pady=(0, 10))

        self.options_frame = ttk.LabelFrame(self, text="Spread Candidates")
        self.options_frame.pack(padx=30, pady=10, fill="both", expand=True)
        self.options_frame.columnconfigure(0, weight=1)
        self.options_frame.columnconfigure(1, weight=0)
        self.options_frame.rowconfigure(0, weight=1)

        list_frame = ttk.Frame(self.options_frame)
        list_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=8)
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        filter_frame = ttk.Frame(self.options_frame)
        filter_frame.grid(row=0, column=1, sticky="nsew", padx=(0, 10), pady=8)
        filter_frame.columnconfigure(1, weight=1)

        ttk.Label(filter_frame, text="Max Loss / Spread Price").grid(
            row=0, column=0, padx=5, pady=2, sticky="w"
        )
        self.max_loss_var = tk.StringVar()
        self.max_loss_entry = ttk.Entry(filter_frame, textvariable=self.max_loss_var, width=12)
        self.max_loss_entry.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        self.max_loss_entry.bind("<KeyRelease>", self.on_filter_change)

        ttk.Label(filter_frame, text="Min Likelihood (%)").grid(
            row=1, column=0, padx=5, pady=2, sticky="w"
        )
        self.likelihood_var = tk.StringVar()
        self.likelihood_entry = ttk.Entry(
            filter_frame, textvariable=self.likelihood_var, width=12
        )
        self.likelihood_entry.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        self.likelihood_entry.bind("<KeyRelease>", self.on_filter_change)

        ttk.Label(filter_frame, text="Expiration").grid(
            row=2, column=0, padx=5, pady=2, sticky="w"
        )
        self.expiration_var = tk.StringVar(value="All")
        self.expiration_dropdown = ttk.Combobox(
            filter_frame, textvariable=self.expiration_var, state="readonly", width=18
        )
        self.expiration_dropdown.grid(row=2, column=1, padx=5, pady=2, sticky="ew")
        self.expiration_dropdown.bind("<<ComboboxSelected>>", self.on_filter_change)

        ttk.Label(filter_frame, text="Strike").grid(
            row=3, column=0, padx=5, pady=2, sticky="w"
        )
        self.strike_var = tk.StringVar(value="All")
        self.strike_dropdown = ttk.Combobox(
            filter_frame, textvariable=self.strike_var, state="readonly", width=12
        )
        self.strike_dropdown.grid(row=3, column=1, padx=5, pady=2, sticky="ew")
        self.strike_dropdown.bind("<<ComboboxSelected>>", self.on_filter_change)

        ttk.Label(filter_frame, text="Type").grid(row=4, column=0, padx=5, pady=2, sticky="w")
        self.type_var = tk.StringVar(value="All")
        self.type_dropdown = ttk.Combobox(
            filter_frame, textvariable=self.type_var, state="readonly", width=10
        )
        self.type_dropdown.grid(row=4, column=1, padx=5, pady=2, sticky="ew")
        self.type_dropdown.bind("<<ComboboxSelected>>", self.on_filter_change)

        self.options_list = tk.Listbox(list_frame, height=12, width=48)
        options_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.options_list.yview)
        self.options_list.configure(yscrollcommand=options_scroll.set)
        self.options_list.grid(row=0, column=0, sticky="nsew")
        options_scroll.grid(row=0, column=1, sticky="ns")
        self.options_list.bind("<<ListboxSelect>>", self.on_spread_select)

        self.spread_info_frame = ttk.LabelFrame(self, text="Selected Spread")
        self.spread_info_frame.pack(padx=30, pady=(5, 10), fill="x")
        self.spread_values: dict[str, ttk.Label] = {}
        self._build_info_grid(
            self.spread_info_frame,
            [
                ("Strategy", "strategy"),
                ("Expiration(s)", "expiration"),
                ("Type", "type"),
                ("Strikes", "strikes"),
                ("Spread Price", "premium"),
                ("Likelihood", "likelihood"),
            ],
            self.spread_values,
            columns=3,
        )

        self.legs_frame = ttk.LabelFrame(self, text="Leg Details")
        self.legs_frame.pack(padx=30, pady=(5, 10), fill="x")
        self.leg_values: dict[str, ttk.Label] = {}
        self._build_info_grid(
            self.legs_frame,
            [
                ("Long Leg", "long_leg"),
                ("Short Leg", "short_leg"),
            ],
            self.leg_values,
            columns=1,
        )

        self.greeks_frame = ttk.LabelFrame(self, text="Spread Greeks")
        self.greeks_frame.pack(padx=30, pady=(5, 10), fill="x")
        self.greeks_values: dict[str, ttk.Label] = {}
        self._build_info_grid(
            self.greeks_frame,
            [
                ("Delta", "delta"),
                ("Gamma", "gamma"),
                ("Theta", "theta"),
                ("Vega", "vega"),
                ("Rho", "rho"),
                ("IV", "iv"),
            ],
            self.greeks_values,
            columns=3,
        )

        button_row = ttk.Frame(self)
        button_row.pack(pady=10)
        ttk.Button(button_row, text="Refresh", command=self.load_market_data).grid(
            row=0, column=0, padx=10
        )
        ttk.Button(
            button_row,
            text="Select Stock",
            command=lambda: controller.show_frame("TickerSelectPage"),
        ).grid(row=0, column=1, padx=10)
        ttk.Button(
            button_row,
            text="Back to Analysis",
            command=lambda: controller.show_frame("AnalysisPage"),
        ).grid(row=0, column=2, padx=10)
        ttk.Button(
            button_row,
            text="Back to Main Menu",
            command=lambda: controller.show_frame("MainMenu"),
        ).grid(row=0, column=3, padx=10)

    def _build_info_grid(
        self,
        parent: ttk.Frame,
        rows: list[tuple[str, str]],
        target: dict[str, ttk.Label],
        columns: int = 1,
    ) -> None:
        for item_index, (label, key) in enumerate(rows):
            row_index = item_index // columns
            column_index = (item_index % columns) * 2
            ttk.Label(parent, text=label).grid(
                row=row_index, column=column_index, padx=10, pady=4, sticky="w"
            )
            value_label = ttk.Label(parent, text="--", foreground="#b00020")
            value_label.grid(
                row=row_index, column=column_index + 1, padx=10, pady=4, sticky="w"
            )
            target[key] = value_label
        for index in range(columns * 2):
            parent.columnconfigure(index, weight=1)

    def _format_float(self, value: float) -> str:
        decimals = 2 if abs(value) >= 1 else 4
        multiplier = 10**decimals
        truncated = math.trunc(value * multiplier) / multiplier
        return f"{truncated:.{decimals}f}".rstrip("0").rstrip(".")

    def _set_value(self, label: ttk.Label, value: str | int | float | None) -> None:
        if value in (None, "", "--"):
            label.config(text="--", foreground="#b00020")
        else:
            if isinstance(value, float):
                text = self._format_float(value)
            elif isinstance(value, int):
                text = str(value)
            else:
                text = str(value)
            label.config(text=text, foreground="#0a7a2f")

    def refresh(self) -> None:
        self.api_client = MassiveApiClient(load_api_key()) if load_api_key() else None
        self.load_market_data()

    def load_market_data(self) -> None:
        if not self.api_client:
            messagebox.showinfo(
                "Missing key", "Enter or set a Massive API key to load options data."
            )
            return
        ticker = self.controller.state.selected_ticker
        if not ticker:
            messagebox.showinfo("Missing ticker", "Select a ticker first.")
            return
        strategy = self.controller.state.option_strategy
        self.strategy_label.config(text=f"Strategy: {strategy}")
        try:
            option_records = load_option_records(self.api_client, ticker)
        except HTTPError as exc:
            self._show_api_error(exc, "Massive", "Verify your Massive API key.")
            return
        except URLError as exc:
            self._show_error_dialog(
                "Connection Error",
                f"Could not reach Massive API endpoint: {exc.reason}",
            )
            return
        self.all_spread_records = self._build_spread_records(option_records, strategy)
        self._refresh_spread_filters(reset=True)

    def _build_spread_records(self, option_records: list[dict], strategy: str) -> list[dict]:
        spreads: list[dict] = []
        if strategy == "Calendar Spread":
            grouped: dict[tuple[str | None, str | None], list[dict]] = {}
            for record in option_records:
                key = (normalize_contract_type(record.get("contract_type")), record.get("strike_price"))
                grouped.setdefault(key, []).append(record)
            for (contract_type, strike), records in grouped.items():
                records_sorted = sorted(
                    records,
                    key=lambda r: r.get("expiration_date") or "",
                )
                for short_leg, long_leg in zip(records_sorted, records_sorted[1:]):
                    premium_long = option_mid_price(long_leg)
                    premium_short = option_mid_price(short_leg)
                    net_premium = None
                    if isinstance(premium_long, (int, float)) or isinstance(
                        premium_short, (int, float)
                    ):
                        net_premium = (premium_long or 0) - (premium_short or 0)
                        net_premium = abs(net_premium)
                    likelihood_values = [
                        option_likelihood(long_leg),
                        option_likelihood(short_leg),
                    ]
                    likelihoods = [
                        value for value in likelihood_values if isinstance(value, (int, float))
                    ]
                    likelihood = sum(likelihoods) / len(likelihoods) if likelihoods else None
                    expiration_label = "{short}→{long}".format(
                        short=short_leg.get("expiration_date", "--"),
                        long=long_leg.get("expiration_date", "--"),
                    )
                    strike_label = format_strike(strike)
                    spreads.append(
                        {
                            "strategy": strategy,
                            "ticker": short_leg.get("ticker") or long_leg.get("ticker"),
                            "expiration_label": expiration_label,
                            "contract_type": contract_type,
                            "strike_pair": strike_label,
                            "long_leg": long_leg,
                            "short_leg": short_leg,
                            "net_premium": net_premium,
                            "likelihood": likelihood,
                            "greeks": combine_greeks(long_leg, short_leg),
                        }
                    )
        else:
            grouped: dict[tuple[str | None, str | None], list[dict]] = {}
            for record in option_records:
                key = (
                    normalize_contract_type(record.get("contract_type")),
                    record.get("expiration_date"),
                )
                grouped.setdefault(key, []).append(record)
            for (contract_type, expiration), records in grouped.items():
                records_sorted = sorted(
                    records,
                    key=lambda r: float(r.get("strike_price") or 0),
                )
                for lower, higher in zip(records_sorted, records_sorted[1:]):
                    if contract_type == "PUT":
                        long_leg = higher
                        short_leg = lower
                    else:
                        long_leg = lower
                        short_leg = higher
                    premium_long = option_mid_price(long_leg)
                    premium_short = option_mid_price(short_leg)
                    net_premium = None
                    if isinstance(premium_long, (int, float)) or isinstance(
                        premium_short, (int, float)
                    ):
                        net_premium = (premium_long or 0) - (premium_short or 0)
                        net_premium = abs(net_premium)
                    likelihood_values = [
                        option_likelihood(long_leg),
                        option_likelihood(short_leg),
                    ]
                    likelihoods = [
                        value for value in likelihood_values if isinstance(value, (int, float))
                    ]
                    likelihood = sum(likelihoods) / len(likelihoods) if likelihoods else None
                    strike_pair = "{low}/{high}".format(
                        low=format_strike(lower.get("strike_price")),
                        high=format_strike(higher.get("strike_price")),
                    )
                    spreads.append(
                        {
                            "strategy": strategy,
                            "ticker": lower.get("ticker") or higher.get("ticker"),
                            "expiration_label": expiration,
                            "contract_type": contract_type,
                            "strike_pair": strike_pair,
                            "long_leg": long_leg,
                            "short_leg": short_leg,
                            "net_premium": net_premium,
                            "likelihood": likelihood,
                            "greeks": combine_greeks(long_leg, short_leg),
                        }
                    )
        return spreads

    def _get_filter_value(self, var: tk.StringVar) -> str | None:
        value = var.get()
        return None if value == "All" else value

    def _record_matches_filters(self, record: dict, filters: dict[str, str | None]) -> bool:
        expiration = record.get("expiration_label")
        strike = record.get("strike_pair")
        contract_type = normalize_contract_type(record.get("contract_type"))
        if filters.get("expiration") and filters["expiration"] != expiration:
            return False
        if filters.get("strike") and filters["strike"] != strike:
            return False
        if filters.get("type") and filters["type"] != contract_type:
            return False
        return True

    def _record_matches_constraints(self, record: dict) -> bool:
        max_loss = parse_float(self.max_loss_var.get())
        min_likelihood = normalize_likelihood_threshold(self.likelihood_var.get())
        premium = record.get("net_premium")
        likelihood = record.get("likelihood")
        if max_loss is not None:
            if not isinstance(premium, (int, float)) or premium > max_loss:
                return False
        if min_likelihood is not None:
            if not isinstance(likelihood, (int, float)) or likelihood < min_likelihood:
                return False
        return True

    def _compute_filter_options(
        self, records: list[dict], current: dict[str, str | None]
    ) -> dict[str, list[str]]:
        options: dict[str, set[str]] = {"expiration": set(), "strike": set(), "type": set()}
        for record in records:
            expiration = record.get("expiration_label")
            strike = record.get("strike_pair")
            contract_type = normalize_contract_type(record.get("contract_type"))
            if self._record_matches_filters(record, {**current, "expiration": None}):
                if expiration:
                    options["expiration"].add(expiration)
            if self._record_matches_filters(record, {**current, "strike": None}):
                if strike:
                    options["strike"].add(strike)
            if self._record_matches_filters(record, {**current, "type": None}):
                if contract_type:
                    options["type"].add(contract_type)
        return {
            "expiration": sorted(options["expiration"]),
            "strike": sorted(options["strike"]),
            "type": sorted(options["type"]),
        }

    def _refresh_spread_filters(self, reset: bool = False) -> None:
        if reset:
            self.expiration_var.set("All")
            self.strike_var.set("All")
            self.type_var.set("All")
        eligible_records = [
            record for record in self.all_spread_records if self._record_matches_constraints(record)
        ]
        filters = {
            "expiration": self._get_filter_value(self.expiration_var),
            "strike": self._get_filter_value(self.strike_var),
            "type": self._get_filter_value(self.type_var),
        }
        options = self._compute_filter_options(eligible_records, filters)
        for key, dropdown, var in (
            ("expiration", self.expiration_dropdown, self.expiration_var),
            ("strike", self.strike_dropdown, self.strike_var),
            ("type", self.type_dropdown, self.type_var),
        ):
            values = ["All"] + options[key]
            dropdown["values"] = values
            if var.get() not in values:
                var.set("All")
        self._apply_spread_filters(eligible_records)

    def _apply_spread_filters(self, eligible_records: list[dict]) -> None:
        filters = {
            "expiration": self._get_filter_value(self.expiration_var),
            "strike": self._get_filter_value(self.strike_var),
            "type": self._get_filter_value(self.type_var),
        }
        self.spread_records = [
            record
            for record in eligible_records
            if self._record_matches_filters(record, filters)
        ]
        self.options_list.delete(0, tk.END)
        if not self.spread_records:
            self.options_list.insert(tk.END, "No spreads returned.")
            self.selected_spread = None
        else:
            for spread in self.spread_records:
                likelihood = spread.get("likelihood")
                likelihood_label = (
                    f"{likelihood * 100:.0f}%" if isinstance(likelihood, (int, float)) else "--"
                )
                premium = spread.get("net_premium")
                premium_label = f"{premium:.2f}" if isinstance(premium, (int, float)) else "--"
                self.options_list.insert(
                    tk.END,
                    "{ticker} {expiration} {type} {strike} | Loss {loss} | Likely {likelihood}".format(
                        ticker=spread.get("ticker", "--"),
                        expiration=spread.get("expiration_label", "--"),
                        type=str(spread.get("contract_type", "--")).upper(),
                        strike=spread.get("strike_pair", "--"),
                        loss=premium_label,
                        likelihood=likelihood_label,
                    ),
                )
            self.options_list.selection_set(0)
            self.options_list.see(0)
            self.selected_spread = self.spread_records[0]
        self._sync_spread_snapshot()

    def _sync_spread_snapshot(self) -> None:
        spread = self.selected_spread or {}
        self._set_value(self.spread_values["strategy"], spread.get("strategy"))
        self._set_value(self.spread_values["expiration"], spread.get("expiration_label"))
        contract_type = normalize_contract_type(spread.get("contract_type"))
        self._set_value(self.spread_values["type"], contract_type)
        self._set_value(self.spread_values["strikes"], spread.get("strike_pair"))
        premium = spread.get("net_premium")
        self._set_value(self.spread_values["premium"], premium)
        likelihood = spread.get("likelihood")
        likelihood_label = (
            f"{likelihood * 100:.1f}%" if isinstance(likelihood, (int, float)) else None
        )
        self._set_value(self.spread_values["likelihood"], likelihood_label)
        long_leg = spread.get("long_leg") or {}
        short_leg = spread.get("short_leg") or {}
        long_label = "{ticker} {exp} {type} {strike}".format(
            ticker=long_leg.get("ticker", "--"),
            exp=long_leg.get("expiration_date", "--"),
            type=str(long_leg.get("contract_type", "--")).upper(),
            strike=long_leg.get("strike_price", "--"),
        )
        short_label = "{ticker} {exp} {type} {strike}".format(
            ticker=short_leg.get("ticker", "--"),
            exp=short_leg.get("expiration_date", "--"),
            type=str(short_leg.get("contract_type", "--")).upper(),
            strike=short_leg.get("strike_price", "--"),
        )
        self._set_value(self.leg_values["long_leg"], long_label)
        self._set_value(self.leg_values["short_leg"], short_label)
        greeks = spread.get("greeks") or {}
        self._set_value(self.greeks_values["delta"], greeks.get("delta"))
        self._set_value(self.greeks_values["gamma"], greeks.get("gamma"))
        self._set_value(self.greeks_values["theta"], greeks.get("theta"))
        self._set_value(self.greeks_values["vega"], greeks.get("vega"))
        self._set_value(self.greeks_values["rho"], greeks.get("rho"))
        self._set_value(self.greeks_values["iv"], greeks.get("iv"))

    def on_spread_select(self, _event: object) -> None:
        selection = self.options_list.curselection()
        if not selection:
            return
        index = selection[0]
        if index >= len(self.spread_records):
            return
        self.selected_spread = self.spread_records[index]
        self._sync_spread_snapshot()

    def on_filter_change(self, _event: object) -> None:
        self._refresh_spread_filters()

    def _show_api_error(self, exc: HTTPError, service: str, hint: str | None = None) -> None:
        detail = format_http_error_detail(exc)
        detail_msg = f"\nDetails: {detail}" if detail else ""
        hint_msg = f"\n{hint}" if hint else ""
        self._show_error_dialog(
            "API Error",
            f"{service} API returned an error: {exc.code} {exc.reason}.{detail_msg}{hint_msg}",
        )

    def _show_error_dialog(self, title: str, message: str) -> None:
        dialog = tk.Toplevel(self)
        dialog.title(title)
        dialog.geometry("620x320")
        dialog.transient(self)

        dialog.rowconfigure(0, weight=1)
        dialog.columnconfigure(0, weight=1)

        text_frame = ttk.Frame(dialog)
        text_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

        text_widget = tk.Text(text_frame, wrap="word")
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        text_widget.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        text_widget.insert("1.0", message)
        text_widget.focus_set()

        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=1, column=0, pady=(0, 10))

        def copy_to_clipboard() -> None:
            dialog.clipboard_clear()
            dialog.clipboard_append(text_widget.get("1.0", "end-1c"))
            dialog.update_idletasks()

        ttk.Button(button_frame, text="Copy", command=copy_to_clipboard).grid(
            row=0, column=0, padx=5
        )
        ttk.Button(button_frame, text="Close", command=dialog.destroy).grid(
            row=0, column=1, padx=5
        )

if __name__ == "__main__":
    app = StoptionsApp()
    app.mainloop()

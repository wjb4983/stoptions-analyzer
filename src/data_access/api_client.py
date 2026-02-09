import json
from datetime import date, datetime, time as dt_time, timedelta
from urllib.parse import urlencode
from urllib.request import urlopen
from zoneinfo import ZoneInfo

from config import API_BASE_URL


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

    def _request_url(self, url: str) -> dict:
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
        start_dt = datetime.combine(start_date, dt_time.min, tzinfo=ZoneInfo("America/New_York"))
        end_dt = datetime.combine(
            end_date, dt_time.max.replace(microsecond=0), tzinfo=ZoneInfo("America/New_York")
        )
        start_ms = int(start_dt.timestamp() * 1000)
        end_ms = int(end_dt.timestamp() * 1000)
        params = {"adjusted": "true", "sort": "asc", "limit": "50000"}
        data = self._request(
            f"/v2/aggs/ticker/{ticker}/range/{minutes_per_bar}/minute/{start_ms}/{end_ms}",
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

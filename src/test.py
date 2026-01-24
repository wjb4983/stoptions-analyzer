import json
import math
import os
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Dict, List, Optional, Any
from urllib.parse import urlencode
from urllib.request import urlopen
from zoneinfo import ZoneInfo

import numpy as np
from scipy.stats import norm  # pip install scipy

# ============================
# Config
# ============================
API_BASE_URL = os.getenv("MASSIVE_BASE_URL", "https://api.polygon.io").rstrip("/")


from pathlib import Path
CONFIG_DIR = Path.home() / ".stoptions_analyzer"
API_KEY_PATH = CONFIG_DIR / "api_key.txt"
DATA_DIR = Path(__file__).resolve().parent / "data"

def load_api_key() -> str:
    env_key = os.getenv("MASSIVE_API_KEY", "").strip()
    if env_key:
        return env_key
    if API_KEY_PATH.exists():
        return API_KEY_PATH.read_text().strip()
    return ""

# API_KEY = os.getenv("MASSIVE_API_KEY", "").strip()
API_KEY  = load_api_key()

USE_AMERICAN = False       # American binomial pricing (CRR)
BINOMIAL_STEPS = 200      # Increase for accuracy, decrease for speed


# ============================
# Date helpers (copied style from your code)
# ============================
def effective_market_date() -> date:
    """If before 4pm ET, use yesterday; else today."""
    now = datetime.now(ZoneInfo("America/New_York"))
    market_close = now.replace(hour=16, minute=0, second=0, microsecond=0)
    if now < market_close:
        return (now - timedelta(days=1)).date()
    return now.date()


# ============================
# Massive/Polygon-compatible HTTP client (copied style from your code)
# ============================
class MassiveApiClient:
    def __init__(self, api_key: str, base_url: str = API_BASE_URL) -> None:
        self.api_key = api_key
        self.base_url = base_url.rstrip("/")

    def _request(self, path: str, params: Dict[str, str]) -> dict:
        params = {**params, "apiKey": self.api_key}
        url = f"{self.base_url}{path}?{urlencode(params)}"
        with urlopen(url, timeout=20) as response:
            payload = response.read().decode("utf-8")
        return json.loads(payload)

    def _request_url(self, url: str) -> dict:
        with urlopen(url, timeout=20) as response:
            payload = response.read().decode("utf-8")
        return json.loads(payload)

    def fetch_previous_close(self, ticker: str) -> dict:
        # EXACTLY like your working code
        data = self._request(f"/v2/aggs/ticker/{ticker}/prev", {"adjusted": "true"})
        result = (data.get("results") or [{}])[0]
        return {
            "close": result.get("c"),
            "open": result.get("o"),
            "high": result.get("h"),
            "low": result.get("l"),
            "volume": result.get("v"),
        }

    def fetch_option_snapshots(self, ticker: str, limit: int = 250) -> List[dict]:
        """Fetch /v3/snapshot/options/{ticker} and normalize to your schema."""
        results: List[dict] = []
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

    def _normalize_option_snapshots(self, snapshots: List[dict]) -> List[dict]:
        # Copied from your working code with minimal changes
        normalized: List[dict] = []
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
                    "bid": last_quote.get("bid") or last_quote.get("bid_price") or last_quote.get("bp"),
                    "ask": last_quote.get("ask") or last_quote.get("ask_price") or last_quote.get("ap"),
                    "last": last_trade.get("price") or last_trade.get("p"),
                }
            )
        return normalized


# ============================
# Option mid proxy (copied from your working code)
# ============================
def option_mid_price(contract: dict) -> Optional[float]:
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


def extract_expiry(contract: dict) -> Optional[date]:
    s = contract.get("expiration_date")
    if isinstance(s, date):
        return s
    if isinstance(s, str):
        try:
            return date.fromisoformat(s)
        except Exception:
            return None
    return None


def extract_iv(contract: dict) -> Optional[float]:
    greeks = contract.get("greeks") or {}
    iv = greeks.get("iv")
    if iv is None:
        iv = contract.get("implied_volatility")
    if isinstance(iv, (int, float)) and iv > 0:
        return float(iv) / 100.0 if iv > 3 else float(iv)
    return None


def extract_delta(contract: dict) -> Optional[float]:
    greeks = contract.get("greeks") or {}
    d = greeks.get("delta")
    return float(d) if isinstance(d, (int, float)) else None


# ============================
# Black-Scholes + American Binomial
# ============================
def bs_price(S: float, K: float, tau: float, r: float, q: float, sigma: float, is_call: bool) -> float:
    if tau <= 0:
        return max(0.0, S - K) if is_call else max(0.0, K - S)
    if sigma <= 0:
        forward = S * math.exp((r - q) * tau)
        intrinsic_fwd = max(0.0, forward - K) if is_call else max(0.0, K - forward)
        return intrinsic_fwd * math.exp(-r * tau)

    d1 = (math.log(S / K) + (r - q + 0.5 * sigma * sigma) * tau) / (sigma * math.sqrt(tau))
    d2 = d1 - sigma * math.sqrt(tau)
    if is_call:
        return S * math.exp(-q * tau) * norm.cdf(d1) - K * math.exp(-r * tau) * norm.cdf(d2)
    return K * math.exp(-r * tau) * norm.cdf(-d2) - S * math.exp(-q * tau) * norm.cdf(-d1)


def american_binomial_crr(S: float, K: float, tau: float, r: float, q: float, sigma: float, is_call: bool, steps: int = 200) -> float:
    if tau <= 0:
        return max(0.0, S - K) if is_call else max(0.0, K - S)
    if sigma <= 0:
        return max(0.0, S - K) if is_call else max(0.0, K - S)

    dt = tau / steps
    u = math.exp(sigma * math.sqrt(dt))
    d = 1.0 / u
    disc = math.exp(-r * dt)
    p = (math.exp((r - q) * dt) - d) / (u - d)
    p = min(1.0, max(0.0, p))

    prices = np.array([S * (u ** j) * (d ** (steps - j)) for j in range(steps + 1)], dtype=float)
    values = np.maximum(0.0, prices - K) if is_call else np.maximum(0.0, K - prices)

    for _ in range(steps - 1, -1, -1):
        prices = prices[:-1] / u
        cont = disc * (p * values[1:] + (1 - p) * values[:-1])
        exer = np.maximum(0.0, prices - K) if is_call else np.maximum(0.0, K - prices)
        values = np.maximum(cont, exer)

    return float(values[0])


def option_model_price(S: float, K: float, tau: float, r: float, q: float, sigma: float, is_call: bool) -> float:
    if USE_AMERICAN:
        return american_binomial_crr(S, K, tau, r, q, sigma, is_call, steps=BINOMIAL_STEPS)
    return bs_price(S, K, tau, r, q, sigma, is_call)


# ============================
# IV surface + Monte Carlo spot
# ============================
def build_iv_surface(contracts: List[dict]) -> Dict[date, Dict[float, float]]:
    surf: Dict[date, Dict[float, float]] = {}
    for c in contracts:
        exp = extract_expiry(c)
        K = c.get("strike_price")
        iv = extract_iv(c)
        if exp and isinstance(K, (int, float)) and iv:
            surf.setdefault(exp, {})[float(K)] = float(iv)
    return surf


def nearest_iv(iv_by_strike: Dict[float, float], K: float) -> float:
    strikes = np.array(sorted(iv_by_strike.keys()))
    idx = int(np.argmin(np.abs(strikes - K)))
    return float(iv_by_strike[float(strikes[idx])])


def simulate_spot_gbm(S0: float, t_years: float, sigma: float, drift: float, n: int, seed: int = 7) -> np.ndarray:
    rng = np.random.default_rng(seed)
    Z = rng.standard_normal(n)
    return S0 * np.exp((drift - 0.5 * sigma * sigma) * t_years + sigma * math.sqrt(t_years) * Z)


# ============================
# Position representation
# ============================
@dataclass(frozen=True)
class OptionLeg:
    ticker: str
    is_call: bool
    strike: float
    expiry: date
    qty: int
    iv: float
    entry_mid: float  # uses option_mid_price() which prefers bid/ask, then last, then day_close


@dataclass
class Position:
    underlying: str
    legs: List[OptionLeg]
    shares_per_contract: int = 100


def position_entry_cost(pos: Position) -> float:
    return sum(leg.qty * leg.entry_mid * pos.shares_per_contract for leg in pos.legs)


def position_value_at(
    pos: Position,
    S: float,
    asof: date,
    r: float,
    q: float,
    iv_surface_by_expiry: Dict[date, Dict[float, float]],
) -> float:
    total = 0.0
    for leg in pos.legs:
        tau_days = (leg.expiry - asof).days
        tau = max(0.0, tau_days / 365.0)

        iv_map = iv_surface_by_expiry.get(leg.expiry, {})
        sigma = nearest_iv(iv_map, leg.strike) if iv_map else leg.iv

        px = option_model_price(S=S, K=leg.strike, tau=tau, r=r, q=q, sigma=sigma, is_call=leg.is_call)
        total += leg.qty * px * pos.shares_per_contract
    return total


# ============================
# Build example spread from normalized snapshot dicts
# ============================
def choose_put_credit_spread(contracts: List[dict], S0: float, r: float = 0.03, q: float = 0.0) -> Position:
    today = date.today()

    puts = []
    for c in contracts:
        if (c.get("contract_type") or "").lower() != "put":
            continue
        exp = extract_expiry(c)
        K = c.get("strike_price")
        iv = extract_iv(c)
        delta = extract_delta(c)
        tkr = c.get("ticker")
        if tkr and exp and isinstance(K, (int, float)) and iv:
            puts.append((tkr, exp, float(K), float(iv), delta, c))

    if not puts:
        raise RuntimeError("No usable put contracts (need ticker, expiry, strike, IV).")

    target_days = 45
    expiries = sorted(set(exp for _, exp, *_ in puts))
    expiry = min(expiries, key=lambda e: abs((e - today).days - target_days))

    puts_e = [(tkr, K, iv, delta, raw) for (tkr, exp, K, iv, delta, raw) in puts if exp == expiry]
    puts_e.sort(key=lambda x: x[1])

    with_delta = [(tkr, K, iv, d, raw) for (tkr, K, iv, d, raw) in puts_e if isinstance(d, (int, float))]
    if with_delta:
        t_short, K_short, iv_short, _, raw_short = min(with_delta, key=lambda x: abs(abs(x[3]) - 0.20))
    else:
        target_short = 0.93 * S0
        t_short, K_short, iv_short, _, raw_short = min(puts_e, key=lambda x: abs(x[1] - target_short))

    target_long = K_short - 10.0
    t_long, K_long, iv_long, _, raw_long = min(puts_e, key=lambda x: abs(x[1] - target_long))

    mid_short = option_mid_price(raw_short)
    mid_long = option_mid_price(raw_long)

    # If day_close is missing (rare), fall back to model price at entry
    if mid_short is None:
        tau = max(0.0, (expiry - effective_market_date()).days / 365.0)
        mid_short = option_model_price(S0, K_short, tau, r, q, iv_short, is_call=False)
    if mid_long is None:
        tau = max(0.0, (expiry - effective_market_date()).days / 365.0)
        mid_long = option_model_price(S0, K_long, tau, r, q, iv_long, is_call=False)

    legs = [
        OptionLeg(ticker=t_short, is_call=False, strike=K_short, expiry=expiry, qty=-1, iv=iv_short, entry_mid=float(mid_short)),
        OptionLeg(ticker=t_long,  is_call=False, strike=K_long,  expiry=expiry, qty=+1, iv=iv_long,  entry_mid=float(mid_long)),
    ]
    return Position(underlying="NVDA", legs=legs)


# ============================
# Horizon analytics
# ============================
def analyze_position_over_horizons(
    pos: Position,
    iv_surface_by_expiry: Dict[date, Dict[float, float]],
    S0: float,
    r: float = 0.03,
    q: float = 0.0,
    use_risk_neutral: bool = True,
    n_sims: int = 50_000,
    horizons_days: List[int] = [7, 30, 365],
):
    today = effective_market_date()
    entry_cost = position_entry_cost(pos)

    chosen_expiry = pos.legs[0].expiry
    iv_map = iv_surface_by_expiry.get(chosen_expiry, {})
    atm_iv = nearest_iv(iv_map, S0) if iv_map else max(0.20, pos.legs[0].iv)

    drift = (r - q) if use_risk_neutral else 0.08

    results = []
    for hd in horizons_days:
        t = hd / 365.0
        asof = today + timedelta(days=hd)

        S_t = simulate_spot_gbm(S0=S0, t_years=t, sigma=atm_iv, drift=drift, n=n_sims, seed=7 + hd)

        vals = np.array([position_value_at(pos, float(s), asof, r, q, iv_surface_by_expiry) for s in S_t])
        pnl = vals - entry_cost

        pop = float(np.mean(pnl >= 0.0))
        ev = float(np.mean(pnl))
        p50 = float(np.quantile(pnl, 0.50))
        var95 = float(np.quantile(pnl, 0.05))
        cvar95 = float(np.mean(pnl[pnl <= var95])) if np.any(pnl <= var95) else var95

        short_leg = [l for l in pos.legs if l.qty < 0][0]
        p_itm_short = float(np.mean(S_t < short_leg.strike)) if not short_leg.is_call else float(np.mean(S_t > short_leg.strike))

        results.append(
            {
                "horizon_days": hd,
                "POP(P/L>=0)": pop,
                "P(ITM short leg)": p_itm_short,
                "EV($)": ev,
                "Median($)": p50,
                "VaR95($)": var95,
                "CVaR95($)": cvar95,
            }
        )

    return results, atm_iv, entry_cost


# ============================
# Main
# ============================
def main():
    if not API_KEY:
        raise RuntimeError("Set MASSIVE_API_KEY in your environment.")

    api = MassiveApiClient(API_KEY)

    # Stock close via /v2/aggs/ticker/NVDA/prev (your working approach)
    prev = api.fetch_previous_close("NVDA")
    S0 = prev.get("close")
    if not isinstance(S0, (int, float)) or S0 <= 0:
        raise RuntimeError("Could not fetch NVDA previous close via /v2/aggs/ticker/NVDA/prev")

    contracts = api.fetch_option_snapshots("NVDA", limit=250)
    if not contracts:
        raise RuntimeError("Empty option snapshot results for NVDA.")

    iv_surface = build_iv_surface(contracts)
    pos = choose_put_credit_spread(contracts, S0=float(S0))

    rows, atm_iv, entry_cost = analyze_position_over_horizons(
        pos=pos,
        iv_surface_by_expiry=iv_surface,
        S0=float(S0),
        r=0.03,
        q=0.00,
        use_risk_neutral=True,
        n_sims=50_000,
        horizons_days=[7, 30, 365],
    )

    print(f"\nS0 (NVDA prev close): {float(S0):.2f}")
    print("\n=== Example Position (NVDA) ===")
    print(f"Entry cost (positive=debit, negative=credit): ${entry_cost:,.2f}")
    for leg in pos.legs:
        side = "LONG" if leg.qty > 0 else "SHORT"
        typ = "CALL" if leg.is_call else "PUT"
        print(f"{side:5s} {typ} {leg.ticker} K={leg.strike:.2f} exp={leg.expiry.isoformat()} iv={leg.iv:.3f} entry_mid_proxy={leg.entry_mid:.2f}")

    print(f"\nATM IV used for spot simulation (quick & dirty): {atm_iv:.3f}")
    print("\n=== Horizon Stats ===")
    for r_ in rows:
        print(
            f"{r_['horizon_days']:4d}d | POP={r_['POP(P/L>=0)']:.3f} | "
            f"P(ITM short)={r_['P(ITM short leg)']:.3f} | EV=${r_['EV($)']:.2f} | "
            f"Median=${r_['Median($)']:.2f} | VaR95=${r_['VaR95($)']:.2f} | CVaR95=${r_['CVaR95($)']:.2f}"
        )


if __name__ == "__main__":
    main()

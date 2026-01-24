"""
All-options EV + risk scanner (naked long calls/puts) with early-exit policies.

What it does (for ONE underlying, e.g. NVDA):
- Pulls option snapshots (/v3/snapshot/options/{ticker}) -> strike, expiry, type, IV, day_close (if present)
- Pulls underlying prev close (/v2/aggs/ticker/{ticker}/prev) -> S0
- Simulates many joint paths of (S_t, IV_factor_t) once per horizon (daily steps)
- Prices ALL options in batches with vectorized European Black–Scholes (GPU-ready structure, pure NumPy)
- Applies an exit policy (profit target / stop loss / max-hold) on each path
- Computes distributions and summary stats: EV, POP, VaR/CVaR, hit rates, expected hold, expected exit price, etc.
- Ranks and prints “best” options under max-loss (premium) and tail-risk constraints.

Notes:
- Because you often won’t have tradable quotes, entry uses snapshot day_close when available, else model price.
- Exit uses model mark (with optional slippage model).
- American: included as a stub hook (no-op) so you can add later.

Requires:
  pip install numpy scipy
Set:
  MASSIVE_API_KEY
  MASSIVE_BASE_URL (optional; defaults to https://api.polygon.io)
"""

import json
import math
import os
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Dict, List, Optional, Tuple
from urllib.parse import urlencode
from urllib.request import urlopen
from zoneinfo import ZoneInfo

from pathlib import Path
import numpy as np
from scipy.special import ndtr  # fast Normal CDF


# ============================
# User-tunable configuration
# ============================
UNDERLYING = os.getenv("UNDERLYING", "NVDA").upper()

API_BASE_URL = os.getenv("MASSIVE_BASE_URL", "https://api.polygon.io").rstrip("/")
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

# Simulation
N_PATHS = int(os.getenv("N_PATHS", "50000"))           # samples
DT_DAYS = 1                                            # daily steps
R = float(os.getenv("R", "0.03"))                      # annual risk-free
Q = float(os.getenv("Q", "0.00"))                      # dividend yield

# Horizons (max hold)
HORIZONS = [
    {"name": "1D",  "max_hold_days": 1},
    {"name": "1W",  "max_hold_days": 7},
    {"name": "1M",  "max_hold_days": 30},
]

# Exit policy (applied to option P/L relative to entry premium)
# Example: +25% take-profit, -50% stop-loss, or max hold.
TAKE_PROFIT_PCT = float(os.getenv("TAKE_PROFIT_PCT", "0.25"))  # +25%
STOP_LOSS_PCT   = float(os.getenv("STOP_LOSS_PCT",   "0.50"))  # -50%

# Slippage model (very simple; you can upgrade)
# Applied at entry and exit as a "half-spread" fraction of option price.
HALF_SPREAD_FRAC = float(os.getenv("HALF_SPREAD_FRAC", "0.02"))  # 2% half-spread

# Filters
MIN_DTE = int(os.getenv("MIN_DTE", "7"))               # ignore expiring too soon
MAX_DTE = int(os.getenv("MAX_DTE", "90"))              # ignore too far out (speed + relevance)
MIN_IV  = float(os.getenv("MIN_IV", "0.05"))           # ignore garbage IV
MAX_PREMIUM = float(os.getenv("MAX_PREMIUM", "5000"))  # max loss per contract ($) you tolerate

# Ranking
TOP_N = int(os.getenv("TOP_N", "30"))
RISK_METRIC = os.getenv("RISK_METRIC", "CVaR95_$")  # "CVaR95" or "VaR95"


# ============================
# Market date helper (same logic you used)
# ============================
def effective_market_date() -> date:
    now = datetime.now(ZoneInfo("America/New_York"))
    market_close = now.replace(hour=16, minute=0, second=0, microsecond=0)
    if now < market_close:
        return (now - timedelta(days=1)).date()
    return now.date()


# ============================
# API client (Polygon-compatible, same style as your working code)
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
        data = self._request(f"/v2/aggs/ticker/{ticker}/prev", {"adjusted": "true"})
        result = (data.get("results") or [{}])[0]
        return {"close": result.get("c"), "open": result.get("o"), "high": result.get("h"), "low": result.get("l"), "volume": result.get("v")}

    def fetch_option_snapshots(self, ticker: str, limit: int = 250) -> List[dict]:
        results: List[dict] = []
        data = self._request(f"/v3/snapshot/options/{ticker}", {"limit": str(limit)})
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

            normalized.append(
                {
                    "ticker": details.get("ticker") or snapshot.get("ticker"),
                    "expiration_date": details.get("expiration_date"),
                    "contract_type": details.get("contract_type"),
                    "strike_price": details.get("strike_price"),
                    "greeks": greeks,
                    "implied_volatility": implied_vol,
                    "day_close": snapshot.get("close") or day.get("close") or day.get("c"),
                    "bid": last_quote.get("bid") or last_quote.get("bid_price") or last_quote.get("bp"),
                    "ask": last_quote.get("ask") or last_quote.get("ask_price") or last_quote.get("ap"),
                    "last": last_trade.get("price") or last_trade.get("p"),
                }
            )
        return normalized


# ============================
# Pricing: vectorized European BS
# ============================
def bs_price_vec(S: np.ndarray, K: np.ndarray, tau: np.ndarray, r: float, q: float, sigma: np.ndarray, is_call: np.ndarray) -> np.ndarray:
    """
    Vectorized BS over broadcastable shapes.
    Inputs typically shaped:
      S:    (N, 1) or (N, T, 1)
      K:    (1, 1, C)
      tau:  (1, T, C)
      sigma:(N, T, C) or (1, T, C)
      is_call:(1, 1, C) boolean
    Returns: price with broadcasted shape.
    """
    eps = 1e-12
    tau_pos = np.maximum(tau, 0.0)
    sig = np.maximum(sigma, eps)

    # intrinsic for tau==0
    intrinsic_call = np.maximum(S - K, 0.0)
    intrinsic_put  = np.maximum(K - S, 0.0)
    intrinsic = np.where(is_call, intrinsic_call, intrinsic_put)

    # safe compute where tau>0
    sqrt_tau = np.sqrt(np.maximum(tau_pos, eps))
    d1 = (np.log(np.maximum(S, eps) / np.maximum(K, eps)) + (r - q + 0.5 * sig * sig) * tau_pos) / (sig * sqrt_tau)
    d2 = d1 - sig * sqrt_tau

    disc_r = np.exp(-r * tau_pos)
    disc_q = np.exp(-q * tau_pos)

    call = S * disc_q * ndtr(d1) - K * disc_r * ndtr(d2)
    put  = K * disc_r * ndtr(-d2) - S * disc_q * ndtr(-d1)

    price = np.where(is_call, call, put)
    return np.where(tau_pos <= 0.0, intrinsic, price)


# ============================
# American stub hook
# ============================
def american_price_stub(price_eur: np.ndarray, *args, **kwargs) -> np.ndarray:
    """
    Placeholder: return European price.
    You can replace this with an American correction later (grid/interp or BAW/CRR).
    """
    return price_eur


# ============================
# Exit policy application (vectorized)
# ============================
def apply_exit_policy(
    price_path: np.ndarray,      # (N, T+1) option mark over time
    entry_price: np.ndarray,     # (N,) entry per path (after slippage)
    take_profit_pct: float,
    stop_loss_pct: float,
) -> Tuple[np.ndarray, np.ndarray, np.ndarray, np.ndarray]:
    """
    Hybrid policy: exit at first time index where:
      price >= entry*(1+tp) OR price <= entry*(1-sl) OR end
    Returns:
      exit_idx (N,) int
      exit_price (N,)
      hit_tp (N,) bool
      hit_sl (N,) bool
    """
    N, T1 = price_path.shape
    tp_level = entry_price * (1.0 + take_profit_pct)
    sl_level = entry_price * (1.0 - stop_loss_pct)

    # boolean triggers across time
    hit_tp_mat = price_path >= tp_level[:, None]
    hit_sl_mat = price_path <= sl_level[:, None]

    # first hit index; if never hit, we'll set to last index
    any_tp = hit_tp_mat.any(axis=1)
    any_sl = hit_sl_mat.any(axis=1)

    first_tp = np.where(any_tp, hit_tp_mat.argmax(axis=1), T1 - 1)
    first_sl = np.where(any_sl, hit_sl_mat.argmax(axis=1), T1 - 1)

    # choose earliest of tp vs sl vs end
    exit_idx = np.minimum(first_tp, first_sl)
    exit_idx = np.minimum(exit_idx, T1 - 1)

    # if both never hit, exit at end already
    row_idx = np.arange(N)
    exit_price = price_path[row_idx, exit_idx]

    hit_tp = any_tp & (first_tp <= first_sl)
    hit_sl = any_sl & (first_sl < first_tp)

    return exit_idx, exit_price, hit_tp, hit_sl


# ============================
# Vol dynamics (fast, realistic enough)
# ============================
def simulate_joint_paths(
    S0: float,
    iv_atm0: float,
    max_days: int,
    n_paths: int,
    r: float,
    q: float,
    # vol dynamics parameters:
    rho: float = -0.6,      # corr between stock shock and iv shock (equity skew)
    kappa: float = 3.0,     # mean reversion speed (annual)
    volvol: float = 0.7,    # iv-of-iv (annual)
    beta_skew: float = -1.5 # iv responds to return: sigma *= exp(beta*log(S/S0))
) -> Tuple[np.ndarray, np.ndarray]:
    """
    Returns:
      S_paths: (N, T+1)
      iv_factor_paths: (N, T+1)   where sigma_t = iv0 * exp(iv_factor_t + beta*log(S_t/S0))
    """
    dt = DT_DAYS / 365.0
    T = max_days

    rng = np.random.default_rng(12345 + T)

    # correlated shocks
    z1 = rng.standard_normal((n_paths, T))
    z2 = rng.standard_normal((n_paths, T))
    eps_s = z1
    eps_v = rho * z1 + math.sqrt(max(1e-12, 1.0 - rho * rho)) * z2

    # simulate logS
    mu_rn = (r - q)
    logS = np.empty((n_paths, T + 1), dtype=np.float64)
    logS[:, 0] = math.log(S0)

    # simulate iv_factor as mean-reverting OU in log space around 0
    x = np.empty((n_paths, T + 1), dtype=np.float64)
    x[:, 0] = 0.0

    # use atm vol for stock diffusion
    sig_s = max(1e-6, iv_atm0)

    for t in range(T):
        # stock
        logS[:, t + 1] = logS[:, t] + (mu_rn - 0.5 * sig_s * sig_s) * dt + sig_s * math.sqrt(dt) * eps_s[:, t]
        # iv factor (log-vol factor)
        x[:, t + 1] = x[:, t] + (-kappa * x[:, t]) * dt + volvol * math.sqrt(dt) * eps_v[:, t]

    S_paths = np.exp(logS)
    return S_paths, x


# ============================
# Data model
# ============================
@dataclass(frozen=True)
class OptRec:
    ticker: str
    is_call: bool
    strike: float
    expiry: date
    iv0: float
    entry_mkt: Optional[float]  # day_close proxy if present


def parse_opt_records(raw: List[dict], today: date) -> List[OptRec]:
    out: List[OptRec] = []
    for c in raw:
        tkr = c.get("ticker")
        exp_s = c.get("expiration_date")
        typ = (c.get("contract_type") or "").lower()
        K = c.get("strike_price")
        iv = (c.get("greeks") or {}).get("iv")
        if iv is None:
            iv = c.get("implied_volatility")

        if not (isinstance(tkr, str) and tkr):
            continue
        if not isinstance(K, (int, float)):
            continue
        if not isinstance(iv, (int, float)):
            continue

        # normalize expiry
        try:
            exp = exp_s if isinstance(exp_s, date) else date.fromisoformat(exp_s)
        except Exception:
            continue

        dte = (exp - today).days
        if dte < MIN_DTE or dte > MAX_DTE:
            continue

        iv0 = float(iv) / 100.0 if float(iv) > 3 else float(iv)
        if iv0 < MIN_IV:
            continue

        is_call = True if typ == "call" else False if typ == "put" else None
        if is_call is None:
            continue

        entry_close = c.get("day_close")
        entry_mkt = float(entry_close) if isinstance(entry_close, (int, float)) and entry_close > 0 else None

        out.append(OptRec(ticker=tkr, is_call=is_call, strike=float(K), expiry=exp, iv0=iv0, entry_mkt=entry_mkt))
    return out


# ============================
# Scanner core: price all options in batches for a horizon
# ============================
def scan_all_options_for_horizon(
    opts: List[OptRec],
    S0: float,
    today: date,
    horizon_name: str,
    max_hold_days: int,
    atm_iv0: float,
    n_paths: int,
    r: float,
    q: float,
    take_profit_pct: float,
    stop_loss_pct: float,
    half_spread_frac: float,
    use_american_stub: bool = True,
    batch_size: int = 256,
) -> List[dict]:
    """
    Returns list of per-option summary dicts for this horizon.
    """
    # simulate shared paths once
    S_paths, iv_factor = simulate_joint_paths(
        S0=S0,
        iv_atm0=atm_iv0,
        max_days=max_hold_days,
        n_paths=n_paths,
        r=r,
        q=q,
    )  # shapes (N, T+1)

    N = n_paths
    T1 = max_hold_days + 1

    # precompute time grid dates for tau calculation
    # tau for each option depends on expiry vs (today + t_days)
    day_offsets = np.arange(T1, dtype=np.int32)
    asof_dates = [today + timedelta(days=int(d)) for d in day_offsets]

    results: List[dict] = []

    # chunk options to manage memory: compute (N, T1, C) prices at once
    for start in range(0, len(opts), batch_size):
        chunk = opts[start : start + batch_size]
        C = len(chunk)
        if C == 0:
            continue

        strikes = np.array([o.strike for o in chunk], dtype=np.float64)              # (C,)
        is_call = np.array([o.is_call for o in chunk], dtype=bool)                   # (C,)
        iv0s    = np.array([o.iv0 for o in chunk], dtype=np.float64)                 # (C,)

        # tau grid (T1, C)
        tau = np.empty((T1, C), dtype=np.float64)
        for ti, d in enumerate(asof_dates):
            tau[ti, :] = np.maximum(0.0, np.array([(o.expiry - d).days for o in chunk], dtype=np.float64) / 365.0)

        # Broadcast shapes
        S = S_paths[:, :, None]                 # (N, T1, 1)
        K = strikes[None, None, :]              # (1, 1, C)

        # sigma dynamics per option:
        # sigma_t(option) = iv0 * exp(iv_factor_t + beta*log(S_t/S0))
        beta = -1.5
        log_moneyness = np.log(np.maximum(S_paths, 1e-12) / max(S0, 1e-12))  # (N, T1)
        sigma = iv0s[None, None, :] * np.exp(iv_factor[:, :, None] + beta * log_moneyness[:, :, None])  # (N, T1, C)
        sigma = np.clip(sigma, 1e-4, 5.0)

        tau_b = tau[None, :, :]                 # (1, T1, C)
        is_call_b = is_call[None, None, :]      # (1, 1, C)

        # European BS (vectorized)
        eur_prices = bs_price_vec(S=S, K=K, tau=tau_b, r=r, q=q, sigma=sigma, is_call=is_call_b)

        # American stub hook
        prices = american_price_stub(eur_prices) if use_american_stub else eur_prices  # (N, T1, C)

        # entry price: prefer market day_close, else model at t=0 (path 0 time 0 is fine, but entry should be scalar)
        # We'll compute entry_model using S0 and tau0 and sigma0 (iv0, no factor, no skew).
        tau0 = tau[0, :]  # (C,)
        entry_model = bs_price_vec(
            S=np.full((1, 1, C), S0, dtype=np.float64),
            K=K,
            tau=tau0[None, None, :],
            r=r,
            q=q,
            sigma=iv0s[None, None, :],
            is_call=is_call_b,
        ).reshape(C)

        entry_mkt = np.array([o.entry_mkt if o.entry_mkt is not None else np.nan for o in chunk], dtype=np.float64)
        entry_base = np.where(np.isfinite(entry_mkt), entry_mkt, entry_model)  # (C,)

        # apply entry/exit slippage (buying -> pay worse)
        entry_fill = entry_base * (1.0 + half_spread_frac)  # (C,)
        exit_fill_prices = prices * (1.0 - half_spread_frac)  # (N, T1, C)

        # Apply exit policy per option in chunk
        # We do this in a loop over C (batch-size), but inside each we stay vectorized over N,T.
        # That’s fast because C is small (256) and N,T are in NumPy.
        for j, o in enumerate(chunk):
            path_px = exit_fill_prices[:, :, j]  # (N, T1)
            entry_j = np.full((N,), entry_fill[j], dtype=np.float64)

            exit_idx, exit_px, hit_tp, hit_sl = apply_exit_policy(
                price_path=path_px,
                entry_price=entry_j,
                take_profit_pct=take_profit_pct,
                stop_loss_pct=stop_loss_pct,
            )

            pnl = (exit_px - entry_j) * 100.0  # 1 contract
            pop = float(np.mean(pnl > 0.0))
            ev = float(np.mean(pnl))
            med = float(np.quantile(pnl, 0.50))
            var95 = float(np.quantile(pnl, 0.05))
            cvar95 = float(np.mean(pnl[pnl <= var95])) if np.any(pnl <= var95) else var95

            # max loss for a long option (premium paid incl slippage)
            max_loss = float(entry_fill[j] * 100.0)

            # hold time stats
            hold_days = exit_idx.astype(np.float64)  # index==days because dt=1 day
            eh = float(np.mean(hold_days))
            h50 = float(np.quantile(hold_days, 0.50))
            h90 = float(np.quantile(hold_days, 0.90))

            res = {
                "horizon": horizon_name,
                "ticker": o.ticker,
                "type": "CALL" if o.is_call else "PUT",
                "strike": o.strike,
                "expiry": o.expiry.isoformat(),
                "dte": int((o.expiry - today).days),
                "iv0": o.iv0,
                "entry_price_used": float(entry_fill[j]),
                "entry_source": "day_close" if o.entry_mkt is not None else "model",
                "max_loss_$": max_loss,
                "EV_$": ev,
                "POP": pop,
                "Median_$": med,
                "VaR95_$": var95,
                "CVaR95_$": cvar95,
                "P_hit_TP": float(np.mean(hit_tp)),
                "P_hit_SL": float(np.mean(hit_sl)),
                "E_exit_price": float(np.mean(exit_px)),
                "E_hold_days": eh,
                "P50_hold_days": h50,
                "P90_hold_days": h90,
            }
            results.append(res)

    return results


# ============================
# Ranking & display
# ============================
# def rank_and_print(results: List[dict], horizon_name: str):
#     # Filter by max-loss tolerance
#     filtered = [r for r in results if r["horizon"] == horizon_name and r["max_loss_$"] <= MAX_PREMIUM]
#     if not filtered:
#         print(f"\n[{horizon_name}] No results after filters (MAX_PREMIUM={MAX_PREMIUM}).")
#         return

#     # Sort by EV desc, then tail risk best (less negative CVaR/VaR)
#     def risk_key(x):
#         return x[RISK_METRIC]

#     filtered.sort(key=lambda x: (x["EV_$"], risk_key(x)), reverse=True)

#     # Print summary header
#     print(f"\n=== {UNDERLYING} | Horizon={horizon_name} | N={len(filtered)} (filtered) ===")
#     print(f"Exit policy: +{TAKE_PROFIT_PCT*100:.0f}% TP / -{STOP_LOSS_PCT*100:.0f}% SL / max-hold")
#     print(f"Slippage: ±{HALF_SPREAD_FRAC*100:.1f}% (half-spread) | Risk metric={RISK_METRIC}")
#     print(
#         "Rank | Type | Strike | DTE | IV  | Entry(src) | MaxLoss$ |   EV$ | POP  | "
#         "VaR95$ | CVaR95$ | P(TP) | P(SL) | EHold | Ticker"
#     )

#     top = filtered[:TOP_N]
#     for i, r in enumerate(top, 1):
#         print(
#             f"{i:4d} | {r['type']:4s} | {r['strike']:6.2f} | {r['dte']:3d} | {r['iv0']:.3f} | "
#             f"{r['entry_price_used']:.2f}({r['entry_source'][:3]}) | {r['max_loss_$']:8.0f} | "
#             f"{r['EV_$']:6.1f} | {r['POP']:.3f} | {r['VaR95_$']:6.1f} | {r['CVaR95_$']:7.1f} | "
#             f"{r['P_hit_TP']:.3f} | {r['P_hit_SL']:.3f} | {r['E_hold_days']:.1f} | {r['ticker']}"
#         )
def rank_and_print(results: List[dict], horizon_name: str):
    # Filter by max-loss tolerance
    filtered = [r for r in results if r["horizon"] == horizon_name and r["max_loss_$"] <= MAX_PREMIUM]
    if not filtered:
        print(f"\n[{horizon_name}] No results after filters (MAX_PREMIUM={MAX_PREMIUM}).")
        return

    # Allow flexible risk metric names
    metric_alias = {
        "CVaR95": "CVaR95_$",
        "VaR95": "VaR95_$",
        "EV": "EV_$",
        "Median": "Median_$",
        "POP": "POP",
        "MaxLoss": "max_loss_$",
    }
    metric_key = metric_alias.get(RISK_METRIC, RISK_METRIC)

    # Safety: if user sets something weird, fall back
    if metric_key not in filtered[0]:
        print(f"[WARN] RISK_METRIC='{RISK_METRIC}' not found; falling back to 'CVaR95_$'")
        metric_key = "CVaR95_$"

    # Sort by EV desc, then by risk metric (higher is "better" if it's less negative)
    filtered.sort(key=lambda x: (x["EV_$"], x[metric_key]), reverse=True)

    # Print summary header
    print(f"\n=== {UNDERLYING} | Horizon={horizon_name} | N={len(filtered)} (filtered) ===")
    print(f"Exit policy: +{TAKE_PROFIT_PCT*100:.0f}% TP / -{STOP_LOSS_PCT*100:.0f}% SL / max-hold")
    print(f"Slippage: ±{HALF_SPREAD_FRAC*100:.1f}% (half-spread) | Risk metric={metric_key}")
    print(
        "Rank | Type | Strike | DTE | IV  | Entry(src) | MaxLoss$ |   EV$ | POP  | "
        "VaR95$ | CVaR95$ | P(TP) | P(SL) | EHold | Ticker"
    )

    top = filtered[:TOP_N]
    for i, r in enumerate(top, 1):
        print(
            f"{i:4d} | {r['type']:4s} | {r['strike']:6.2f} | {r['dte']:3d} | {r['iv0']:.3f} | "
            f"{r['entry_price_used']:.2f}({r['entry_source'][:3]}) | {r['max_loss_$']:8.0f} | "
            f"{r['EV_$']:6.1f} | {r['POP']:.3f} | {r['VaR95_$']:6.1f} | {r['CVaR95_$']:7.1f} | "
            f"{r['P_hit_TP']:.3f} | {r['P_hit_SL']:.3f} | {r['E_hold_days']:.1f} | {r['ticker']}"
        )


# ============================
# Main
# ============================
def main():
    if not API_KEY:
        raise RuntimeError("Set MASSIVE_API_KEY in your environment.")

    api = MassiveApiClient(API_KEY)
    today = effective_market_date()

    prev = api.fetch_previous_close(UNDERLYING)
    S0 = prev.get("close")
    if not isinstance(S0, (int, float)) or S0 <= 0:
        raise RuntimeError(f"Could not fetch {UNDERLYING} previous close via /v2/aggs/ticker/{UNDERLYING}/prev")
    S0 = float(S0)

    raw_opts = api.fetch_option_snapshots(UNDERLYING, limit=250)
    opts = parse_opt_records(raw_opts, today=today)

    if not opts:
        raise RuntimeError("No options after parsing/filters. Try loosening MIN_DTE/MAX_DTE/MIN_IV.")

    # Choose an ATM-ish IV0 baseline for simulating the underlying diffusion:
    # pick median IV of near-the-money options (closest strikes).
    strikes = np.array([o.strike for o in opts], dtype=np.float64)
    ivs = np.array([o.iv0 for o in opts], dtype=np.float64)
    idx_atm = np.argsort(np.abs(strikes - S0))[:min(50, len(opts))]
    atm_iv0 = float(np.median(ivs[idx_atm])) if len(idx_atm) else float(np.median(ivs))
    atm_iv0 = max(0.10, min(2.50, atm_iv0))

    print(f"\nUnderlying: {UNDERLYING}")
    print(f"Market date: {today.isoformat()} | S0(prev close): {S0:.2f} | ATM_IV0~{atm_iv0:.3f}")
    print(f"Options considered: {len(opts)} (DTE {MIN_DTE}-{MAX_DTE}, IV>={MIN_IV})")
    print(f"Simulation: {N_PATHS} paths | daily steps | horizons: {[h['name'] for h in HORIZONS]}")
    print("Pricing: European BS (vectorized). American: stub hook.")

    all_results: List[dict] = []

    # Run horizon scans
    for h in HORIZONS:
        res = scan_all_options_for_horizon(
            opts=opts,
            S0=S0,
            today=today,
            horizon_name=h["name"],
            max_hold_days=h["max_hold_days"],
            atm_iv0=atm_iv0,
            n_paths=N_PATHS,
            r=R,
            q=Q,
            take_profit_pct=TAKE_PROFIT_PCT,
            stop_loss_pct=STOP_LOSS_PCT,
            half_spread_frac=HALF_SPREAD_FRAC,
            use_american_stub=True,
            batch_size=256,  # tune for RAM/speed
        )
        all_results.extend(res)
        rank_and_print(all_results, h["name"])

    # Save full results for later slicing/plotting
    out_path = f"{UNDERLYING}_all_options_scan.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(all_results, f, indent=2, default=str)
    print(f"\nSaved full results to: {out_path}")


if __name__ == "__main__":
    main()

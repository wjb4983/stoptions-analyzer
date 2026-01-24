"""
RAM-friendly ALL-options EV scanner (long calls/puts only) with two-phase workflow.

Phase A (SCAN, RAM-friendly):
- Simulate shared paths of (S_t, IV_factor_t) step-by-step (no S_paths matrix)
- Price options in small batches each step with vectorized European Black–Scholes
- Apply early-exit policy (TP/SL/max-hold) online
- Compute: EV, POP, median (approx), TP/SL hit rates, expected hold, expected exit price
- Keep only lightweight aggregates + Top-N ranking in memory
- Stream per-option summaries to disk as gzipped JSONL (one row per option)

Phase B (DEEP RISK on Top-K):
- Rerun only top-K options per horizon
- Store full P/L samples (K × N_paths) and compute VaR/CVaR/quantiles accurately
- Output a separate gzipped JSONL for deep stats

Defaults are chosen to be laptop-friendly without killing sample counts:
- Phase A: N_PATHS_SCAN=25000
- Phase B: N_PATHS_DEEP=100000 (for top K only)
- Float32 everywhere, small option batch size.

Set env vars (optional):
  MASSIVE_API_KEY
  MASSIVE_BASE_URL (default https://api.polygon.io)
  UNDERLYING (default NVDA)

Run:
  python test3.py
"""

import gzip
import json
import math
import os
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Dict, List, Optional, Tuple
from urllib.parse import urlencode
from urllib.request import urlopen
from zoneinfo import ZoneInfo

import numpy as np
from scipy.special import ndtr  # fast Normal CDF


# ============================
# Config
# ============================
UNDERLYING = os.getenv("UNDERLYING", "NVDA").upper()

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

# Rates / carry
R = float(os.getenv("R", "0.03"))
Q = float(os.getenv("Q", "0.00"))

# Horizons (max hold in days). Add longer ones if you want; this stays RAM-safe.
HORIZONS = [
    {"name": "1D", "max_hold_days": 1},
    {"name": "1W", "max_hold_days": 7},
    {"name": "1M", "max_hold_days": 30},
    # {"name": "3M", "max_hold_days": 90},   # works fine
    # {"name": "12M", "max_hold_days": 365}, # works, but time cost grows; still RAM-safe
]

# Exit policy on option *price* relative to entry fill
TAKE_PROFIT_PCT = float(os.getenv("TAKE_PROFIT_PCT", "0.25"))  # +25% target
STOP_LOSS_PCT = float(os.getenv("STOP_LOSS_PCT", "0.50"))      # -50% stop

# Slippage model (simple half-spread fraction)
HALF_SPREAD_FRAC = float(os.getenv("HALF_SPREAD_FRAC", "0.02"))

# Filters
MIN_DTE = int(os.getenv("MIN_DTE", "7"))
MAX_DTE = int(os.getenv("MAX_DTE", "90"))
MIN_IV = float(os.getenv("MIN_IV", "0.05"))
MAX_PREMIUM = float(os.getenv("MAX_PREMIUM", "5000"))  # max loss per contract ($)

# Phase A / Phase B samples
N_PATHS_SCAN = int(os.getenv("N_PATHS_SCAN", "25000"))
N_PATHS_DEEP = int(os.getenv("N_PATHS_DEEP", "100000"))

# Option chunking (RAM safe)
BATCH_SIZE = int(os.getenv("BATCH_SIZE", "64"))  # small for laptops

# Top lists
TOP_N_PRINT = int(os.getenv("TOP_N_PRINT", "30"))
TOP_K_DEEP = int(os.getenv("TOP_K_DEEP", "150"))  # deep risk computed for top K per horizon

# Deep risk quantiles
DEEP_Q = [0.01, 0.05, 0.10, 0.50, 0.90, 0.95, 0.99]
DEEP_VAR_Q = 0.05  # VaR95 is 5th percentile of P/L

# Vol dynamics (fast, decent realism)
RHO = float(os.getenv("RHO", "-0.6"))       # corr stock shock and iv shock
KAPPA = float(os.getenv("KAPPA", "3.0"))    # mean reversion speed (annual)
VOLVOL = float(os.getenv("VOLVOL", "0.7"))  # iv-of-iv (annual)
BETA_SKEW = float(os.getenv("BETA_SKEW", "-1.5"))  # IV responds to returns


# ============================
# Date helper
# ============================
def effective_market_date() -> date:
    now = datetime.now(ZoneInfo("America/New_York"))
    market_close = now.replace(hour=16, minute=0, second=0, microsecond=0)
    if now < market_close:
        return (now - timedelta(days=1)).date()
    return now.date()


# ============================
# API client (Polygon-compatible)
# ============================
class MassiveApiClient:
    def __init__(self, api_key: str, base_url: str = API_BASE_URL) -> None:
        self.api_key = api_key
        self.base_url = base_url.rstrip("/")

    def _request(self, path: str, params: Dict[str, str]) -> dict:
        params = {**params, "apiKey": self.api_key}
        url = f"{self.base_url}{path}?{urlencode(params)}"
        with urlopen(url, timeout=30) as response:
            payload = response.read().decode("utf-8")
        return json.loads(payload)

    def _request_url(self, url: str) -> dict:
        with urlopen(url, timeout=30) as response:
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
# Option record
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

        if typ == "call":
            is_call = True
        elif typ == "put":
            is_call = False
        else:
            continue

        entry_close = c.get("day_close")
        entry_mkt = float(entry_close) if isinstance(entry_close, (int, float)) and entry_close > 0 else None

        out.append(OptRec(ticker=tkr, is_call=is_call, strike=float(K), expiry=exp, iv0=iv0, entry_mkt=entry_mkt))
    return out


# ============================
# Pricing (European BS, vectorized)
# ============================
def bs_price_vec_1d(S: np.ndarray, K: np.ndarray, tau: np.ndarray, r: float, q: float, sigma: np.ndarray, is_call: np.ndarray) -> np.ndarray:
    """
    Vectorized BS for shapes:
      S:     (N,) float32/64
      K:     (C,) float32/64
      tau:   (C,) float32/64  (time to expiry at current day)
      sigma: (N,C) float32/64 (path-specific sigma per option)
      is_call:(C,) bool
    Returns:
      prices: (N,C)
    """
    eps = 1e-12
    S2 = np.maximum(S[:, None], eps)
    K2 = np.maximum(K[None, :], eps)
    tau2 = np.maximum(tau[None, :], 0.0)
    sig = np.maximum(sigma, 1e-6)

    intrinsic_call = np.maximum(S2 - K2, 0.0)
    intrinsic_put = np.maximum(K2 - S2, 0.0)
    intrinsic = np.where(is_call[None, :], intrinsic_call, intrinsic_put)

    # where tau == 0 -> intrinsic
    sqrt_tau = np.sqrt(np.maximum(tau2, eps))
    d1 = (np.log(S2 / K2) + (r - q + 0.5 * sig * sig) * tau2) / (sig * sqrt_tau)
    d2 = d1 - sig * sqrt_tau

    disc_r = np.exp(-r * tau2)
    disc_q = np.exp(-q * tau2)

    call = S2 * disc_q * ndtr(d1) - K2 * disc_r * ndtr(d2)
    put = K2 * disc_r * ndtr(-d2) - S2 * disc_q * ndtr(-d1)
    price = np.where(is_call[None, :], call, put)

    return np.where(tau2 <= 0.0, intrinsic, price)


# ============================
# American stub hook
# ============================
def american_stub(prices_eur: np.ndarray, *args, **kwargs) -> np.ndarray:
    return prices_eur


# ============================
# Simulation (streaming, no huge matrices)
# ============================
def simulate_streaming_paths(
    S0: float,
    atm_iv0: float,
    max_days: int,
    n_paths: int,
    r: float,
    q: float,
    rho: float,
    kappa: float,
    volvol: float,
    seed: int,
) -> Tuple[np.ndarray, np.ndarray]:
    """
    Generate correlated shocks for the entire horizon once:
      eps_s, eps_v shaped (N, T)
    Then you can step S and iv_factor forward cheaply in the loop.

    This keeps RAM to about:
      2 * N * T float32 (shock matrices)
    For N=25k, T=30 => ~6MB. Even for T=365 => ~73MB (still reasonable).
    """
    dt = 1.0 / 365.0
    T = max_days

    rng = np.random.default_rng(seed)
    z1 = rng.standard_normal((n_paths, T), dtype=np.float32)
    z2 = rng.standard_normal((n_paths, T), dtype=np.float32)

    eps_s = z1
    eps_v = rho * z1 + math.sqrt(max(1e-12, 1.0 - rho * rho)) * z2

    return eps_s, eps_v


# ============================
# Phase A: scan all options, streaming exit policy + aggregates
# ============================
def scan_horizon_phase_a(
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
    batch_size: int,
    out_jsonl_gz_path: str,
) -> List[dict]:
    """
    Returns Top-K candidates (light stats) to be used in Phase B.
    Also streams ALL option summaries (light stats) to gzipped JSONL.
    """
    dt = 1.0 / 365.0
    T = max_hold_days
    mu_rn = (r - q)
    sig_s = float(max(1e-6, atm_iv0))

    eps_s, eps_v = simulate_streaming_paths(
        S0=S0,
        atm_iv0=atm_iv0,
        max_days=T,
        n_paths=n_paths,
        r=r,
        q=q,
        rho=RHO,
        kappa=KAPPA,
        volvol=VOLVOL,
        seed=1000 + T,
    )

    # streaming writer
    f_out = gzip.open(out_jsonl_gz_path, "wt", encoding="utf-8")

    # We'll keep a top list in memory (for Phase B selection)
    top_candidates: List[dict] = []

    def push_top(row: dict):
        # Only consider within max premium (max loss) filter
        if row["max_loss_$"] > MAX_PREMIUM:
            return
        top_candidates.append(row)
        # keep limited size
        if len(top_candidates) > (TOP_K_DEEP * 5):
            top_candidates.sort(key=lambda x: (x["EV_$"], x["POP"]), reverse=True)
            del top_candidates[(TOP_K_DEEP * 5):]

    # For each chunk of options, maintain per-path exit state (N,Cchunk) as bool/idx
    # This is the main memory cost per chunk; with N=25k, C=64 -> 1.6M booleans -> ~1.6MB
    for start in range(0, len(opts), batch_size):
        chunk = opts[start:start + batch_size]
        C = len(chunk)
        if C == 0:
            continue

        strikes = np.array([o.strike for o in chunk], dtype=np.float32)
        is_call = np.array([o.is_call for o in chunk], dtype=bool)
        iv0s = np.array([o.iv0 for o in chunk], dtype=np.float32)
        dtes0 = np.array([(o.expiry - today).days for o in chunk], dtype=np.int32)

        # Entry price per option: market day_close else model at t=0
        entry_mkt = np.array([o.entry_mkt if o.entry_mkt is not None else np.nan for o in chunk], dtype=np.float32)

        tau0 = np.maximum(dtes0.astype(np.float32) / 365.0, 0.0)
        # entry model (use S0 scalar replicated in S vector)
        # For model entry, we don't need per-path sigma; it's iv0s and S0.
        # We'll compute with a dummy sigma matrix of shape (1,C) then broadcast.
        S_dummy = np.array([S0], dtype=np.float32)
        sigma_dummy = (iv0s[None, :]).astype(np.float32)
        entry_model = bs_price_vec_1d(
            S=S_dummy,
            K=strikes,
            tau=tau0,
            r=r,
            q=q,
            sigma=sigma_dummy,
            is_call=is_call,
        ).reshape(C).astype(np.float32)

        entry_base = np.where(np.isfinite(entry_mkt), entry_mkt, entry_model)
        entry_fill = entry_base * (1.0 + half_spread_frac)  # buy worse
        max_loss = entry_fill * 100.0

        # Prepare exit state
        active = np.ones((n_paths, C), dtype=bool)
        exit_day = np.full((n_paths, C), T, dtype=np.int16)         # default exit at end
        exit_price = np.zeros((n_paths, C), dtype=np.float32)
        hit_tp = np.zeros((n_paths, C), dtype=bool)
        hit_sl = np.zeros((n_paths, C), dtype=bool)

        # Initialize state
        logS = np.full((n_paths,), math.log(S0), dtype=np.float32)
        x = np.zeros((n_paths,), dtype=np.float32)  # iv_factor log

        # Price at day 0
        S_t = np.exp(logS).astype(np.float32)

        # sigma_t per path+option at day 0
        # sigma = iv0 * exp(x + beta*log(S/S0))
        log_m = np.log(np.maximum(S_t, 1e-12) / max(S0, 1e-12)).astype(np.float32)  # (N,)
        sigma = iv0s[None, :] * np.exp((x[:, None]) + (BETA_SKEW * log_m)[:, None]).astype(np.float32)
        sigma = np.clip(sigma, 1e-4, 5.0).astype(np.float32)

        tau_t = np.maximum((dtes0 - 0).astype(np.float32) / 365.0, 0.0)
        px0 = bs_price_vec_1d(S_t, strikes, tau_t, r, q, sigma, is_call)
        px0 = american_stub(px0)  # no-op now
        px0 = (px0 * (1.0 - half_spread_frac)).astype(np.float32)  # exit worse if you sell

        tp_level = entry_fill[None, :] * (1.0 + take_profit_pct)
        sl_level = entry_fill[None, :] * (1.0 - stop_loss_pct)

        # apply exits at day 0? (usually none, but keep consistent)
        tp_hit_now = (px0 >= tp_level) & active
        sl_hit_now = (px0 <= sl_level) & active
        hit_now = tp_hit_now | sl_hit_now
        if hit_now.any():
            exit_day[hit_now] = 0
            exit_price[hit_now] = px0[hit_now]
            hit_tp[tp_hit_now] = True
            hit_sl[sl_hit_now & (~tp_hit_now)] = True
            active[hit_now] = False

        # Step through days 1..T
        sqrt_dt = math.sqrt(dt)
        for t in range(1, T + 1):
            if not active.any():
                break

            # evolve logS and iv factor x
            e_s = eps_s[:, t - 1]
            e_v = eps_v[:, t - 1]

            logS = logS + (mu_rn - 0.5 * sig_s * sig_s) * dt + sig_s * sqrt_dt * e_s
            x = x + (-KAPPA * x) * dt + VOLVOL * sqrt_dt * e_v

            S_t = np.exp(logS).astype(np.float32)

            # sigma per path+option
            log_m = np.log(np.maximum(S_t, 1e-12) / max(S0, 1e-12)).astype(np.float32)
            sigma = iv0s[None, :] * np.exp((x[:, None]) + (BETA_SKEW * log_m)[:, None]).astype(np.float32)
            sigma = np.clip(sigma, 1e-4, 5.0).astype(np.float32)

            tau_t = np.maximum((dtes0 - t).astype(np.float32) / 365.0, 0.0)
            px = bs_price_vec_1d(S_t, strikes, tau_t, r, q, sigma, is_call)
            px = american_stub(px)
            px = (px * (1.0 - half_spread_frac)).astype(np.float32)

            # apply exit logic only for active
            tp_hit = (px >= tp_level) & active
            sl_hit = (px <= sl_level) & active
            hit = tp_hit | sl_hit

            if hit.any():
                exit_day[hit] = t
                exit_price[hit] = px[hit]
                hit_tp[tp_hit] = True
                # if both hit same time, we treat TP as priority (common convention)
                hit_sl[sl_hit & (~tp_hit)] = True
                active[hit] = False

        # Aggregate per-option stats from exit arrays (reduce over paths)
        entry_fill_100 = (entry_fill * 100.0).astype(np.float32)  # (C,)
        pnl = (exit_price - entry_fill[None, :]) * 100.0          # (N,C)

        # Lightweight stats per option
        ev = pnl.mean(axis=0).astype(float)
        pop = (pnl > 0).mean(axis=0).astype(float)
        med = np.quantile(pnl, 0.50, axis=0).astype(float)  # ok in Phase A; no need for full distro storage
        p_tp = hit_tp.mean(axis=0).astype(float)
        p_sl = hit_sl.mean(axis=0).astype(float)
        e_hold = exit_day.mean(axis=0).astype(float)
        e_exit_px = exit_price.mean(axis=0).astype(float)

        # Write out all option summaries (Phase A)
        for j, o in enumerate(chunk):
            row = {
                "phase": "A",
                "horizon": horizon_name,
                "ticker": o.ticker,
                "type": "CALL" if o.is_call else "PUT",
                "strike": float(o.strike),
                "expiry": o.expiry.isoformat(),
                "dte": int(dtes0[j]),
                "iv0": float(o.iv0),
                "entry_price_used": float(entry_fill[j]),
                "entry_source": "day_close" if o.entry_mkt is not None else "model",
                "max_loss_$": float(max_loss[j] * 100.0 / 100.0),  # ensure float
                "EV_$": float(ev[j]),
                "POP": float(pop[j]),
                "Median_$": float(med[j]),
                "P_hit_TP": float(p_tp[j]),
                "P_hit_SL": float(p_sl[j]),
                "E_exit_price": float(e_exit_px[j]),
                "E_hold_days": float(e_hold[j]),
            }
            f_out.write(json.dumps(row) + "\n")
            push_top(row)

    f_out.close()

    # Select top K for Phase B
    top_candidates.sort(key=lambda x: (x["EV_$"], x["POP"]), reverse=True)
    top_k = top_candidates[:TOP_K_DEEP]
    return top_k


# ============================
# Phase B: deep risk for top K (store full pnl for accurate VaR/CVaR/quantiles)
# ============================
def deep_risk_phase_b(
    top_k: List[dict],
    opt_lookup: Dict[str, OptRec],
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
    out_jsonl_gz_path: str,
):
    if not top_k:
        return

    dt = 1.0 / 365.0
    T = max_hold_days
    mu_rn = (r - q)
    sig_s = float(max(1e-6, atm_iv0))

    eps_s, eps_v = simulate_streaming_paths(
        S0=S0,
        atm_iv0=atm_iv0,
        max_days=T,
        n_paths=n_paths,
        r=r,
        q=q,
        rho=RHO,
        kappa=KAPPA,
        volvol=VOLVOL,
        seed=5000 + T,
    )

    f_out = gzip.open(out_jsonl_gz_path, "wt", encoding="utf-8")

    # Run each option independently in Phase B (K is small; compute full pnl vector)
    # This is time-heavier but RAM-light and gives accurate tail risk.
    for row in top_k:
        o = opt_lookup.get(row["ticker"])
        if o is None:
            continue

        strike = np.float32(o.strike)
        is_call = bool(o.is_call)
        iv0 = np.float32(o.iv0)

        dte0 = int((o.expiry - today).days)
        tau0 = np.float32(max(0.0, dte0 / 365.0))

        # entry
        entry_base = np.float32(o.entry_mkt if o.entry_mkt is not None else float(
            bs_price_vec_1d(
                S=np.array([S0], dtype=np.float32),
                K=np.array([strike], dtype=np.float32),
                tau=np.array([tau0], dtype=np.float32),
                r=r,
                q=q,
                sigma=np.array([[iv0]], dtype=np.float32),
                is_call=np.array([is_call], dtype=bool),
            )[0, 0]
        ))
        entry_fill = entry_base * (1.0 + half_spread_frac)

        # exit state per path (N,)
        active = np.ones((n_paths,), dtype=bool)
        exit_day = np.full((n_paths,), T, dtype=np.int16)
        exit_price = np.zeros((n_paths,), dtype=np.float32)
        hit_tp = np.zeros((n_paths,), dtype=bool)
        hit_sl = np.zeros((n_paths,), dtype=bool)

        # init S/x
        logS = np.full((n_paths,), math.log(S0), dtype=np.float32)
        x = np.zeros((n_paths,), dtype=np.float32)

        tp_level = entry_fill * (1.0 + take_profit_pct)
        sl_level = entry_fill * (1.0 - stop_loss_pct)

        sqrt_dt = math.sqrt(dt)

        # step 0
        S_t = np.exp(logS).astype(np.float32)
        log_m = np.log(np.maximum(S_t, 1e-12) / max(S0, 1e-12)).astype(np.float32)
        sigma = iv0 * np.exp(x + (BETA_SKEW * log_m)).astype(np.float32)
        sigma = np.clip(sigma, 1e-4, 5.0).astype(np.float32)

        tau_t = np.float32(max(0.0, (dte0 - 0) / 365.0))
        px = bs_price_vec_1d(
            S=S_t,
            K=np.array([strike], dtype=np.float32),
            tau=np.array([tau_t], dtype=np.float32),
            r=r,
            q=q,
            sigma=sigma[:, None],
            is_call=np.array([is_call], dtype=bool),
        )[:, 0]
        px = american_stub(px)
        px = (px * (1.0 - half_spread_frac)).astype(np.float32)

        tp_hit = (px >= tp_level) & active
        sl_hit = (px <= sl_level) & active
        hit = tp_hit | sl_hit
        if hit.any():
            exit_day[hit] = 0
            exit_price[hit] = px[hit]
            hit_tp[tp_hit] = True
            hit_sl[sl_hit & (~tp_hit)] = True
            active[hit] = False

        # steps 1..T
        for t in range(1, T + 1):
            if not active.any():
                break
            e_s = eps_s[:, t - 1]
            e_v = eps_v[:, t - 1]

            logS = logS + (mu_rn - 0.5 * sig_s * sig_s) * dt + sig_s * sqrt_dt * e_s
            x = x + (-KAPPA * x) * dt + VOLVOL * sqrt_dt * e_v

            S_t = np.exp(logS).astype(np.float32)
            log_m = np.log(np.maximum(S_t, 1e-12) / max(S0, 1e-12)).astype(np.float32)
            sigma = iv0 * np.exp(x + (BETA_SKEW * log_m)).astype(np.float32)
            sigma = np.clip(sigma, 1e-4, 5.0).astype(np.float32)

            tau_t = np.float32(max(0.0, (dte0 - t) / 365.0))
            px = bs_price_vec_1d(
                S=S_t,
                K=np.array([strike], dtype=np.float32),
                tau=np.array([tau_t], dtype=np.float32),
                r=r,
                q=q,
                sigma=sigma[:, None],
                is_call=np.array([is_call], dtype=bool),
            )[:, 0]
            px = american_stub(px)
            px = (px * (1.0 - half_spread_frac)).astype(np.float32)

            tp_hit = (px >= tp_level) & active
            sl_hit = (px <= sl_level) & active
            hit = tp_hit | sl_hit
            if hit.any():
                exit_day[hit] = t
                exit_price[hit] = px[hit]
                hit_tp[tp_hit] = True
                hit_sl[sl_hit & (~tp_hit)] = True
                active[hit] = False

        pnl = (exit_price - entry_fill) * 100.0  # (N,)

        # deep stats
        ev = float(np.mean(pnl))
        pop = float(np.mean(pnl > 0.0))
        quants = {f"q{int(qv*100):02d}": float(np.quantile(pnl, qv)) for qv in DEEP_Q}
        var95 = float(np.quantile(pnl, DEEP_VAR_Q))
        cvar95 = float(np.mean(pnl[pnl <= var95])) if np.any(pnl <= var95) else var95

        out_row = {
            "phase": "B",
            "horizon": horizon_name,
            "ticker": o.ticker,
            "type": "CALL" if is_call else "PUT",
            "strike": float(o.strike),
            "expiry": o.expiry.isoformat(),
            "dte": dte0,
            "iv0": float(o.iv0),
            "entry_price_used": float(entry_fill),
            "entry_source": "day_close" if o.entry_mkt is not None else "model",
            "max_loss_$": float(entry_fill * 100.0),
            "EV_$": ev,
            "POP": pop,
            "VaR95_$": var95,
            "CVaR95_$": cvar95,
            "P_hit_TP": float(np.mean(hit_tp)),
            "P_hit_SL": float(np.mean(hit_sl)),
            "E_exit_price": float(np.mean(exit_price)),
            "E_hold_days": float(np.mean(exit_day.astype(np.float32))),
            **quants,
        }
        f_out.write(json.dumps(out_row) + "\n")

    f_out.close()


# ============================
# Utilities
# ============================
def print_top_table(rows: List[dict], horizon_name: str, title: str):
    if not rows:
        print(f"\n[{horizon_name}] No rows to print.")
        return
    rows = [r for r in rows if r["horizon"] == horizon_name and r["max_loss_$"] <= MAX_PREMIUM]
    rows.sort(key=lambda x: (x["EV_$"], x["POP"]), reverse=True)
    rows = rows[:TOP_N_PRINT]

    print(f"\n=== {title} | {UNDERLYING} | Horizon={horizon_name} | Top {len(rows)} ===")
    print(f"Policy: +{TAKE_PROFIT_PCT*100:.0f}% TP / -{STOP_LOSS_PCT*100:.0f}% SL | Slippage half-spread {HALF_SPREAD_FRAC*100:.1f}%")
    print("Rank | Type | Strike | DTE | IV  | Entry(src) | MaxLoss$ | EV$  | POP  | P(TP) | P(SL) | EHold | Ticker")
    for i, r in enumerate(rows, 1):
        print(
            f"{i:4d} | {r['type']:4s} | {r['strike']:6.2f} | {r['dte']:3d} | {r['iv0']:.3f} | "
            f"{r['entry_price_used']:.2f}({r['entry_source'][:3]}) | {r['max_loss_$']:8.0f} | "
            f"{r['EV_$']:5.1f} | {r['POP']:.3f} | {r.get('P_hit_TP',0):.3f} | {r.get('P_hit_SL',0):.3f} | "
            f"{r.get('E_hold_days',0):.1f} | {r['ticker']}"
        )


def read_jsonl_gz(path: str) -> List[dict]:
    rows: List[dict] = []
    with gzip.open(path, "rt", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                rows.append(json.loads(line))
    return rows


# ============================
# Main
# ============================
def main():
    if not API_KEY:
        raise RuntimeError("Set MASSIVE_API_KEY in your environment.")

    today = effective_market_date()
    api = MassiveApiClient(API_KEY)

    prev = api.fetch_previous_close(UNDERLYING)
    S0 = prev.get("close")
    if not isinstance(S0, (int, float)) or S0 <= 0:
        raise RuntimeError(f"Could not fetch {UNDERLYING} previous close via /v2/aggs/ticker/{UNDERLYING}/prev")
    S0 = float(S0)

    raw = api.fetch_option_snapshots(UNDERLYING, limit=250)
    opts = parse_opt_records(raw, today=today)
    if not opts:
        raise RuntimeError("No options after parsing/filters. Try loosening MIN_DTE/MAX_DTE/MIN_IV.")

    # ATM-ish IV estimate for stock diffusion:
    strikes = np.array([o.strike for o in opts], dtype=np.float32)
    ivs = np.array([o.iv0 for o in opts], dtype=np.float32)
    idx_atm = np.argsort(np.abs(strikes - S0))[:min(50, len(opts))]
    atm_iv0 = float(np.median(ivs[idx_atm])) if len(idx_atm) else float(np.median(ivs))
    atm_iv0 = max(0.10, min(2.50, atm_iv0))

    print(f"\nUnderlying: {UNDERLYING}")
    print(f"Market date: {today.isoformat()} | S0(prev close): {S0:.2f} | ATM_IV0~{atm_iv0:.3f}")
    print(f"Options considered: {len(opts)} (DTE {MIN_DTE}-{MAX_DTE}, IV>={MIN_IV})")
    print(f"Phase A (scan): {N_PATHS_SCAN} paths | batch={BATCH_SIZE} | horizons={[h['name'] for h in HORIZONS]}")
    print(f"Phase B (deep): {N_PATHS_DEEP} paths | topK={TOP_K_DEEP} per horizon")
    print("Pricing: European BS (vectorized). American: stub hook.")

    # lookup for phase B
    opt_lookup = {o.ticker: o for o in opts}

    # Run per horizon
    for h in HORIZONS:
        name = h["name"]
        T = h["max_hold_days"]

        out_a = f"{UNDERLYING}_{name}_phaseA.jsonl.gz"
        out_b = f"{UNDERLYING}_{name}_phaseB.jsonl.gz"

        print(f"\n--- Horizon {name} (max_hold_days={T}) ---")
        print(f"Writing Phase A to {out_a}")

        top_k = scan_horizon_phase_a(
            opts=opts,
            S0=S0,
            today=today,
            horizon_name=name,
            max_hold_days=T,
            atm_iv0=atm_iv0,
            n_paths=N_PATHS_SCAN,
            r=R,
            q=Q,
            take_profit_pct=TAKE_PROFIT_PCT,
            stop_loss_pct=STOP_LOSS_PCT,
            half_spread_frac=HALF_SPREAD_FRAC,
            batch_size=BATCH_SIZE,
            out_jsonl_gz_path=out_a,
        )

        # show top from Phase A (light stats)
        phase_a_rows = read_jsonl_gz(out_a)
        print_top_table(phase_a_rows, name, "Phase A (light)")

        print(f"Running Phase B deep risk for topK={len(top_k)} -> {out_b}")
        deep_risk_phase_b(
            top_k=top_k,
            opt_lookup=opt_lookup,
            S0=S0,
            today=today,
            horizon_name=name,
            max_hold_days=T,
            atm_iv0=atm_iv0,
            n_paths=N_PATHS_DEEP,
            r=R,
            q=Q,
            take_profit_pct=TAKE_PROFIT_PCT,
            stop_loss_pct=STOP_LOSS_PCT,
            half_spread_frac=HALF_SPREAD_FRAC,
            out_jsonl_gz_path=out_b,
        )

        # read Phase B and print top
        phase_b_rows = read_jsonl_gz(out_b)
        # augment print with VaR/CVaR if present
        if phase_b_rows:
            phase_b_rows.sort(key=lambda x: (x["EV_$"], x["POP"]), reverse=True)
            top = [r for r in phase_b_rows if r["max_loss_$"] <= MAX_PREMIUM][:TOP_N_PRINT]
            print(f"\n=== Phase B (deep) | {UNDERLYING} | Horizon={name} | Top {len(top)} ===")
            print("Rank | Type | Strike | DTE | IV  | Entry(src) | MaxLoss$ | EV$  | POP  | VaR95$ | CVaR95$ | EHold | Ticker")
            for i, r in enumerate(top, 1):
                print(
                    f"{i:4d} | {r['type']:4s} | {r['strike']:6.2f} | {r['dte']:3d} | {r['iv0']:.3f} | "
                    f"{r['entry_price_used']:.2f}({r['entry_source'][:3]}) | {r['max_loss_$']:8.0f} | "
                    f"{r['EV_$']:5.1f} | {r['POP']:.3f} | {r['VaR95_$']:6.1f} | {r['CVaR95_$']:7.1f} | "
                    f"{r['E_hold_days']:.1f} | {r['ticker']}"
                )
        else:
            print("[WARN] No Phase B rows produced (maybe lookup mismatch or filtered out).")

    print("\nDone. Outputs are gzipped JSONL files per horizon:")
    print("  *_phaseA.jsonl.gz  (all options, light stats)")
    print("  *_phaseB.jsonl.gz  (topK options, deep tail risk stats)")


if __name__ == "__main__":
    main()

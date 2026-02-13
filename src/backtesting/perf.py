"""Performance profiling helpers for reference vs optimized execution paths."""

from __future__ import annotations

import cProfile
import json
import pstats
from dataclasses import dataclass
from io import StringIO
from pathlib import Path
from tempfile import TemporaryDirectory
from time import perf_counter
from statistics import NormalDist

import numpy as np

from .execution import BpsSlippage, FixedCommission, ShortBorrowCost
from .vectorized import backtest_vectorized


@dataclass(frozen=True)
class ProfileSummary:
    mode: str
    output: str
    elapsed_seconds: float


DEFAULT_PROFILE_THRESHOLDS: dict[str, float] = {
    "optimized_to_reference_max_ratio": 1.20,
}


def deflated_sharpe_ratio(
    *,
    observed_sharpe: float,
    n_returns: int,
    skew: float,
    kurtosis: float,
    n_trials: int,
) -> float:
    """Compute Deflated Sharpe Ratio probability after multiple testing correction."""
    if n_returns <= 2:
        return 0.0

    sr = float(observed_sharpe)
    variance_term = 1.0 - skew * sr + ((kurtosis - 1.0) / 4.0) * (sr**2)
    sr_std = np.sqrt(max(variance_term, 1e-12) / max(1, n_returns - 1))

    effective_trials = max(1, int(n_trials))
    if effective_trials == 1:
        sr_star = 0.0
    else:
        norm = NormalDist()
        q1 = norm.inv_cdf(1.0 - 1.0 / effective_trials)
        q2 = norm.inv_cdf(1.0 - 1.0 / (effective_trials * np.e))
        euler_gamma = 0.5772156649
        sr_star = (1.0 - euler_gamma) * q1 + euler_gamma * q2
    return float(NormalDist().cdf((sr - sr_star) / max(sr_std, 1e-12)))


def probabilistic_sharpe_ratio(
    *,
    observed_sharpe: float,
    benchmark_sharpe: float,
    n_returns: int,
    skew: float = 0.0,
    kurtosis: float = 3.0,
) -> float:
    """Return Probabilistic Sharpe Ratio: P(SR > benchmark_sharpe)."""
    if n_returns <= 2:
        return 0.0
    sr = float(observed_sharpe)
    sr_ref = float(benchmark_sharpe)
    variance_term = 1.0 - skew * sr + ((kurtosis - 1.0) / 4.0) * (sr**2)
    sr_std = np.sqrt(max(variance_term, 1e-12) / max(1, n_returns - 1))
    z_score = (sr - sr_ref) / max(sr_std, 1e-12)
    return float(NormalDist().cdf(z_score))


def benjamini_hochberg_adjusted_pvalues(p_values: list[float]) -> list[float]:
    """Compute Benjamini-Hochberg adjusted p-values (q-values)."""
    m = len(p_values)
    if m == 0:
        return []

    clipped = np.clip(np.asarray(p_values, dtype=float), 0.0, 1.0)
    order = np.argsort(clipped)
    ranked = clipped[order]

    adjusted = np.empty(m, dtype=float)
    running = 1.0
    for idx in range(m - 1, -1, -1):
        rank = idx + 1
        candidate = (ranked[idx] * m) / rank
        running = min(running, float(candidate))
        adjusted[idx] = running

    out = np.empty(m, dtype=float)
    out[order] = np.clip(adjusted, 0.0, 1.0)
    return out.tolist()


def profile_backtest_hotspots(n_periods: int = 20_000, n_assets: int = 128) -> list[ProfileSummary]:
    rng = np.random.default_rng(42)
    returns = rng.normal(loc=0.0002, scale=0.01, size=(n_periods, n_assets))
    prices = 100.0 * np.cumprod(1.0 + returns, axis=0)
    signals = np.sign(rng.normal(size=(n_periods, n_assets)))
    weights = np.full(n_assets, 1.0 / n_assets)

    outputs: list[ProfileSummary] = []
    for mode in ("reference", "optimized"):
        pr = cProfile.Profile()
        start = perf_counter()
        pr.enable()
        backtest_vectorized(
            prices,
            signals,
            slippage_model=BpsSlippage(5.0),
            fee_model=FixedCommission(0.0005),
            borrow_cost_model=ShortBorrowCost(0.03),
            weights=weights,
            execution_mode=mode,
        )
        pr.disable()
        elapsed = perf_counter() - start

        buf = StringIO()
        stats = pstats.Stats(pr, stream=buf).sort_stats("cumtime")
        stats.print_stats(15)
        outputs.append(ProfileSummary(mode=mode, output=buf.getvalue(), elapsed_seconds=float(elapsed)))
    return outputs




def benchmark_serialization_boundaries(
    *,
    n_periods: int = 20_000,
    n_assets: int = 32,
) -> dict[str, dict[str, float] | int]:
    """Benchmark serialization and I/O boundaries for representative arrays."""

    rng = np.random.default_rng(123)
    payload = {
        "open": (100 + rng.normal(size=(n_periods, n_assets))).astype(np.float64),
        "close": (100 + rng.normal(size=(n_periods, n_assets))).astype(np.float64),
        "signal": np.sign(rng.normal(size=(n_periods, n_assets))).astype(np.float64),
    }

    results: dict[str, dict[str, float] | int] = {}

    with TemporaryDirectory() as tmp:
        root = Path(tmp)

        npz_path = root / "bundle.npz"
        start = perf_counter()
        np.savez_compressed(npz_path, **payload)
        npz_write = perf_counter() - start
        start = perf_counter()
        with np.load(npz_path) as loaded:
            _ = {k: loaded[k] for k in loaded.files}
        npz_read = perf_counter() - start
        results["npz"] = {
            "write_seconds": float(npz_write),
            "read_seconds": float(npz_read),
            "size_bytes": float(npz_path.stat().st_size),
        }

        json_path = root / "bundle.json"
        start = perf_counter()
        json_path.write_text(json.dumps({k: v.tolist() for k, v in payload.items()}))
        json_write = perf_counter() - start
        start = perf_counter()
        _ = json.loads(json_path.read_text())
        json_read = perf_counter() - start
        results["json"] = {
            "write_seconds": float(json_write),
            "read_seconds": float(json_read),
            "size_bytes": float(json_path.stat().st_size),
        }

        try:
            import pyarrow as pa
            import pyarrow.parquet as pq

            parquet_path = root / "bundle.parquet"
            flat_data = {
                name: values.reshape(-1)
                for name, values in payload.items()
            }
            table = pa.table(flat_data)
            start = perf_counter()
            pq.write_table(table, parquet_path)
            parquet_write = perf_counter() - start
            start = perf_counter()
            _ = pq.read_table(parquet_path)
            parquet_read = perf_counter() - start
            results["parquet"] = {
                "write_seconds": float(parquet_write),
                "read_seconds": float(parquet_read),
                "size_bytes": float(parquet_path.stat().st_size),
            }
        except Exception:
            results["parquet"] = {
                "write_seconds": -1.0,
                "read_seconds": -1.0,
                "size_bytes": -1.0,
            }

    return {
        "n_periods": int(n_periods),
        "n_assets": int(n_assets),
        "results": results,
    }


def check_profile_regression(
    summaries: list[ProfileSummary],
    *,
    thresholds: dict[str, float] | None = None,
) -> dict[str, float | bool]:
    """Validate profiled mode runtimes against configurable performance thresholds."""
    limits = dict(DEFAULT_PROFILE_THRESHOLDS)
    if thresholds:
        limits.update({k: float(v) for k, v in thresholds.items()})

    by_mode = {item.mode: item for item in summaries}
    if "reference" not in by_mode or "optimized" not in by_mode:
        raise ValueError("summaries must include both 'reference' and 'optimized' modes")

    ref = max(by_mode["reference"].elapsed_seconds, 1e-12)
    opt = by_mode["optimized"].elapsed_seconds
    ratio = float(opt / ref)
    max_ratio = float(limits["optimized_to_reference_max_ratio"])
    return {
        "optimized_to_reference_ratio": ratio,
        "optimized_to_reference_max_ratio": max_ratio,
        "pass": ratio <= max_ratio,
    }

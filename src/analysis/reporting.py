from __future__ import annotations

from datetime import datetime
from statistics import NormalDist
from typing import Iterable

import numpy as np

from backtesting.perf import benjamini_hochberg_adjusted_pvalues, deflated_sharpe_ratio, probabilistic_sharpe_ratio

from .cross_sectional.base import CrossSectionalResult
from .time_series.base import TimeSeriesResult


def format_cross_sectional_report(
    *,
    title: str,
    as_of: str,
    universe: Iterable[str],
    settings: dict[str, object],
    result: CrossSectionalResult,
) -> str:
    lines: list[str] = []
    universe_list = list(universe)
    lines.append(title)
    lines.append("=" * len(title))
    lines.append(f"Generated: {datetime.now().isoformat(timespec='seconds')}")
    lines.append(f"As of: {as_of}")
    lines.append(f"Universe size: {len(universe_list)}")
    lines.append("")
    lines.append("Parameters:")
    for key, value in settings.items():
        lines.append(f"  - {key}: {value}")
    lines.append("")

    max_ticker_len = _max_ticker_len(universe_list, result)

    lines.append("Top Winners:")
    if result.longs:
        for ticker in sorted(result.longs, key=lambda t: result.scores.get(t, 0.0), reverse=True):
            lines.append(_format_ticker_line(ticker, result, max_ticker_len))
    else:
        lines.append("  (none)")
    lines.append("")

    lines.append("Top Losers:")
    if result.shorts:
        for ticker in sorted(result.shorts, key=lambda t: result.scores.get(t, 0.0)):
            lines.append(_format_ticker_line(ticker, result, max_ticker_len))
    else:
        lines.append("  (none)")
    lines.append("")

    lines.append("Ranking (best to worst):")
    if result.ranking:
        for ticker, score in result.ranking:
            lines.append(_format_ticker_line(ticker, result, max_ticker_len, score_override=score))
    else:
        lines.append("  (none)")

    lines.append("")
    lines.append("All Surveyed (best to worst, including skipped):")
    scored = {**{ticker: score for ticker, score in result.ranking}}
    skipped = dict(result.skipped)
    for ticker in sorted(universe_list):
        if ticker not in scored and ticker not in skipped:
            skipped[ticker] = "no_data"
    surveyed_scores = [
        (ticker, scored.get(ticker)) for ticker in universe_list if ticker in scored
    ]
    surveyed_scores.sort(key=lambda item: item[1], reverse=True)
    for ticker, score in surveyed_scores:
        lines.append(_format_ticker_line(ticker, result, max_ticker_len, score_override=score))
    for ticker, reason in sorted(skipped.items()):
        if ticker not in scored:
            lines.append(f"  - {ticker}: skipped ({reason})")

    return "\n".join(lines)


def format_time_series_report(
    *,
    title: str,
    as_of: str,
    universe: Iterable[str],
    settings: dict[str, object],
    result: TimeSeriesResult,
) -> str:
    lines: list[str] = []
    universe_list = list(universe)
    lines.append(title)
    lines.append("=" * len(title))
    lines.append(f"Generated: {datetime.now().isoformat(timespec='seconds')}")
    lines.append(f"As of: {as_of}")
    lines.append(f"Universe size: {len(universe_list)}")
    lines.append("")
    lines.append("Parameters:")
    for key, value in settings.items():
        lines.append(f"  - {key}: {value}")
    lines.append("")

    max_ticker_len = _max_ticker_len_time_series(universe_list, result)

    lines.append("Top Winners:")
    if result.longs:
        for ticker in sorted(result.longs, key=lambda t: result.scores.get(t, 0.0), reverse=True):
            lines.append(_format_ticker_line_time_series(ticker, result, max_ticker_len))
    else:
        lines.append("  (none)")
    lines.append("")

    lines.append("Top Losers:")
    if result.shorts:
        for ticker in sorted(result.shorts, key=lambda t: result.scores.get(t, 0.0)):
            lines.append(_format_ticker_line_time_series(ticker, result, max_ticker_len))
    else:
        lines.append("  (none)")
    lines.append("")

    lines.append("Ranking (best to worst):")
    if result.ranking:
        for ticker, score in result.ranking:
            lines.append(
                _format_ticker_line_time_series(
                    ticker, result, max_ticker_len, score_override=score
                )
            )
    else:
        lines.append("  (none)")

    lines.append("")
    lines.append("All Surveyed (best to worst, including skipped):")
    scored = {**{ticker: score for ticker, score in result.ranking}}
    skipped = dict(result.skipped)
    for ticker in sorted(universe_list):
        if ticker not in scored and ticker not in skipped:
            skipped[ticker] = "no_data"
    surveyed_scores = [
        (ticker, scored.get(ticker)) for ticker in universe_list if ticker in scored
    ]
    surveyed_scores.sort(key=lambda item: item[1], reverse=True)
    for ticker, score in surveyed_scores:
        lines.append(
            _format_ticker_line_time_series(
                ticker, result, max_ticker_len, score_override=score
            )
        )
    for ticker, reason in sorted(skipped.items()):
        if ticker not in scored:
            lines.append(f"  - {ticker}: skipped ({reason})")

    return "\n".join(lines)


def format_backtest_report(
    *,
    title: str,
    params: dict[str, object],
    metrics: dict[str, float],
    drawdown_rows: list[dict[str, object]],
    turnover_stats: dict[str, float],
    cost_totals: dict[str, float],
    robustness_report: dict[str, object] | None = None,
) -> str:
    lines: list[str] = []
    lines.append(title)
    lines.append("=" * len(title))
    lines.append(f"Generated: {datetime.now().isoformat(timespec='seconds')}")
    lines.append("")

    lines.append("Parameters:")
    for key, value in params.items():
        lines.append(f"  - {key}: {value}")
    lines.append("")

    lines.append("Summary Metrics:")
    for key, value in metrics.items():
        lines.append(f"  - {key}: {value:.6f}")
    lines.append("")

    if any(key in metrics for key in ("rolling_sharpe_mean", "rolling_sharpe_min", "rolling_sharpe_max")):
        lines.append("Rolling Sharpe Summary:")
        lines.append(
            "  - mean={mean:.6f}, min={minv:.6f}, max={maxv:.6f}, window={window:.0f}".format(
                mean=metrics.get("rolling_sharpe_mean", 0.0),
                minv=metrics.get("rolling_sharpe_min", 0.0),
                maxv=metrics.get("rolling_sharpe_max", 0.0),
                window=metrics.get("rolling_window", 0.0),
            )
        )
        lines.append("")

    if any(key in metrics for key in ("rolling_drawdown_mean", "rolling_drawdown_worst")):
        lines.append("Rolling Drawdown Summary:")
        lines.append(
            "  - mean={mean:.6f}, worst={worst:.6f}, window={window:.0f}".format(
                mean=metrics.get("rolling_drawdown_mean", 0.0),
                worst=metrics.get("rolling_drawdown_worst", 0.0),
                window=metrics.get("rolling_window", 0.0),
            )
        )
        lines.append("")

    lines.append("Drawdown Table:")
    if drawdown_rows:
        for row in drawdown_rows:
            lines.append(
                "  - {timestamp}: drawdown={drawdown:.6f}, equity={equity:.6f}, "
                "running_peak={running_peak:.6f}".format(
                    timestamp=row["timestamp"],
                    drawdown=float(row["drawdown"]),
                    equity=float(row["equity"]),
                    running_peak=float(row["running_peak"]),
                )
            )
    else:
        lines.append("  (none)")
    lines.append("")

    lines.append("Turnover and Cost Attribution:")
    lines.append(
        "  - turnover_mean={mean:.6f}, turnover_total={total:.6f}, turnover_max={max:.6f}".format(
            mean=turnover_stats.get("mean", 0.0),
            total=turnover_stats.get("total", 0.0),
            max=turnover_stats.get("max", 0.0),
        )
    )
    lines.append(
        "  - costs_total={total:.6f}, slippage={slippage:.6f}, fees={fees:.6f}, borrow={borrow:.6f}".format(
            total=cost_totals.get("total", 0.0),
            slippage=cost_totals.get("slippage", 0.0),
            fees=cost_totals.get("fees", 0.0),
            borrow=cost_totals.get("borrow", 0.0),
        )
    )

    if robustness_report:
        ci = robustness_report.get("bootstrap_confidence_intervals", {})
        if isinstance(ci, dict) and ci:
            lines.append("Bootstrap Confidence Intervals (95%):")
            for metric_name in ("sharpe", "cagr", "max_drawdown"):
                bucket = ci.get(metric_name, {}) if isinstance(ci, dict) else {}
                if isinstance(bucket, dict):
                    lines.append(
                        "  - {metric}: lower={lower:.6f}, median={median:.6f}, upper={upper:.6f}".format(
                            metric=metric_name,
                            lower=float(bucket.get("lower", 0.0)),
                            median=float(bucket.get("median", 0.0)),
                            upper=float(bucket.get("upper", 0.0)),
                        )
                    )
            lines.append("")

        dsr = float(robustness_report.get("deflated_sharpe_ratio", 0.0))
        lines.append(f"Deflated Sharpe Ratio: {dsr:.6f}")

        capacity = robustness_report.get("capacity_diagnostics", {})
        if isinstance(capacity, dict) and capacity:
            lines.append("Capacity Diagnostics:")
            lines.append(
                "  - average_participation={part:.6f}, realized_slippage_bps={slip:.6f}".format(
                    part=float(capacity.get("average_participation_rate", 0.0)),
                    slip=float(capacity.get("realized_slippage_bps", 0.0)),
                )
            )
            curve = capacity.get("expected_slippage_curve", [])
            if isinstance(curve, list):
                for row in curve[:4]:
                    if isinstance(row, dict):
                        lines.append(
                            "  - expected_slippage @ participation={p:.2%}: {bps:.3f} bps".format(
                                p=float(row.get("participation_rate", 0.0)),
                                bps=float(row.get("expected_slippage_bps", 0.0)),
                            )
                        )
            frontier = capacity.get("capacity_frontier", [])
            if isinstance(frontier, list) and frontier:
                lines.append("  - Capacity frontier (alpha net costs):")
                for row in frontier[:4]:
                    if isinstance(row, dict):
                        lines.append(
                            "    - AUM={aum:.0f}, scale={scale:.2f}x, net_alpha={net:.3f} bps".format(
                                aum=float(row.get("aum", 0.0)),
                                scale=float(row.get("aum_scale", 0.0)),
                                net=float(row.get("expected_alpha_net_cost_bps", 0.0)),
                            )
                        )
            lines.append("")

        drift = robustness_report.get("model_drift_diagnostics", {})
        if isinstance(drift, dict) and drift:
            lines.append("Model Drift Diagnostics:")
            lines.append(
                "  - baseline_mean={base_mean:.6f}, baseline_vol={base_vol:.6f}, "
                "current_mean={curr_mean:.6f}, current_vol={curr_vol:.6f}".format(
                    base_mean=float(drift.get("baseline_mean", 0.0)),
                    base_vol=float(drift.get("baseline_vol", 0.0)),
                    curr_mean=float(drift.get("current_mean", 0.0)),
                    curr_vol=float(drift.get("current_vol", 0.0)),
                )
            )
            lines.append(
                "  - drift_z_score={z:.6f}, retraining_triggered={trigger}".format(
                    z=float(drift.get("drift_z_score", 0.0)),
                    trigger=bool(drift.get("retraining_triggered", False)),
                )
            )
            lines.append("")

    return "\n".join(lines)


def build_drawdown_rows(
    timestamps: np.ndarray,
    equity_curve: np.ndarray,
    *,
    top_n: int = 10,
) -> list[dict[str, object]]:
    if equity_curve.size == 0:
        return []
    running_peak = np.maximum.accumulate(equity_curve)
    safe_peak = np.where(running_peak == 0.0, 1.0, running_peak)
    drawdown = equity_curve / safe_peak - 1.0
    order = np.argsort(drawdown)
    selected = order[: min(top_n, drawdown.size)]
    rows: list[dict[str, object]] = []
    for idx in selected:
        ts = datetime.utcfromtimestamp(int(timestamps[idx]) / 1000.0).isoformat()
        rows.append(
            {
                "timestamp": ts,
                "drawdown": float(drawdown[idx]),
                "equity": float(equity_curve[idx]),
                "running_peak": float(running_peak[idx]),
            }
        )
    return rows


def compute_bootstrap_confidence_intervals(
    *,
    returns: np.ndarray,
    periods_per_year: float,
    n_bootstrap: int = 500,
    confidence: float = 0.95,
    seed: int = 42,
) -> dict[str, dict[str, float]]:
    samples = _to_1d_float(returns)
    if samples.size == 0 or n_bootstrap <= 1:
        return {
            "sharpe": {"lower": 0.0, "median": 0.0, "upper": 0.0},
            "cagr": {"lower": 0.0, "median": 0.0, "upper": 0.0},
            "max_drawdown": {"lower": 0.0, "median": 0.0, "upper": 0.0},
        }

    ann_factor = max(float(periods_per_year), 1.0)
    rng = np.random.default_rng(seed)
    sharpe_samples: list[float] = []
    cagr_samples: list[float] = []
    drawdown_samples: list[float] = []

    for _ in range(int(n_bootstrap)):
        boot = samples[rng.integers(0, samples.size, size=samples.size)]
        mean = float(np.mean(boot))
        vol = float(np.std(boot, ddof=1)) if boot.size > 1 else 0.0
        sharpe_samples.append(mean / vol * np.sqrt(ann_factor) if vol else 0.0)

        equity = np.cumprod(1.0 + boot)
        start = float(equity[0]) if equity.size else 1.0
        end = float(equity[-1]) if equity.size else 1.0
        cagr_samples.append((end / start) ** (ann_factor / max(1, boot.size)) - 1.0 if start else 0.0)

        peak = np.maximum.accumulate(equity)
        safe_peak = np.where(peak == 0.0, 1.0, peak)
        drawdown = equity / safe_peak - 1.0
        drawdown_samples.append(float(np.min(drawdown)) if drawdown.size else 0.0)

    alpha = (1.0 - float(confidence)) / 2.0
    lo_q = 100.0 * max(0.0, alpha)
    hi_q = 100.0 * min(1.0, 1.0 - alpha)

    return {
        "sharpe": _summarize_distribution(sharpe_samples, lo_q, hi_q),
        "cagr": _summarize_distribution(cagr_samples, lo_q, hi_q),
        "max_drawdown": _summarize_distribution(drawdown_samples, lo_q, hi_q),
    }


def compute_deflated_sharpe_ratio(
    *,
    observed_sharpe: float,
    n_returns: int,
    skew: float,
    kurtosis: float,
    n_trials: int,
) -> float:
    return deflated_sharpe_ratio(
        observed_sharpe=observed_sharpe,
        n_returns=n_returns,
        skew=skew,
        kurtosis=kurtosis,
        n_trials=n_trials,
    )


def build_backtest_robustness_report(
    *,
    returns: np.ndarray,
    metrics: dict[str, float],
    turnover_stats: dict[str, float],
    cost_totals: dict[str, float],
    fills: list[dict[str, object]],
    n_bootstrap: int = 500,
    seed: int = 42,
) -> dict[str, object]:
    periods_per_year = float(metrics.get("periods_per_year", 252.0))
    ci = compute_bootstrap_confidence_intervals(
        returns=_to_1d_float(returns),
        periods_per_year=periods_per_year,
        n_bootstrap=n_bootstrap,
        confidence=0.95,
        seed=seed,
    )

    dsr = compute_deflated_sharpe_ratio(
        observed_sharpe=float(metrics.get("sharpe", 0.0)),
        n_returns=int(_to_1d_float(returns).size),
        skew=float(metrics.get("skew", 0.0)),
        kurtosis=float(metrics.get("kurtosis", 0.0) + 3.0),
        n_trials=1,
    )

    capacity = _build_capacity_diagnostics(
        metrics=metrics,
        turnover_stats=turnover_stats,
        cost_totals=cost_totals,
        fills=fills,
    )
    drift = compute_model_drift_diagnostics(returns=_to_1d_float(returns))
    return {
        "bootstrap_confidence_intervals": ci,
        "deflated_sharpe_ratio": dsr,
        "capacity_diagnostics": capacity,
        "model_drift_diagnostics": drift,
    }






def compute_model_drift_diagnostics(
    *,
    returns: np.ndarray,
    baseline_mean: float | None = None,
    baseline_vol: float | None = None,
    drift_z_threshold: float = 2.0,
) -> dict[str, object]:
    samples = _to_1d_float(returns)
    if samples.size == 0:
        return {
            "baseline_mean": 0.0,
            "baseline_vol": 0.0,
            "current_mean": 0.0,
            "current_vol": 0.0,
            "drift_z_score": 0.0,
            "retraining_triggered": False,
        }

    split = max(1, samples.size // 2)
    base = samples[:split]
    current = samples[split:] if split < samples.size else samples

    base_mean = float(np.mean(base)) if baseline_mean is None else float(baseline_mean)
    base_vol = float(np.std(base, ddof=1)) if base.size > 1 else 0.0
    if baseline_vol is not None:
        base_vol = float(baseline_vol)

    current_mean = float(np.mean(current)) if current.size else 0.0
    current_vol = float(np.std(current, ddof=1)) if current.size > 1 else 0.0

    denom = max(base_vol, 1e-12)
    drift_z = (current_mean - base_mean) / denom
    retrain = bool(abs(drift_z) >= float(drift_z_threshold))

    return {
        "baseline_mean": base_mean,
        "baseline_vol": base_vol,
        "current_mean": current_mean,
        "current_vol": current_vol,
        "drift_z_score": float(drift_z),
        "retraining_triggered": retrain,
    }
def compute_white_reality_check(
    *,
    candidate_returns: np.ndarray,
    n_bootstrap: int = 500,
    seed: int = 42,
) -> dict[str, float]:
    arr = np.asarray(candidate_returns, dtype=float)
    if arr.ndim != 2 or arr.size == 0:
        return {"observed_max_mean": 0.0, "p_value": 1.0, "n_candidates": 0.0, "n_observations": 0.0}
    n_obs, n_candidates = arr.shape
    centered = arr - np.mean(arr, axis=0, keepdims=True)
    observed = float(np.max(np.mean(arr, axis=0)))
    rng = np.random.default_rng(seed)
    boot_stats: list[float] = []
    for _ in range(max(1, int(n_bootstrap))):
        idx = rng.integers(0, n_obs, size=n_obs)
        sample = centered[idx, :]
        boot_stats.append(float(np.max(np.mean(sample, axis=0))))
    p_value = float(np.mean(np.asarray(boot_stats) >= observed)) if boot_stats else 1.0
    return {
        "observed_max_mean": observed,
        "p_value": p_value,
        "n_candidates": float(n_candidates),
        "n_observations": float(n_obs),
    }


def compute_spa_pvalue(
    *,
    candidate_returns: np.ndarray,
    benchmark_returns: np.ndarray | None = None,
    n_bootstrap: int = 500,
    seed: int = 42,
) -> dict[str, float]:
    arr = np.asarray(candidate_returns, dtype=float)
    if arr.ndim != 2 or arr.size == 0:
        return {"observed_stat": 0.0, "p_value": 1.0, "n_candidates": 0.0, "n_observations": 0.0}
    n_obs, n_candidates = arr.shape
    bench = np.zeros(n_obs, dtype=float) if benchmark_returns is None else np.asarray(benchmark_returns, dtype=float).reshape(-1)
    if bench.size != n_obs:
        raise ValueError("benchmark_returns must match candidate_returns length")

    excess = arr - bench[:, None]
    means = np.mean(excess, axis=0)
    stds = np.std(excess, axis=0, ddof=1)
    denom = stds / np.sqrt(max(1, n_obs))
    t_stats = np.divide(means, denom, out=np.zeros_like(means), where=denom > 1e-12)
    observed = float(np.max(np.maximum(t_stats, 0.0)))

    centered = excess - means[None, :]
    rng = np.random.default_rng(seed)
    boot_max: list[float] = []
    for _ in range(max(1, int(n_bootstrap))):
        idx = rng.integers(0, n_obs, size=n_obs)
        sample = centered[idx, :]
        s_mean = np.mean(sample, axis=0)
        s_std = np.std(sample, axis=0, ddof=1)
        s_denom = s_std / np.sqrt(max(1, n_obs))
        s_t = np.divide(s_mean, s_denom, out=np.zeros_like(s_mean), where=s_denom > 1e-12)
        boot_max.append(float(np.max(np.maximum(s_t, 0.0))))
    p_value = float(np.mean(np.asarray(boot_max) >= observed)) if boot_max else 1.0
    return {
        "observed_stat": observed,
        "p_value": p_value,
        "n_candidates": float(n_candidates),
        "n_observations": float(n_obs),
    }
def build_sweep_robustness_report(
    *,
    ranked_rows: list[dict[str, object]],
    score_key: str = "sharpe",
    n_monte_carlo: int = 200,
    seed: int = 42,
) -> dict[str, object]:
    scores = np.array([float(row.get(score_key, 0.0)) for row in ranked_rows], dtype=float)
    if scores.size == 0:
        return {
            "deflated_sharpe_ratio": 0.0,
            "probabilistic_sharpe_ratio": 0.0,
            "pbo_style": {
                "n_combinations": 0,
                "n_monte_carlo": int(n_monte_carlo),
                "probability_of_overfitting": 0.0,
                "median_logit": 0.0,
            },
            "white_reality_check": {
                "observed_max_mean": 0.0,
                "p_value": 1.0,
                "n_candidates": 0.0,
                "n_observations": 0.0,
            },
            "spa": {
                "observed_stat": 0.0,
                "p_value": 1.0,
                "n_candidates": 0.0,
                "n_observations": 0.0,
            },
            "multiple_testing": {
                "method": "benjamini_hochberg",
                "n_hypotheses": 0,
                "raw_pvalues": [],
                "bh_adjusted_pvalues": [],
                "min_raw_pvalue": 1.0,
                "min_bh_adjusted_pvalue": 1.0,
            },
        }

    centered = scores - float(np.mean(scores))
    m2 = float(np.mean(centered**2)) if scores.size > 1 else 0.0
    skew = float(np.mean(centered**3) / (m2 ** 1.5)) if scores.size > 2 and m2 > 0 else 0.0
    kurt = float(np.mean(centered**4) / (m2**2)) if scores.size > 3 and m2 > 0 else 3.0

    n_trials = int(scores.size)
    best_sharpe = float(scores[0])
    dsr = compute_deflated_sharpe_ratio(
        observed_sharpe=best_sharpe,
        n_returns=max(3, int(scores.size)),
        skew=skew,
        kurtosis=kurt,
        n_trials=n_trials,
    )
    psr = probabilistic_sharpe_ratio(
        observed_sharpe=best_sharpe,
        benchmark_sharpe=0.0,
        n_returns=max(3, int(scores.size)),
        skew=skew,
        kurtosis=kurt,
    )

    pbo = _compute_pbo_style(ranked_rows=ranked_rows, score_key=score_key, n_monte_carlo=n_monte_carlo, seed=seed)

    wr_cols = [
        key for key in sorted({k for row in ranked_rows for k in row.keys()})
        if key.startswith("ret_")
    ]
    if wr_cols:
        candidate_returns = np.asarray(
            [[float(row.get(col, 0.0)) for col in wr_cols] for row in ranked_rows],
            dtype=float,
        ).T
    else:
        candidate_returns = scores.reshape(-1, 1)
    white = compute_white_reality_check(candidate_returns=candidate_returns, n_bootstrap=max(200, n_monte_carlo), seed=seed)
    spa = compute_spa_pvalue(candidate_returns=candidate_returns, n_bootstrap=max(200, n_monte_carlo), seed=seed)

    if wr_cols:
        mean_scores = np.mean(candidate_returns, axis=0)
        std_scores = np.std(candidate_returns, axis=0, ddof=1)
        n_obs = max(1, int(candidate_returns.shape[0]))
        denom = std_scores / np.sqrt(n_obs)
        t_stats = np.divide(mean_scores, denom, out=np.zeros_like(mean_scores), where=denom > 1e-12)
        p_values = [float(1.0 - NormalDist().cdf(float(t))) for t in t_stats]
    else:
        p_values = [float(np.mean(scores >= s)) for s in scores]
    bh_adjusted = benjamini_hochberg_adjusted_pvalues(p_values)

    return {
        "deflated_sharpe_ratio": dsr,
        "probabilistic_sharpe_ratio": psr,
        "pbo_style": pbo,
        "white_reality_check": white,
        "spa": spa,
        "multiple_testing": {
            "method": "benjamini_hochberg",
            "n_hypotheses": int(len(p_values)),
            "raw_pvalues": p_values,
            "bh_adjusted_pvalues": bh_adjusted,
            "min_raw_pvalue": float(min(p_values)) if p_values else 1.0,
            "min_bh_adjusted_pvalue": float(min(bh_adjusted)) if bh_adjusted else 1.0,
        },
    }




def build_scenario_attribution_and_guardrails(
    *,
    baseline_metrics: dict[str, float],
    scenario_results: list[dict[str, object]],
) -> dict[str, list[dict[str, object]]]:
    baseline_sharpe = float(baseline_metrics.get("sharpe", 0.0))
    baseline_drawdown = float(baseline_metrics.get("max_drawdown", 0.0))
    baseline_return = float(baseline_metrics.get("total_return", 0.0))

    attribution: list[dict[str, object]] = []
    guardrails: list[dict[str, object]] = []
    for row in scenario_results:
        name = str(row.get("name", "unnamed"))
        metrics = row.get("metrics", {})
        if not isinstance(metrics, dict):
            metrics = {}
        sharpe = float(metrics.get("sharpe", 0.0))
        max_drawdown = float(metrics.get("max_drawdown", 0.0))
        total_return = float(metrics.get("total_return", 0.0))
        pnl_total = float(row.get("pnl_total", 0.0))

        attribution.append({
            "scenario": name,
            "pnl_total": pnl_total,
            "delta_total_return": total_return - baseline_return,
            "delta_sharpe": sharpe - baseline_sharpe,
            "delta_max_drawdown": max_drawdown - baseline_drawdown,
        })

        checks = {
            "drawdown": max_drawdown >= min(-0.5, baseline_drawdown - 0.15),
            "sharpe_floor": sharpe >= min(0.0, baseline_sharpe - 1.0),
            "return_floor": total_return >= baseline_return - 0.30,
        }
        guardrails.append({
            "scenario": name,
            "passed": bool(all(checks.values())),
            "checks": checks,
        })

    return {
        "scenario_attribution": attribution,
        "scenario_guardrails": guardrails,
    }
def _compute_pbo_style(
    *,
    ranked_rows: list[dict[str, object]],
    score_key: str,
    n_monte_carlo: int,
    seed: int,
) -> dict[str, object]:
    n = len(ranked_rows)
    if n < 4:
        return {
            "n_combinations": n,
            "n_monte_carlo": int(n_monte_carlo),
            "probability_of_overfitting": 0.0,
            "median_logit": 0.0,
        }
    rng = np.random.default_rng(seed)
    logits: list[float] = []
    overfit_count = 0
    for _ in range(int(n_monte_carlo)):
        perm = rng.permutation(n)
        split = n // 2
        in_idx = perm[:split]
        out_idx = perm[split:]
        if in_idx.size == 0 or out_idx.size == 0:
            continue
        in_scores = [float(ranked_rows[i].get(score_key, 0.0)) for i in in_idx]
        best_local = int(np.argmax(in_scores))
        chosen_score = in_scores[best_local]
        out_scores = np.array([float(ranked_rows[i].get(score_key, 0.0)) for i in out_idx], dtype=float)
        percentile = float(np.mean(out_scores <= chosen_score))
        percentile = min(max(percentile, 1e-6), 1.0 - 1e-6)
        logits.append(float(np.log(percentile / (1.0 - percentile))))
        if percentile < 0.5:
            overfit_count += 1
    median_logit = float(np.median(logits)) if logits else 0.0
    pbo = float(overfit_count / len(logits)) if logits else 0.0
    return {
        "n_combinations": n,
        "n_monte_carlo": int(n_monte_carlo),
        "probability_of_overfitting": pbo,
        "median_logit": median_logit,
    }




def compute_capacity_frontier(
    *,
    expected_alpha_bps: float,
    realized_slippage_bps: float,
    average_participation_rate: float,
    base_aum: float = 1_000_000.0,
    turnover_total: float = 1.0,
    scales: np.ndarray | None = None,
) -> list[dict[str, float]]:
    """Compute a capacity frontier: expected alpha net of trading costs by AUM scale."""

    levels = np.asarray(scales, dtype=float) if scales is not None else np.array([0.25, 0.5, 1.0, 2.0, 4.0, 8.0], dtype=float)
    levels = levels[np.isfinite(levels) & (levels > 0.0)]
    if levels.size == 0:
        levels = np.array([1.0], dtype=float)

    base_participation = max(float(average_participation_rate), 0.01)
    base_slippage = max(float(realized_slippage_bps), 0.0)
    alpha = float(expected_alpha_bps)
    turnover = max(float(turnover_total), 0.0)

    rows: list[dict[str, float]] = []
    for scale in levels:
        aum = float(base_aum) * float(scale)
        scaled_participation = base_participation * float(scale)
        impact_scale = np.sqrt(max(scaled_participation, 1e-12) / base_participation)
        projected_slippage_bps = base_slippage * impact_scale
        net_alpha_bps = alpha - projected_slippage_bps * turnover
        rows.append(
            {
                "aum": float(aum),
                "aum_scale": float(scale),
                "participation_rate": float(scaled_participation),
                "expected_alpha_bps": float(alpha),
                "projected_slippage_bps": float(projected_slippage_bps),
                "expected_alpha_net_cost_bps": float(net_alpha_bps),
            }
        )
    return rows
def _build_capacity_diagnostics(
    *,
    metrics: dict[str, float],
    turnover_stats: dict[str, float],
    cost_totals: dict[str, float],
    fills: list[dict[str, object]],
) -> dict[str, object]:
    realized_parts = np.array([float(row.get("participation_rate", 0.0)) for row in fills], dtype=float)
    avg_participation = float(np.mean(realized_parts)) if realized_parts.size else 0.0
    base_participation = max(avg_participation, 0.01)

    turnover_total = max(float(turnover_stats.get("total", 0.0)), 1e-12)
    realized_slippage_bps = float(cost_totals.get("slippage", 0.0)) / turnover_total * 10_000.0
    cagr = float(metrics.get("cagr", 0.0))

    levels = np.array([0.01, 0.02, 0.05, 0.10, 0.20, 0.40], dtype=float)
    slippage_curve: list[dict[str, float]] = []
    degradation_curve: list[dict[str, float]] = []
    realized_alpha_bps = float(metrics.get("turnover_adjusted_return", metrics.get("cagr", 0.0))) * 10_000.0
    for level in levels:
        scale = np.sqrt(level / base_participation) if base_participation > 0 else 1.0
        expected_bps = realized_slippage_bps * scale
        slippage_curve.append({"participation_rate": float(level), "expected_slippage_bps": float(expected_bps)})
        turnover_penalty = (expected_bps / 10_000.0) * float(turnover_stats.get("total", 0.0))
        degradation_curve.append(
            {
                "participation_rate": float(level),
                "projected_cagr": float(cagr - turnover_penalty),
                "projected_net_return": float(metrics.get("total_return", 0.0) - turnover_penalty),
            }
        )

    frontier = compute_capacity_frontier(
        expected_alpha_bps=realized_alpha_bps,
        realized_slippage_bps=realized_slippage_bps,
        average_participation_rate=avg_participation,
        turnover_total=float(turnover_stats.get("total", 0.0)),
    )

    return {
        "average_participation_rate": avg_participation,
        "realized_slippage_bps": realized_slippage_bps,
        "expected_alpha_bps": realized_alpha_bps,
        "expected_slippage_curve": slippage_curve,
        "performance_degradation_curve": degradation_curve,
        "capacity_frontier": frontier,
    }


def _to_1d_float(values: np.ndarray) -> np.ndarray:
    arr = np.asarray(values, dtype=float)
    return arr.reshape(-1)


def _summarize_distribution(values: list[float], lo_q: float, hi_q: float) -> dict[str, float]:
    arr = np.asarray(values, dtype=float)
    if arr.size == 0:
        return {"lower": 0.0, "median": 0.0, "upper": 0.0}
    return {
        "lower": float(np.percentile(arr, lo_q)),
        "median": float(np.percentile(arr, 50.0)),
        "upper": float(np.percentile(arr, hi_q)),
    }


def _format_ticker_line(
    ticker: str,
    result: CrossSectionalResult,
    max_ticker_len: int,
    score_override: float | None = None,
) -> str:
    score = score_override if score_override is not None else result.scores.get(ticker, 0.0)
    metrics = result.metrics.get(ticker, {})
    metric_parts = [f"score={score:.4f}"]
    for key, value in metrics.items():
        metric_parts.append(f"{key}={value:.4f}")
    metric_text = ", ".join(metric_parts)
    padded_ticker = ticker.ljust(max_ticker_len)
    return f"  - {padded_ticker}: {metric_text}"


def _max_ticker_len(universe: list[str], result: CrossSectionalResult) -> int:
    candidates = list(universe)
    candidates.extend(result.skipped.keys())
    candidates.extend(result.scores.keys())
    if not candidates:
        return 1
    return max(1, max(len(ticker) for ticker in candidates))


def _format_ticker_line_time_series(
    ticker: str,
    result: TimeSeriesResult,
    max_ticker_len: int,
    score_override: float | None = None,
) -> str:
    score = score_override if score_override is not None else result.scores.get(ticker, 0.0)
    metrics = result.metrics.get(ticker, {})
    metric_parts = [f"score={score:.4f}"]
    for key, value in metrics.items():
        metric_parts.append(f"{key}={value:.4f}")
    metric_text = ", ".join(metric_parts)
    padded_ticker = ticker.ljust(max_ticker_len)
    return f"  - {padded_ticker}: {metric_text}"


def _max_ticker_len_time_series(universe: list[str], result: TimeSeriesResult) -> int:
    candidates = list(universe)
    candidates.extend(result.skipped.keys())
    candidates.extend(result.scores.keys())
    if not candidates:
        return 1
    return max(1, max(len(ticker) for ticker in candidates))

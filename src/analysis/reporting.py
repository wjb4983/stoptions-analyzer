from __future__ import annotations

from datetime import datetime
from typing import Iterable

import numpy as np

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

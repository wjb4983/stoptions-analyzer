from __future__ import annotations

from datetime import datetime
from typing import Iterable

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

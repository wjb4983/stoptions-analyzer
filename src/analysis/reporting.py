from __future__ import annotations

from datetime import datetime
from typing import Iterable

from .cross_sectional.base import CrossSectionalResult


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

    lines.append("Top Winners:")
    if result.longs:
        for ticker in sorted(result.longs, key=lambda t: result.scores.get(t, 0.0), reverse=True):
            lines.append(f"  - {ticker}: {result.scores.get(ticker, 0.0):.4f}")
    else:
        lines.append("  (none)")
    lines.append("")

    lines.append("Top Losers:")
    if result.shorts:
        for ticker in sorted(result.shorts, key=lambda t: result.scores.get(t, 0.0)):
            lines.append(f"  - {ticker}: {result.scores.get(ticker, 0.0):.4f}")
    else:
        lines.append("  (none)")
    lines.append("")

    lines.append("Ranking (best to worst):")
    if result.ranking:
        for ticker, score in result.ranking:
            lines.append(f"  - {ticker}: {score:.4f}")
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
        lines.append(f"  - {ticker}: {score:.4f}")
    for ticker, reason in sorted(skipped.items()):
        if ticker not in scored:
            lines.append(f"  - {ticker}: skipped ({reason})")

    return "\n".join(lines)

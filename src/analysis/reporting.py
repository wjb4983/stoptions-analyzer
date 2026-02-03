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
        for ticker in result.longs:
            lines.append(f"  - {ticker}: {result.scores.get(ticker, 0.0):.4f}")
    else:
        lines.append("  (none)")
    lines.append("")

    lines.append("Top Losers:")
    if result.shorts:
        for ticker in result.shorts:
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

    if result.skipped:
        lines.append("")
        lines.append("Skipped:")
        for ticker, reason in result.skipped.items():
            lines.append(f"  - {ticker}: {reason}")

    return "\n".join(lines)

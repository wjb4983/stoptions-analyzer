from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class CrossSectionalResult:
    scores: dict[str, float]
    ranking: list[tuple[str, float]]
    longs: list[str]
    shorts: list[str]
    weights: dict[str, float]
    metrics: dict[str, dict[str, float]] = field(default_factory=dict)
    metadata: dict[str, Any] = field(default_factory=dict)
    skipped: dict[str, str] = field(default_factory=dict)

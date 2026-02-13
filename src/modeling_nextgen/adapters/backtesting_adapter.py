from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from src.backtesting.walk_forward import WalkForwardFold, build_walk_forward_folds


@dataclass(frozen=True)
class BacktestingBridge:
    """Adapter exposing next-gen validation hooks into existing backtesting APIs."""

    def build_walk_forward_folds(self, **kwargs: Any) -> list[WalkForwardFold]:
        return build_walk_forward_folds(**kwargs)

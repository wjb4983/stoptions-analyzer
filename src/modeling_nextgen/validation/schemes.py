from __future__ import annotations

from dataclasses import dataclass

import numpy as np


@dataclass(frozen=True)
class PurgedCrossValidator:
    name: str = "purged_cv"

    def split(self, n_samples: int) -> list[tuple[np.ndarray, np.ndarray]]:
        idx = np.arange(n_samples)
        return [(idx, idx)]


@dataclass(frozen=True)
class WalkForwardValidator:
    name: str = "walk_forward"

    def split(self, n_samples: int) -> list[tuple[np.ndarray, np.ndarray]]:
        idx = np.arange(n_samples)
        return [(idx, idx)]


@dataclass(frozen=True)
class StressValidator:
    name: str = "stress"

    def split(self, n_samples: int) -> list[tuple[np.ndarray, np.ndarray]]:
        idx = np.arange(n_samples)
        return [(idx, idx)]

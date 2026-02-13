from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import numpy as np


@dataclass(frozen=True)
class SurfaceFeatureBuilder:
    name: str = "surface"

    def build(self, frame: Any) -> dict[str, np.ndarray]:
        return {}


@dataclass(frozen=True)
class RegimeFeatureBuilder:
    name: str = "regime"

    def build(self, frame: Any) -> dict[str, np.ndarray]:
        return {}


@dataclass(frozen=True)
class FlowFeatureBuilder:
    name: str = "flow"

    def build(self, frame: Any) -> dict[str, np.ndarray]:
        return {}

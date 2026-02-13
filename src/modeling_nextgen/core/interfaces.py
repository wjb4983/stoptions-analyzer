from __future__ import annotations

from pathlib import Path
from typing import Any, Protocol

from .schemas import (
    OptionSurfaceTensorPayload,
    PanelFeaturesPayload,
    RegimeLabelsPayload,
    UncertaintyOutputPayload,
)


class NextGenModelInterface(Protocol):
    """Protocol for next-generation modeling workflows."""

    name: str

    def fit(
        self,
        panel_features: PanelFeaturesPayload,
        regime_labels: RegimeLabelsPayload,
        option_surfaces: OptionSurfaceTensorPayload | None = None,
    ) -> None:
        ...

    def predict(
        self,
        panel_features: PanelFeaturesPayload,
        option_surfaces: OptionSurfaceTensorPayload | None = None,
    ) -> RegimeLabelsPayload:
        ...

    def predict_proba(
        self,
        panel_features: PanelFeaturesPayload,
        option_surfaces: OptionSurfaceTensorPayload | None = None,
    ) -> UncertaintyOutputPayload:
        ...

    def predict_distribution(
        self,
        panel_features: PanelFeaturesPayload,
        option_surfaces: OptionSurfaceTensorPayload | None = None,
    ) -> UncertaintyOutputPayload:
        ...

    def explain(
        self,
        panel_features: PanelFeaturesPayload,
        option_surfaces: OptionSurfaceTensorPayload | None = None,
    ) -> dict[str, Any]:
        ...

    def calibrate(
        self,
        panel_features: PanelFeaturesPayload,
        regime_labels: RegimeLabelsPayload,
    ) -> None:
        ...

    def save(self, path: str | Path) -> None:
        ...

    def load(self, path: str | Path) -> None:
        ...

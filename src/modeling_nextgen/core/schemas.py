from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import numpy as np
from numpy.typing import NDArray


@dataclass(frozen=True)
class PanelFeaturesPayload:
    """Panel features aligned as asset × date tensors."""

    assets: NDArray[np.str_]
    dates: NDArray[np.datetime64]
    features: dict[str, NDArray[np.floating[Any]]]


@dataclass(frozen=True)
class OptionSurfaceTensorPayload:
    """Options surface tensor aligned as moneyness × tenor × date."""

    moneyness: NDArray[np.floating[Any]]
    tenors: NDArray[np.floating[Any]]
    dates: NDArray[np.datetime64]
    values: NDArray[np.floating[Any]]


@dataclass(frozen=True)
class RegimeLabelsPayload:
    """Discrete regime labels with optional class probabilities."""

    dates: NDArray[np.datetime64]
    labels: NDArray[np.integer[Any]]
    probabilities: NDArray[np.floating[Any]] | None = None


@dataclass(frozen=True)
class UncertaintyOutputPayload:
    """Model uncertainty outputs for forecasts and regime inference."""

    dates: NDArray[np.datetime64]
    predictive_std: NDArray[np.floating[Any]]
    confidence_lower: NDArray[np.floating[Any]] | None = None
    confidence_upper: NDArray[np.floating[Any]] | None = None
    regime_uncertainty: NDArray[np.floating[Any]] | None = None

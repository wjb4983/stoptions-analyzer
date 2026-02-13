"""Next-generation modeling package.

This package provides opt-in contracts and adapters for progressively
introducing advanced modeling workflows without changing legacy paradigms.
"""

from .core.contracts import FeatureBuilder, Model, ProbabilisticModel, Validator
from .core.interfaces import NextGenModelInterface
from .core.schemas import (
    OptionSurfaceTensorPayload,
    PanelFeaturesPayload,
    RegimeLabelsPayload,
    UncertaintyOutputPayload,
)
from .core.config import NextGenModelingConfig

__all__ = [
    "FeatureBuilder",
    "Model",
    "ProbabilisticModel",
    "Validator",
    "NextGenModelingConfig",
    "NextGenModelInterface",
    "PanelFeaturesPayload",
    "OptionSurfaceTensorPayload",
    "RegimeLabelsPayload",
    "UncertaintyOutputPayload",
]

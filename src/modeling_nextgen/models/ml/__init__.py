from dataclasses import dataclass

from ..base import NextGenModelBase
from .panel_baselines import (
    ElasticNetBaseline,
    LogitBaseline,
    PanelSplit,
    PanelWalkForwardSplitter,
    RandomForestBaseline,
    TreeBoostingBaseline,
)


@dataclass
class MLModel(NextGenModelBase):
    name: str = "ml"


__all__ = [
    "MLModel",
    "ElasticNetBaseline",
    "LogitBaseline",
    "TreeBoostingBaseline",
    "RandomForestBaseline",
    "PanelSplit",
    "PanelWalkForwardSplitter",
]

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
from .multitask import CalibrationArtifact, MultiTaskRiskModel, TaskSpec


@dataclass
class MLModel(NextGenModelBase):
    name: str = "ml"


__all__ = [
    "MLModel",
    "ElasticNetBaseline",
    "LogitBaseline",
    "TreeBoostingBaseline",
    "RandomForestBaseline",
    "TaskSpec",
    "CalibrationArtifact",
    "MultiTaskRiskModel",
    "PanelSplit",
    "PanelWalkForwardSplitter",
]

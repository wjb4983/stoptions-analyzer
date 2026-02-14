from dataclasses import dataclass

from ..base import NextGenModelBase
from .sequence_encoder import SequenceEncoder


@dataclass
class DeepModel(NextGenModelBase):
    name: str = "deep"


__all__ = ["DeepModel", "SequenceEncoder"]

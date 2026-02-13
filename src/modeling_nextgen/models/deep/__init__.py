from dataclasses import dataclass

from ..base import NextGenModelBase


@dataclass
class DeepModel(NextGenModelBase):
    name: str = "deep"


__all__ = ["DeepModel"]

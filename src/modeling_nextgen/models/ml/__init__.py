from dataclasses import dataclass

from ..base import NextGenModelBase


@dataclass
class MLModel(NextGenModelBase):
    name: str = "ml"


__all__ = ["MLModel"]

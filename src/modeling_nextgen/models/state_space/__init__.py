from dataclasses import dataclass

from ..base import NextGenModelBase


@dataclass
class StateSpaceModel(NextGenModelBase):
    name: str = "state_space"


__all__ = ["StateSpaceModel"]

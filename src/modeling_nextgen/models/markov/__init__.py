from dataclasses import dataclass

from ..base import NextGenModelBase


@dataclass
class MarkovModel(NextGenModelBase):
    name: str = "markov"


__all__ = ["MarkovModel"]

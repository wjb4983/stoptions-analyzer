from dataclasses import dataclass

from ..base import NextGenModelBase


@dataclass
class BayesModel(NextGenModelBase):
    name: str = "bayes"


__all__ = ["BayesModel"]

from dataclasses import dataclass

from ..base import NextGenModelBase
from .regime_switching import (
    RegimeSwitchingConfig,
    RegimeSwitchingOutput,
    estimate_transition_matrix,
    fit_regime_switching_model,
)


@dataclass
class MarkovModel(NextGenModelBase):
    name: str = "markov"


__all__ = [
    "MarkovModel",
    "RegimeSwitchingConfig",
    "RegimeSwitchingOutput",
    "estimate_transition_matrix",
    "fit_regime_switching_model",
]

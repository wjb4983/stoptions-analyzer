from dataclasses import dataclass

from ..base import NextGenModelBase
from .vol_factor_kalman import (
    VolFactorKalmanConfig,
    VolFactorKalmanOutput,
    infer_observation_matrix,
    run_vol_factor_kalman,
)


@dataclass
class StateSpaceModel(NextGenModelBase):
    name: str = "state_space"


__all__ = [
    "StateSpaceModel",
    "VolFactorKalmanConfig",
    "VolFactorKalmanOutput",
    "infer_observation_matrix",
    "run_vol_factor_kalman",
]

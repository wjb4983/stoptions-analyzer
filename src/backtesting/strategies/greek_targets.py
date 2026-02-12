from __future__ import annotations

from dataclasses import dataclass

import numpy as np


@dataclass(frozen=True)
class GreekNeutralTargetRequest:
    base_targets: np.ndarray
    delta: np.ndarray | None = None
    gamma: np.ndarray | None = None
    vega: np.ndarray | None = None
    theta: np.ndarray | None = None
    neutralize: tuple[str, ...] = ("delta",)


def build_greek_neutral_targets(request: GreekNeutralTargetRequest) -> np.ndarray:
    targets = np.asarray(request.base_targets, dtype=float).copy()
    if targets.ndim != 2:
        raise ValueError("base_targets must be a 2D array [periods, assets]")

    sensitivities = {
        "delta": None if request.delta is None else np.asarray(request.delta, dtype=float),
        "gamma": None if request.gamma is None else np.asarray(request.gamma, dtype=float),
        "vega": None if request.vega is None else np.asarray(request.vega, dtype=float),
        "theta": None if request.theta is None else np.asarray(request.theta, dtype=float),
    }

    for greek in request.neutralize:
        arr = sensitivities.get(greek)
        if arr is None or arr.shape != targets.shape:
            continue
        exposure = np.sum(targets * arr, axis=1, keepdims=True)
        denom = np.sum(arr * arr, axis=1, keepdims=True)
        hedge = np.divide(exposure, denom, out=np.zeros_like(exposure), where=denom > 0.0)
        targets = targets - hedge * arr

    gross = np.sum(np.abs(targets), axis=1, keepdims=True)
    targets = np.divide(targets, gross, out=np.zeros_like(targets), where=gross > 0.0)
    return targets

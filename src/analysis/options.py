from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import numpy as np


@dataclass(frozen=True)
class OptionExposureSnapshot:
    delta: np.ndarray
    gamma: np.ndarray
    vega: np.ndarray
    theta: np.ndarray


def aggregate_option_exposures(
    *,
    positions: np.ndarray,
    portfolio_weights: np.ndarray,
    delta: np.ndarray | None = None,
    gamma: np.ndarray | None = None,
    vega: np.ndarray | None = None,
    theta: np.ndarray | None = None,
    multipliers: np.ndarray | None = None,
    contract_adjustments: np.ndarray | None = None,
) -> OptionExposureSnapshot:
    pos = np.asarray(positions, dtype=float)
    if pos.ndim != 2:
        raise ValueError("positions must be 2D [periods, assets]")

    rows, assets = pos.shape
    w = np.asarray(portfolio_weights, dtype=float)
    if w.shape != (assets,):
        raise ValueError("portfolio_weights must have shape (assets,)")

    multi = np.ones((rows, assets), dtype=float)
    if multipliers is not None:
        m = np.asarray(multipliers, dtype=float)
        if m.shape == (assets,):
            multi *= m[None, :]
        elif m.shape == (rows, assets):
            multi *= m

    if contract_adjustments is not None:
        adj = np.asarray(contract_adjustments, dtype=float)
        if adj.shape == (assets,):
            multi *= adj[None, :]
        elif adj.shape == (rows, assets):
            multi *= adj

    signed_notional = pos * w[None, :] * multi

    def _resolve(values: np.ndarray | None) -> np.ndarray:
        if values is None:
            return np.zeros_like(signed_notional)
        arr = np.asarray(values, dtype=float)
        if arr.shape != signed_notional.shape:
            return np.zeros_like(signed_notional)
        return arr

    return OptionExposureSnapshot(
        delta=np.sum(signed_notional * _resolve(delta), axis=1),
        gamma=np.sum(signed_notional * _resolve(gamma), axis=1),
        vega=np.sum(signed_notional * _resolve(vega), axis=1),
        theta=np.sum(signed_notional * _resolve(theta), axis=1),
    )


def summarize_lifecycle_events(events: list[dict[str, Any]]) -> dict[str, int]:
    counts: dict[str, int] = {
        "expiry_cash_settlement": 0,
        "assignment_exercise": 0,
        "contract_roll": 0,
        "corporate_action_adjustment": 0,
    }
    for row in events:
        name = str(row.get("event", "")).strip()
        if name in counts:
            counts[name] += 1
    return counts


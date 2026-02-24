from __future__ import annotations

from dataclasses import dataclass

import numpy as np

from .greek_targets import GreekNeutralTargetRequest, build_greek_neutral_targets


@dataclass(frozen=True)
class HedgeRebalanceConfig:
    every_n_bars: int | None = None
    delta_threshold: float | None = None
    vega_threshold: float | None = None


@dataclass(frozen=True)
class VolHedgingPolicy:
    hedge_delta: bool = True
    hedge_vega: bool = False
    rebalance: HedgeRebalanceConfig = HedgeRebalanceConfig(every_n_bars=1)
    max_hedge_turnover: float | None = None
    max_hedge_notional: float | None = None


@dataclass(frozen=True)
class HedgedTargetResult:
    hedged_targets: np.ndarray
    hedge_targets: np.ndarray
    rebalance_mask: np.ndarray


def apply_vol_hedging_policy(
    *,
    base_targets: np.ndarray,
    delta: np.ndarray | None,
    vega: np.ndarray | None,
    policy: VolHedgingPolicy | None,
) -> HedgedTargetResult:
    targets = np.asarray(base_targets, dtype=float)
    if targets.ndim != 2:
        raise ValueError("base_targets must be a 2D array [periods, assets]")

    rows, _ = targets.shape
    if policy is None or (not policy.hedge_delta and not policy.hedge_vega):
        return HedgedTargetResult(
            hedged_targets=targets.copy(),
            hedge_targets=np.zeros_like(targets),
            rebalance_mask=np.zeros(rows, dtype=bool),
        )

    delta_arr = None if delta is None else np.asarray(delta, dtype=float)
    vega_arr = None if vega is None else np.asarray(vega, dtype=float)

    hedges = np.zeros_like(targets)
    rebalance_mask = np.zeros(rows, dtype=bool)
    prev_hedge = np.zeros(targets.shape[1], dtype=float)
    every_n = None if policy.rebalance.every_n_bars is None else max(1, int(policy.rebalance.every_n_bars))

    for idx in range(rows):
        base_row = targets[idx]
        should_rebalance = idx == 0 or every_n is None or (idx % every_n == 0)
        if not should_rebalance:
            hedges[idx] = prev_hedge
            continue

        threshold_hit = False
        if policy.rebalance.delta_threshold is not None and delta_arr is not None and delta_arr.shape == targets.shape:
            delta_exp = float(np.sum(base_row * delta_arr[idx]))
            threshold_hit = threshold_hit or abs(delta_exp) >= float(policy.rebalance.delta_threshold)
        if policy.rebalance.vega_threshold is not None and vega_arr is not None and vega_arr.shape == targets.shape:
            vega_exp = float(np.sum(base_row * vega_arr[idx]))
            threshold_hit = threshold_hit or abs(vega_exp) >= float(policy.rebalance.vega_threshold)

        if (
            policy.rebalance.delta_threshold is not None
            or policy.rebalance.vega_threshold is not None
        ) and not threshold_hit and idx != 0:
            hedges[idx] = prev_hedge
            continue

        neutralize: list[str] = []
        if policy.hedge_delta and delta_arr is not None and delta_arr.shape == targets.shape:
            neutralize.append("delta")
        if policy.hedge_vega and vega_arr is not None and vega_arr.shape == targets.shape:
            neutralize.append("vega")

        if not neutralize:
            hedges[idx] = prev_hedge
            continue

        neutral_row = build_greek_neutral_targets(
            GreekNeutralTargetRequest(
                base_targets=base_row.reshape(1, -1),
                delta=None if delta_arr is None else delta_arr[idx].reshape(1, -1),
                vega=None if vega_arr is None else vega_arr[idx].reshape(1, -1),
                neutralize=tuple(neutralize),
            )
        )[0]
        candidate = neutral_row - base_row

        max_notional = policy.max_hedge_notional
        gross = float(np.sum(np.abs(candidate)))
        if max_notional is not None and gross > float(max_notional) and gross > 0.0:
            candidate *= float(max_notional) / gross

        max_turnover = policy.max_hedge_turnover
        delta_turnover = float(np.sum(np.abs(candidate - prev_hedge)))
        if max_turnover is not None and delta_turnover > float(max_turnover) and delta_turnover > 0.0:
            candidate = prev_hedge + (candidate - prev_hedge) * (float(max_turnover) / delta_turnover)

        hedges[idx] = candidate
        prev_hedge = candidate
        rebalance_mask[idx] = True

    return HedgedTargetResult(
        hedged_targets=targets + hedges,
        hedge_targets=hedges,
        rebalance_mask=rebalance_mask,
    )

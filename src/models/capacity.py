from __future__ import annotations

from dataclasses import dataclass

import numpy as np


@dataclass(frozen=True)
class StrategyCapacityInput:
    strategy_id: str
    sharpe: float
    expected_alpha_bps: float
    liquidity: float
    turnover: float
    impact_coefficient: float
    is_niche: bool = False


@dataclass(frozen=True)
class CapacityConfig:
    base_capital: float = 1_000_000.0
    min_liquidity_floor: float = 1e-9
    min_turnover_floor: float = 1e-9
    max_capacity_score: float = 1.0


@dataclass(frozen=True)
class PromotionCapacityPolicy:
    high_sharpe_threshold: float = 1.5
    min_capacity_score: float = 0.25


def estimate_strategy_capacity(
    strategies: list[StrategyCapacityInput],
    *,
    config: CapacityConfig | None = None,
) -> list[dict[str, float | str | bool]]:
    cfg = config or CapacityConfig()
    rows: list[dict[str, float | str | bool]] = []
    for strategy in strategies:
        liquidity = max(float(strategy.liquidity), cfg.min_liquidity_floor)
        turnover = max(float(strategy.turnover), cfg.min_turnover_floor)
        impact = max(float(strategy.impact_coefficient), 0.0)

        modeled_capacity = liquidity / (turnover * (1.0 + impact))
        capacity_score = min(float(modeled_capacity / max(cfg.base_capital, 1.0)), cfg.max_capacity_score)

        rows.append(
            {
                "strategy_id": strategy.strategy_id,
                "sharpe": float(strategy.sharpe),
                "expected_alpha_bps": float(strategy.expected_alpha_bps),
                "liquidity": float(strategy.liquidity),
                "turnover": float(strategy.turnover),
                "impact_coefficient": float(strategy.impact_coefficient),
                "modeled_capacity": float(modeled_capacity),
                "capacity_score": max(float(capacity_score), 0.0),
                "is_niche": bool(strategy.is_niche),
            }
        )
    return rows


def compute_alpha_decay_under_capital(
    *,
    expected_alpha_bps: float,
    capital_levels: np.ndarray,
    modeled_capacity: float,
) -> list[dict[str, float]]:
    levels = np.asarray(capital_levels, dtype=float)
    levels = levels[np.isfinite(levels) & (levels > 0.0)]
    if levels.size == 0:
        return []

    cap = max(float(modeled_capacity), 1e-9)
    alpha0 = float(expected_alpha_bps)
    rows: list[dict[str, float]] = []
    for capital in levels:
        overload = max(float(capital) / cap - 1.0, 0.0)
        decay_factor = float(np.exp(-overload))
        realized_alpha = alpha0 * decay_factor
        rows.append(
            {
                "capital": float(capital),
                "decay_factor": float(decay_factor),
                "realized_alpha_bps": float(realized_alpha),
                "alpha_decay_bps": float(alpha0 - realized_alpha),
            }
        )
    return rows


def rank_and_allocate_by_capacity(
    strategies: list[StrategyCapacityInput],
    *,
    deploy_capital: float,
    config: CapacityConfig | None = None,
) -> list[dict[str, float | str | bool]]:
    estimated = estimate_strategy_capacity(strategies, config=config)
    total_score = 0.0
    rows: list[dict[str, float | str | bool]] = []
    for row in estimated:
        decay = compute_alpha_decay_under_capital(
            expected_alpha_bps=float(row["expected_alpha_bps"]),
            capital_levels=np.array([deploy_capital], dtype=float),
            modeled_capacity=float(row["modeled_capacity"]),
        )
        realized_alpha = float(decay[0]["realized_alpha_bps"]) if decay else float(row["expected_alpha_bps"])
        capacity_adjusted_rr = float(row["sharpe"]) * float(row["capacity_score"]) * (realized_alpha / max(float(row["expected_alpha_bps"]), 1e-9))
        total_score += max(capacity_adjusted_rr, 0.0)
        rows.append(
            {
                **row,
                "deploy_capital": float(deploy_capital),
                "realized_alpha_bps": float(realized_alpha),
                "capacity_adjusted_score": float(capacity_adjusted_rr),
            }
        )

    rows.sort(key=lambda x: float(x["capacity_adjusted_score"]), reverse=True)
    if total_score <= 0.0:
        equal = 1.0 / max(len(rows), 1)
        for row in rows:
            row["allocation_weight"] = float(equal)
        return rows

    for row in rows:
        row["allocation_weight"] = float(max(float(row["capacity_adjusted_score"]), 0.0) / total_score)
    return rows


def is_promotion_blocked_for_capacity(
    *,
    sharpe: float,
    capacity_score: float,
    is_niche: bool,
    policy: PromotionCapacityPolicy | None = None,
) -> bool:
    cfg = policy or PromotionCapacityPolicy()
    if bool(is_niche):
        return False
    return float(sharpe) >= cfg.high_sharpe_threshold and float(capacity_score) < cfg.min_capacity_score

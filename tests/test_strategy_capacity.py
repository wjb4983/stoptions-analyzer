from __future__ import annotations

import numpy as np

from models.capacity import (
    CapacityConfig,
    PromotionCapacityPolicy,
    StrategyCapacityInput,
    compute_alpha_decay_under_capital,
    estimate_strategy_capacity,
    is_promotion_blocked_for_capacity,
    rank_and_allocate_by_capacity,
)


def test_estimate_capacity_and_alpha_decay_monotonicity() -> None:
    rows = estimate_strategy_capacity(
        [
            StrategyCapacityInput(
                strategy_id="liq_high",
                sharpe=1.1,
                expected_alpha_bps=24.0,
                liquidity=20_000_000.0,
                turnover=0.8,
                impact_coefficient=0.2,
            )
        ],
        config=CapacityConfig(base_capital=1_000_000.0),
    )
    assert len(rows) == 1
    assert float(rows[0]["modeled_capacity"]) > 0.0
    assert 0.0 <= float(rows[0]["capacity_score"]) <= 1.0

    decay = compute_alpha_decay_under_capital(
        expected_alpha_bps=24.0,
        capital_levels=np.array([500_000.0, 1_000_000.0, 2_000_000.0], dtype=float),
        modeled_capacity=1_000_000.0,
    )
    alphas = [float(row["realized_alpha_bps"]) for row in decay]
    assert alphas[2] < alphas[1] <= alphas[0]


def test_rank_and_allocate_prefers_higher_capacity_adjusted_score() -> None:
    ranked = rank_and_allocate_by_capacity(
        [
            StrategyCapacityInput(
                strategy_id="stable_large",
                sharpe=1.2,
                expected_alpha_bps=20.0,
                liquidity=30_000_000.0,
                turnover=0.6,
                impact_coefficient=0.15,
            ),
            StrategyCapacityInput(
                strategy_id="fragile_small",
                sharpe=2.0,
                expected_alpha_bps=30.0,
                liquidity=1_000_000.0,
                turnover=3.5,
                impact_coefficient=1.4,
            ),
        ],
        deploy_capital=2_000_000.0,
    )

    assert ranked[0]["strategy_id"] == "stable_large"
    total = sum(float(row["allocation_weight"]) for row in ranked)
    assert np.isclose(total, 1.0)


def test_promotion_block_policy_allows_niche_override() -> None:
    policy = PromotionCapacityPolicy(high_sharpe_threshold=1.5, min_capacity_score=0.3)

    blocked = is_promotion_blocked_for_capacity(sharpe=2.0, capacity_score=0.1, is_niche=False, policy=policy)
    niche_allowed = is_promotion_blocked_for_capacity(sharpe=2.0, capacity_score=0.1, is_niche=True, policy=policy)

    assert blocked is True
    assert niche_allowed is False

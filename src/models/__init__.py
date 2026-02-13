from .base import BaseParadigmModel, ModelExplanation, ModelInterface
from .ensemble import EnsembleOutput, ModelEnsembler
from .deployment import (
    ModelSlots,
    PromotionGates,
    REASON_PROMOTION_GATES_PASSED,
    REASON_RISK_BREACH,
    REASON_ROLLBACK_TO_PRIOR_CHAMPION,
    REASON_ROBUSTNESS_FAILURE,
    REASON_CAPACITY_CONSTRAINT,
    REASON_SHADOW_UNDERPERFORMANCE,
    REASON_STABILITY_FAILURE,
    SlotEvent,
)
from .registry import MODEL_REGISTRY, ModelActivation, ModelActivationConfig, activated_models, create_model
from .capacity import (
    CapacityConfig,
    PromotionCapacityPolicy,
    StrategyCapacityInput,
    compute_alpha_decay_under_capital,
    estimate_strategy_capacity,
    is_promotion_blocked_for_capacity,
    rank_and_allocate_by_capacity,
)
from .robustness import RobustnessThresholds, build_robustness_scorecards

__all__ = [
    "BaseParadigmModel",
    "ModelExplanation",
    "ModelInterface",
    "EnsembleOutput",
    "ModelEnsembler",
    "MODEL_REGISTRY",
    "ModelActivation",
    "ModelActivationConfig",
    "activated_models",
    "create_model",
    "ModelSlots",
    "PromotionGates",
    "SlotEvent",
    "REASON_PROMOTION_GATES_PASSED",
    "REASON_RISK_BREACH",
    "REASON_ROLLBACK_TO_PRIOR_CHAMPION",
    "REASON_ROBUSTNESS_FAILURE",
    "REASON_CAPACITY_CONSTRAINT",
    "REASON_SHADOW_UNDERPERFORMANCE",
    "REASON_STABILITY_FAILURE",
    "RobustnessThresholds",
    "build_robustness_scorecards",
    "StrategyCapacityInput",
    "CapacityConfig",
    "PromotionCapacityPolicy",
    "estimate_strategy_capacity",
    "compute_alpha_decay_under_capital",
    "rank_and_allocate_by_capacity",
    "is_promotion_blocked_for_capacity",
]

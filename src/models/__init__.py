from .base import BaseParadigmModel, ModelExplanation, ModelInterface
from .ensemble import EnsembleOutput, ModelEnsembler
from .deployment import (
    ModelSlots,
    PromotionGates,
    REASON_PROMOTION_GATES_PASSED,
    REASON_RISK_BREACH,
    REASON_ROLLBACK_TO_PRIOR_CHAMPION,
    REASON_ROBUSTNESS_FAILURE,
    REASON_SHADOW_UNDERPERFORMANCE,
    REASON_STABILITY_FAILURE,
    SlotEvent,
)
from .registry import MODEL_REGISTRY, ModelActivation, ModelActivationConfig, activated_models, create_model
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
    "REASON_SHADOW_UNDERPERFORMANCE",
    "REASON_STABILITY_FAILURE",
    "RobustnessThresholds",
    "build_robustness_scorecards",
]

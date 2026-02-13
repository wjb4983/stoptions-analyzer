from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from .base import ModelInterface
from .paradigms import (
    DispersionModel,
    EventDrivenModel,
    FactorNeutralCrossSectionalRankModel,
    MacroRegimeConditionedModel,
    MeanReversionModel,
    MetaLabelClassifierModel,
    MicrostructureImbalanceModel,
    MomentumModel,
    OptionsFlowDrivenModel,
    StatArbPairSpreadModel,
    TermStructureSlopeModel,
    VolatilityCarryModel,
)

MODEL_REGISTRY: dict[str, type[ModelInterface]] = {
    MomentumModel.name: MomentumModel,
    MeanReversionModel.name: MeanReversionModel,
    VolatilityCarryModel.name: VolatilityCarryModel,
    TermStructureSlopeModel.name: TermStructureSlopeModel,
    DispersionModel.name: DispersionModel,
    StatArbPairSpreadModel.name: StatArbPairSpreadModel,
    FactorNeutralCrossSectionalRankModel.name: FactorNeutralCrossSectionalRankModel,
    MacroRegimeConditionedModel.name: MacroRegimeConditionedModel,
    EventDrivenModel.name: EventDrivenModel,
    OptionsFlowDrivenModel.name: OptionsFlowDrivenModel,
    MicrostructureImbalanceModel.name: MicrostructureImbalanceModel,
    MetaLabelClassifierModel.name: MetaLabelClassifierModel,
}


@dataclass(frozen=True)
class ModelActivation:
    name: str
    weight: float = 1.0
    enabled: bool = True


@dataclass(frozen=True)
class ModelActivationConfig:
    paradigms: tuple[ModelActivation, ...]

    @staticmethod
    def from_dict(config: dict[str, Any]) -> "ModelActivationConfig":
        paradigms_cfg = config.get("paradigms", [])
        paradigms: list[ModelActivation] = []
        for entry in paradigms_cfg:
            paradigms.append(
                ModelActivation(
                    name=str(entry["name"]),
                    weight=float(entry.get("weight", 1.0)),
                    enabled=bool(entry.get("enabled", True)),
                )
            )
        return ModelActivationConfig(paradigms=tuple(paradigms))


def create_model(name: str) -> ModelInterface:
    model_cls = MODEL_REGISTRY.get(name)
    if model_cls is None:
        raise KeyError(f"Unknown model paradigm: {name}")
    return model_cls()


def activated_models(config: ModelActivationConfig) -> list[tuple[ModelInterface, float]]:
    active: list[tuple[ModelInterface, float]] = []
    for paradigm in config.paradigms:
        if not paradigm.enabled:
            continue
        active.append((create_model(paradigm.name), paradigm.weight))
    return active

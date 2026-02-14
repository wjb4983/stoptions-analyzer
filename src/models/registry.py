from __future__ import annotations

from dataclasses import dataclass, field
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
    OptionsDirectionalModel,
    OptionsVolatilityModel,
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
    OptionsDirectionalModel.name: OptionsDirectionalModel,
    OptionsVolatilityModel.name: OptionsVolatilityModel,
    MicrostructureImbalanceModel.name: MicrostructureImbalanceModel,
    MetaLabelClassifierModel.name: MetaLabelClassifierModel,
}

DEPLOYMENT_SLOTS = ("champion", "challenger", "candidate")


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


@dataclass(frozen=True)
class ModelRecord:
    model_id: str
    lineage: tuple[str, ...] = ()
    hyperparams: dict[str, Any] = field(default_factory=dict)
    calibration_version: str | None = None
    robustness_passed: bool | None = None
    stress_passed: bool | None = None
    deployment_slot: str | None = None

    def promotion_ready(self) -> bool:
        return bool(self.robustness_passed and self.stress_passed and self.calibration_version)


class ModelMetadataRegistry:
    """In-memory metadata registry for deployment and governance lookups."""

    def __init__(self) -> None:
        self._records: dict[str, ModelRecord] = {}
        self._slots: dict[str, str] = {}

    def register(
        self,
        model_id: str,
        *,
        lineage: tuple[str, ...] | list[str] = (),
        hyperparams: dict[str, Any] | None = None,
        calibration_version: str | None = None,
        robustness_passed: bool | None = None,
        stress_passed: bool | None = None,
        deployment_slot: str | None = None,
    ) -> ModelRecord:
        normalized_slot = self._normalize_slot(deployment_slot)
        normalized_lineage = tuple(str(item) for item in lineage)
        record = ModelRecord(
            model_id=str(model_id),
            lineage=normalized_lineage,
            hyperparams=dict(hyperparams or {}),
            calibration_version=calibration_version,
            robustness_passed=robustness_passed,
            stress_passed=stress_passed,
            deployment_slot=normalized_slot,
        )
        self._records[record.model_id] = record
        if normalized_slot:
            self._slots[normalized_slot] = record.model_id
        return record

    def get(self, model_id: str) -> ModelRecord:
        try:
            return self._records[str(model_id)]
        except KeyError as exc:
            raise KeyError(f"Unknown model id: {model_id}") from exc

    def slot_model(self, slot: str) -> ModelRecord | None:
        model_id = self._slots.get(self._normalize_slot(slot))
        if not model_id:
            return None
        return self._records.get(model_id)

    def assign_slot(self, model_id: str, slot: str) -> ModelRecord:
        record = self.get(model_id)
        normalized_slot = self._normalize_slot(slot)
        self._slots[normalized_slot] = record.model_id
        updated = ModelRecord(
            model_id=record.model_id,
            lineage=record.lineage,
            hyperparams=dict(record.hyperparams),
            calibration_version=record.calibration_version,
            robustness_passed=record.robustness_passed,
            stress_passed=record.stress_passed,
            deployment_slot=normalized_slot,
        )
        self._records[record.model_id] = updated
        return updated

    def promotion_candidate(self, target_slot: str = "champion") -> ModelRecord | None:
        if self._normalize_slot(target_slot) != "champion":
            return None
        challenger = self.slot_model("challenger")
        if challenger and challenger.promotion_ready():
            return challenger
        return None

    def list_records(self) -> tuple[ModelRecord, ...]:
        return tuple(self._records.values())

    @staticmethod
    def _normalize_slot(slot: str | None) -> str | None:
        if slot is None:
            return None
        normalized = str(slot).strip().lower()
        if normalized not in DEPLOYMENT_SLOTS:
            raise ValueError(f"Unknown deployment slot: {slot}")
        return normalized


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


MODEL_METADATA_REGISTRY = ModelMetadataRegistry()

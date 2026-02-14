from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from src.models.registry import MODEL_METADATA_REGISTRY, ModelMetadataRegistry, ModelRecord


@dataclass
class NextGenRegistry:
    """Adapter-friendly interface for model metadata and deployment promotion lookup."""

    backend: ModelMetadataRegistry = MODEL_METADATA_REGISTRY

    def register_model(
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
        return self.backend.register(
            model_id,
            lineage=lineage,
            hyperparams=hyperparams,
            calibration_version=calibration_version,
            robustness_passed=robustness_passed,
            stress_passed=stress_passed,
            deployment_slot=deployment_slot,
        )

    def model_lineage(self, model_id: str) -> tuple[str, ...]:
        return self.backend.get(model_id).lineage

    def model_hyperparams(self, model_id: str) -> dict[str, Any]:
        return dict(self.backend.get(model_id).hyperparams)

    def calibration_version(self, model_id: str) -> str | None:
        return self.backend.get(model_id).calibration_version

    def robustness_and_stress_status(self, model_id: str) -> tuple[bool | None, bool | None]:
        record = self.backend.get(model_id)
        return record.robustness_passed, record.stress_passed

    def assign_deployment_slot(self, model_id: str, slot: str) -> ModelRecord:
        return self.backend.assign_slot(model_id, slot)

    def lookup_slot(self, slot: str) -> ModelRecord | None:
        return self.backend.slot_model(slot)

    def lookup_promotion_candidate(self, target_slot: str = "champion") -> ModelRecord | None:
        return self.backend.promotion_candidate(target_slot)

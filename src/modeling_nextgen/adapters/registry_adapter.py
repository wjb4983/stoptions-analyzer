from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from models.base import ModelInterface
from models.registry import MODEL_REGISTRY, create_model

from ..core.config import NextGenModelingConfig
from ..core.registry import NextGenRegistry


@dataclass
class LegacyRegistryAdapter:
    """Opt-in bridge to legacy model registry without changing paradigm behavior."""

    config: NextGenModelingConfig
    metadata_registry: NextGenRegistry = NextGenRegistry()

    def enabled(self) -> bool:
        return self.config.enabled

    def available(self) -> tuple[str, ...]:
        return tuple(sorted(MODEL_REGISTRY.keys()))

    def create(self, name: str) -> ModelInterface:
        return create_model(name)

    def register_model_metadata(
        self,
        model_id: str,
        *,
        lineage: tuple[str, ...] | list[str] = (),
        hyperparams: dict[str, Any] | None = None,
        calibration_version: str | None = None,
        robustness_passed: bool | None = None,
        stress_passed: bool | None = None,
        deployment_slot: str | None = None,
    ) -> None:
        self.metadata_registry.register_model(
            model_id,
            lineage=lineage,
            hyperparams=hyperparams,
            calibration_version=calibration_version,
            robustness_passed=robustness_passed,
            stress_passed=stress_passed,
            deployment_slot=deployment_slot,
        )

    def lookup_for_slot_promotion(self, target_slot: str = "champion") -> str | None:
        candidate = self.metadata_registry.lookup_promotion_candidate(target_slot)
        if not candidate:
            return None
        return candidate.model_id

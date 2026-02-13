from __future__ import annotations

from dataclasses import dataclass

from src.models.base import ModelInterface
from src.models.registry import MODEL_REGISTRY, create_model

from ..core.config import NextGenModelingConfig


@dataclass
class LegacyRegistryAdapter:
    """Opt-in bridge to legacy model registry without changing paradigm behavior."""

    config: NextGenModelingConfig

    def enabled(self) -> bool:
        return self.config.enabled

    def available(self) -> tuple[str, ...]:
        return tuple(sorted(MODEL_REGISTRY.keys()))

    def create(self, name: str) -> ModelInterface:
        return create_model(name)

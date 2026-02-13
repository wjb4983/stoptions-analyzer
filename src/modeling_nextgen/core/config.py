from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(frozen=True)
class NextGenModelingConfig:
    enabled: bool = False
    feature_pipeline: tuple[str, ...] = ()
    model_pipeline: tuple[str, ...] = ()
    validation_pipeline: tuple[str, ...] = ()
    calibration_pipeline: tuple[str, ...] = ()
    serving_adapter: str = "local"
    extra: dict[str, Any] = field(default_factory=dict)

    @staticmethod
    def from_dict(raw: dict[str, Any]) -> "NextGenModelingConfig":
        return NextGenModelingConfig(
            enabled=bool(raw.get("enabled", False)),
            feature_pipeline=tuple(raw.get("feature_pipeline", ())),
            model_pipeline=tuple(raw.get("model_pipeline", ())),
            validation_pipeline=tuple(raw.get("validation_pipeline", ())),
            calibration_pipeline=tuple(raw.get("calibration_pipeline", ())),
            serving_adapter=str(raw.get("serving_adapter", "local")),
            extra=dict(raw.get("extra", {})),
        )

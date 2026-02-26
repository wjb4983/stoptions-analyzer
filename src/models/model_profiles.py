from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

from models.regime_catalog import list_models_for_leg

ModelProfileSource = Literal["catalog", "preset", "trained_artifact"]


@dataclass(frozen=True)
class ArtifactReference:
    run_id: str
    artifact_path: str
    checksum_metadata: dict[str, str]


@dataclass(frozen=True)
class ModelProfile:
    profile_id: str
    display_name: str
    source: ModelProfileSource
    base_model_id: str
    leg_family: str
    hyperparameters: dict[str, Any] | None = None
    architecture_spec: dict[str, Any] | None = None
    calibration_spec: dict[str, Any] | None = None
    event_process_spec: dict[str, Any] | None = None
    artifact_reference: ArtifactReference | None = None
    provenance: dict[str, str] | None = None

    @property
    def resolved_model_id(self) -> str:
        return self.base_model_id.strip()


@dataclass(frozen=True)
class ModelProfileRegistry:
    catalog_profiles: tuple[ModelProfile, ...]
    preset_profiles: tuple[ModelProfile, ...]
    trained_profiles: tuple[ModelProfile, ...]

    @property
    def profiles(self) -> tuple[ModelProfile, ...]:
        return (*self.catalog_profiles, *self.preset_profiles, *self.trained_profiles)


def build_model_profile_registry(
    *,
    leg_family: str,
    presets: dict[str, object] | None = None,
    training_runs: list[dict[str, object]] | None = None,
) -> ModelProfileRegistry:
    return ModelProfileRegistry(
        catalog_profiles=tuple(_catalog_profiles_for_leg(leg_family)),
        preset_profiles=tuple(_preset_profiles_for_leg(leg_family, presets or {})),
        trained_profiles=tuple(_trained_profiles_for_leg(leg_family, training_runs or [])),
    )


def _catalog_profiles_for_leg(leg_family: str) -> list[ModelProfile]:
    profiles: list[ModelProfile] = []
    for descriptor in list_models_for_leg(leg_family):
        profiles.append(
            ModelProfile(
                profile_id=f"catalog:{leg_family}:{descriptor.model_name}",
                display_name=descriptor.display_name,
                source="catalog",
                base_model_id=descriptor.model_name,
                leg_family=leg_family,
                hyperparameters=dict(descriptor.hyperparameter_template),
            )
        )
    return profiles


def _preset_profiles_for_leg(leg_family: str, presets: dict[str, object]) -> list[ModelProfile]:
    profiles: list[ModelProfile] = []
    for preset_name, payload in presets.items():
        if not isinstance(payload, dict):
            continue
        preset_leg_family = str(payload.get("leg_family", "")).strip()
        if preset_leg_family != leg_family:
            continue
        base_model_id = str(payload.get("base_model_id", "")).strip()
        if not base_model_id:
            continue
        profile_id = str(payload.get("profile_id", "")).strip() or f"preset:{leg_family}:{preset_name}"
        display_name = str(payload.get("display_name", "")).strip() or preset_name
        profiles.append(
            ModelProfile(
                profile_id=profile_id,
                display_name=display_name,
                source="preset",
                base_model_id=base_model_id,
                leg_family=leg_family,
                hyperparameters=_optional_dict(payload.get("hyperparameters")),
                architecture_spec=_optional_dict(payload.get("architecture_spec")),
                calibration_spec=_optional_dict(payload.get("calibration_spec")),
                event_process_spec=_optional_dict(payload.get("event_process_spec")),
                provenance=_optional_str_dict(payload.get("provenance"), allowed_keys=(
                    "paper_title",
                    "citation_key_or_url",
                    "task_fit",
                    "market_assumptions",
                )),
            )
        )
    return profiles


def _trained_profiles_for_leg(leg_family: str, training_runs: list[dict[str, object]]) -> list[ModelProfile]:
    profiles: list[ModelProfile] = []
    for run in reversed(training_runs):
        run_id = str(run.get("run_id", "")).strip()
        if not run_id:
            continue
        manifest = _load_run_manifest(run)
        if manifest is None:
            continue
        request = manifest.get("request")
        if not isinstance(request, dict):
            continue
        legs = request.get("legs")
        if not isinstance(legs, list):
            continue
        reproducibility = _load_reproducibility_payload(manifest)
        for index, leg in enumerate(legs):
            if not isinstance(leg, dict):
                continue
            if str(leg.get("model_type", "")).strip() != leg_family:
                continue
            resolved_model = str(leg.get("model_id", "")).strip() or str(leg.get("selected_model_id", "")).strip()
            if not resolved_model:
                continue
            leg_name = str(leg.get("name", f"Leg {index + 1}")).strip() or f"Leg {index + 1}"
            checksum_metadata = _checksums_for_leg(reproducibility, leg_name=leg_name, index=index)
            artifact_path = str(run.get("artifact_path", "")).strip()
            profiles.append(
                ModelProfile(
                    profile_id=f"trained_artifact:{run_id}:{index}",
                    display_name=f"{leg_name} ({resolved_model}) [{run_id[:10]}]",
                    source="trained_artifact",
                    base_model_id=resolved_model,
                    leg_family=leg_family,
                    hyperparameters=_optional_dict(leg.get("hyperparameters")),
                    architecture_spec=_optional_dict(leg.get("architecture_spec")),
                    calibration_spec=_optional_dict(leg.get("calibration_spec")),
                    event_process_spec=_optional_dict(leg.get("event_process_spec")),
                    artifact_reference=ArtifactReference(
                        run_id=run_id,
                        artifact_path=artifact_path,
                        checksum_metadata=checksum_metadata,
                    ),
                )
            )
    return profiles


def _optional_dict(value: object) -> dict[str, Any] | None:
    if not isinstance(value, dict):
        return None
    return dict(value)


def _optional_str_dict(value: object, *, allowed_keys: tuple[str, ...]) -> dict[str, str] | None:
    if not isinstance(value, dict):
        return None
    result: dict[str, str] = {}
    for key in allowed_keys:
        raw = value.get(key)
        text = str(raw).strip() if raw is not None else ""
        if text:
            result[key] = text
    return result or None


def _load_run_manifest(run: dict[str, object]) -> dict[str, object] | None:
    path_raw = str(run.get("artifact_path", "")).strip()
    if not path_raw:
        return None
    path = Path(path_raw)
    if not path.exists() or not path.is_file():
        return None
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    if not isinstance(payload, dict):
        return None
    return payload


def _load_reproducibility_payload(manifest: dict[str, object]) -> dict[str, object]:
    metadata = manifest.get("metadata")
    if not isinstance(metadata, dict):
        return {}
    reproducibility = metadata.get("reproducibility")
    if not isinstance(reproducibility, dict):
        return {}
    return reproducibility


def _checksums_for_leg(reproducibility: dict[str, object], *, leg_name: str, index: int) -> dict[str, str]:
    legs = reproducibility.get("legs")
    if not isinstance(legs, dict):
        return {}
    key = f"{index:02d}:{leg_name}"
    payload = legs.get(key)
    if not isinstance(payload, dict):
        return {}
    return {str(name): str(value) for name, value in payload.items() if isinstance(name, str) and isinstance(value, str)}


__all__ = [
    "ArtifactReference",
    "ModelProfile",
    "ModelProfileRegistry",
    "ModelProfileSource",
    "build_model_profile_registry",
]

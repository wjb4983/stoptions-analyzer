from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from typing import Any

ALLOWED_SAMPLERS = {"tpe", "cma-es", "random", "grid"}
LEGACY_SAMPLER_MIGRATIONS = {"bayesian": "tpe"}
WORKFLOW_PRESET_SCHEMA_VERSION = 2


@dataclass
class WorkflowPresetValidationResult:
    payload: dict[str, Any]
    warnings: list[str]


def validate_workflow_preset_payload(
    payload: Any,
    *,
    fallback_payload: dict[str, Any],
) -> WorkflowPresetValidationResult:
    warnings: list[str] = []
    fallback = deepcopy(fallback_payload)

    if not isinstance(payload, dict):
        warnings.append(
            f"Invalid payload at $: expected JSON object, got {type(payload).__name__}; using built-in defaults."
        )
        return WorkflowPresetValidationResult(payload=fallback, warnings=warnings)

    migrated, migration_warnings = migrate_workflow_preset_payload(payload, fallback_payload=fallback)
    warnings.extend(migration_warnings)

    presets = migrated.get("presets")
    if not isinstance(presets, dict) or not presets:
        warnings.append("Missing or invalid key at $.presets: expected non-empty object; using built-in defaults.")
        return WorkflowPresetValidationResult(payload=fallback, warnings=warnings)

    sanitized = deepcopy(migrated)
    sanitized_presets = sanitized.get("presets")
    assert isinstance(sanitized_presets, dict)

    fallback_presets = fallback.get("presets", {})
    fallback_default_name = str(fallback.get("default_preset", "balanced_baseline"))
    fallback_default = fallback_presets.get(fallback_default_name)
    if not isinstance(fallback_default, dict):
        fallback_default = next((v for v in fallback_presets.values() if isinstance(v, dict)), {})

    for preset_name, preset_value in list(sanitized_presets.items()):
        if not isinstance(preset_value, dict):
            warnings.append(f"Invalid key at $.presets.{preset_name}: expected object, got {type(preset_value).__name__}; dropped preset.")
            sanitized_presets.pop(preset_name, None)
            continue
        _sanitize_sampler(preset_name, preset_value, fallback_default, warnings)
        _sanitize_numeric_bounds(preset_name, preset_value, fallback_default, warnings)

    if not sanitized_presets:
        warnings.append("All presets were invalid after validation; using built-in defaults.")
        return WorkflowPresetValidationResult(payload=fallback, warnings=warnings)

    default_preset = str(sanitized.get("default_preset") or next(iter(sanitized_presets)))
    if default_preset not in sanitized_presets:
        replacement = next(iter(sanitized_presets))
        warnings.append(
            f"Invalid key at $.default_preset: '{default_preset}' is not present in $.presets; using '{replacement}'."
        )
        default_preset = replacement

    return WorkflowPresetValidationResult(
        payload={
            "schema_version": WORKFLOW_PRESET_SCHEMA_VERSION,
            "default_preset": default_preset,
            "presets": sanitized_presets,
        },
        warnings=warnings,
    )


def migrate_workflow_preset_payload(
    payload: dict[str, Any],
    *,
    fallback_payload: dict[str, Any],
) -> tuple[dict[str, Any], list[str]]:
    migrated = deepcopy(payload)
    warnings: list[str] = []

    raw_schema = migrated.get("schema_version", 1)
    try:
        schema_version = int(raw_schema)
    except (TypeError, ValueError):
        warnings.append(
            f"Invalid key at $.schema_version: expected integer, got {raw_schema!r}; treating as legacy schema 1."
        )
        schema_version = 1

    if "default_preset" not in migrated:
        fallback_default_name = str(fallback_payload.get("default_preset", "balanced_baseline"))
        migrated["default_preset"] = fallback_default_name
        warnings.append(
            f"Missing key at $.default_preset; using fallback '{fallback_default_name}'."
        )

    presets = migrated.get("presets")
    if not isinstance(presets, dict):
        return migrated, warnings

    fallback_presets = fallback_payload.get("presets", {})
    fallback_default_name = str(fallback_payload.get("default_preset", "balanced_baseline"))
    fallback_default = fallback_presets.get(fallback_default_name)
    if not isinstance(fallback_default, dict):
        fallback_default = next((v for v in fallback_presets.values() if isinstance(v, dict)), {})

    if schema_version < WORKFLOW_PRESET_SCHEMA_VERSION:
        for preset_name, preset_value in presets.items():
            if not isinstance(preset_value, dict):
                continue
            _migrate_missing_keys(
                preset_name=preset_name,
                preset=preset_value,
                fallback_default=fallback_default,
                warnings=warnings,
            )
        warnings.append(
            f"Migrated $.schema_version from {schema_version} to {WORKFLOW_PRESET_SCHEMA_VERSION}."
        )

    migrated["schema_version"] = WORKFLOW_PRESET_SCHEMA_VERSION
    return migrated, warnings


def _sanitize_sampler(
    preset_name: str,
    preset: dict[str, Any],
    fallback_default: dict[str, Any],
    warnings: list[str],
) -> None:
    optimization = preset.get("optimization")
    if not isinstance(optimization, dict):
        warnings.append(
            f"Invalid key at $.presets.{preset_name}.optimization: expected object, got {type(optimization).__name__}; using defaults where required."
        )
        return
    sampler = str(optimization.get("sampler", "")).strip().lower()
    if not sampler:
        return
    if sampler in LEGACY_SAMPLER_MIGRATIONS:
        migrated = LEGACY_SAMPLER_MIGRATIONS[sampler]
        optimization["sampler"] = migrated
        warnings.append(
            f"Migrated key at $.presets.{preset_name}.optimization.sampler from '{sampler}' to '{migrated}'."
        )
        return
    if sampler in ALLOWED_SAMPLERS:
        optimization["sampler"] = sampler
        return

    fallback_sampler = "tpe"
    default_opt = fallback_default.get("optimization") if isinstance(fallback_default, dict) else {}
    if isinstance(default_opt, dict) and str(default_opt.get("sampler", "")) in ALLOWED_SAMPLERS:
        fallback_sampler = str(default_opt["sampler"])
    optimization["sampler"] = fallback_sampler
    warnings.append(
        f"Invalid key at $.presets.{preset_name}.optimization.sampler: unsupported value '{sampler}'; using '{fallback_sampler}'."
    )


def _sanitize_numeric_bounds(
    preset_name: str,
    preset: dict[str, Any],
    fallback_default: dict[str, Any],
    warnings: list[str],
) -> None:
    walk_forward = preset.get("walk_forward")
    fallback_wf = fallback_default.get("walk_forward", {}) if isinstance(fallback_default, dict) else {}
    if isinstance(walk_forward, dict):
        _validate_number(
            walk_forward,
            fallback_wf,
            key="train_fraction",
            min_value=0.0,
            max_value=1.0,
            inclusive_min=False,
            preset_name=preset_name,
            section="walk_forward",
            warnings=warnings,
        )
        _validate_number(walk_forward, fallback_wf, key="validation_fraction", min_value=0.0, max_value=1.0, inclusive_min=False, preset_name=preset_name, section="walk_forward", warnings=warnings)
        _validate_number(walk_forward, fallback_wf, key="test_fraction", min_value=0.0, max_value=1.0, inclusive_min=False, preset_name=preset_name, section="walk_forward", warnings=warnings)
        _validate_number(walk_forward, fallback_wf, key="step_fraction", min_value=0.0, max_value=1.0, inclusive_min=False, preset_name=preset_name, section="walk_forward", warnings=warnings)
    elif walk_forward is not None:
        warnings.append(
            f"Invalid key at $.presets.{preset_name}.walk_forward: expected object, got {type(walk_forward).__name__}; using defaults where required."
        )

    stress = preset.get("stress_controls")
    fallback_stress = fallback_default.get("stress_controls", {}) if isinstance(fallback_default, dict) else {}
    if isinstance(stress, dict):
        _validate_number(stress, fallback_stress, key="historical_window_fraction", min_value=0.0, max_value=1.0, inclusive_min=False, preset_name=preset_name, section="stress_controls", warnings=warnings)
        _validate_number(stress, fallback_stress, key="historical_replay_window_bars", min_value=1.0, as_int=True, preset_name=preset_name, section="stress_controls", warnings=warnings)
        _validate_number(stress, fallback_stress, key="synthetic_jump_magnitude", min_value=0.0, max_value=1.0, inclusive_min=True, preset_name=preset_name, section="stress_controls", warnings=warnings)
        _validate_number(stress, fallback_stress, key="synthetic_jump_interval", min_value=1.0, as_int=True, preset_name=preset_name, section="stress_controls", warnings=warnings)
        _validate_number(stress, fallback_stress, key="synthetic_vol_cluster_multiplier", min_value=0.0, max_value=10.0, inclusive_min=False, preset_name=preset_name, section="stress_controls", warnings=warnings)
        _validate_number(stress, fallback_stress, key="overlay_spread_multiplier", min_value=0.0, max_value=10.0, inclusive_min=False, preset_name=preset_name, section="stress_controls", warnings=warnings)
        _validate_number(stress, fallback_stress, key="overlay_liquidity_multiplier", min_value=0.0, max_value=1.0, inclusive_min=True, preset_name=preset_name, section="stress_controls", warnings=warnings)
    elif stress is not None:
        warnings.append(
            f"Invalid key at $.presets.{preset_name}.stress_controls: expected object, got {type(stress).__name__}; using defaults where required."
        )


def _validate_number(
    container: dict[str, Any],
    fallback_container: dict[str, Any],
    *,
    key: str,
    min_value: float,
    max_value: float | None = None,
    inclusive_min: bool = True,
    as_int: bool = False,
    preset_name: str,
    section: str,
    warnings: list[str],
) -> None:
    if key not in container:
        return
    raw = container[key]
    try:
        value = float(raw)
    except (TypeError, ValueError):
        _replace_with_fallback(container, fallback_container, key)
        warnings.append(
            f"Invalid key at $.presets.{preset_name}.{section}.{key}: expected numeric value, got {raw!r}; using fallback."
        )
        return

    min_ok = value >= min_value if inclusive_min else value > min_value
    max_ok = True if max_value is None else value <= max_value
    if not (min_ok and max_ok):
        _replace_with_fallback(container, fallback_container, key)
        warnings.append(
            f"Invalid key at $.presets.{preset_name}.{section}.{key}: out-of-range value {raw!r}; using fallback."
        )
        return

    container[key] = int(round(value)) if as_int else float(value)


def _replace_with_fallback(container: dict[str, Any], fallback_container: dict[str, Any], key: str) -> None:
    if key in fallback_container:
        container[key] = fallback_container[key]
    else:
        container.pop(key, None)


def _migrate_missing_keys(
    *,
    preset_name: str,
    preset: dict[str, Any],
    fallback_default: dict[str, Any],
    warnings: list[str],
) -> None:
    for key in ("entry_signals", "exit_signals", "benchmark_selection"):
        if key not in preset and key in fallback_default:
            preset[key] = deepcopy(fallback_default[key])
            warnings.append(
                f"Missing key at $.presets.{preset_name}.{key}; backfilled from defaults."
            )

    _migrate_missing_section_keys(
        preset_name=preset_name,
        preset=preset,
        fallback_default=fallback_default,
        section="optimization",
        warnings=warnings,
    )
    _migrate_missing_section_keys(
        preset_name=preset_name,
        preset=preset,
        fallback_default=fallback_default,
        section="walk_forward",
        warnings=warnings,
    )
    _migrate_missing_section_keys(
        preset_name=preset_name,
        preset=preset,
        fallback_default=fallback_default,
        section="stress_controls",
        warnings=warnings,
    )


def _migrate_missing_section_keys(
    *,
    preset_name: str,
    preset: dict[str, Any],
    fallback_default: dict[str, Any],
    section: str,
    warnings: list[str],
) -> None:
    fallback_section = fallback_default.get(section)
    if not isinstance(fallback_section, dict):
        return
    current = preset.get(section)
    if current is None:
        preset[section] = deepcopy(fallback_section)
        warnings.append(f"Missing key at $.presets.{preset_name}.{section}; backfilled from defaults.")
        return
    if not isinstance(current, dict):
        return
    for key, fallback_value in fallback_section.items():
        if key not in current:
            current[key] = deepcopy(fallback_value)
            warnings.append(
                f"Missing key at $.presets.{preset_name}.{section}.{key}; backfilled from defaults."
            )

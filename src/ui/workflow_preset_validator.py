from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from typing import Any

ALLOWED_SAMPLERS = {"tpe", "cma-es", "random", "grid"}
LEGACY_SAMPLER_MIGRATIONS = {"bayesian": "tpe"}


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
        warnings.append("Preset payload is not a JSON object; using built-in defaults.")
        return WorkflowPresetValidationResult(payload=fallback, warnings=warnings)

    presets = payload.get("presets")
    if not isinstance(presets, dict) or not presets:
        warnings.append("Preset payload missing non-empty 'presets'; using built-in defaults.")
        return WorkflowPresetValidationResult(payload=fallback, warnings=warnings)

    sanitized = deepcopy(payload)
    sanitized_presets = sanitized.get("presets")
    assert isinstance(sanitized_presets, dict)

    fallback_presets = fallback.get("presets", {})
    fallback_default_name = str(fallback.get("default_preset", "balanced_baseline"))
    fallback_default = fallback_presets.get(fallback_default_name)
    if not isinstance(fallback_default, dict):
        fallback_default = next((v for v in fallback_presets.values() if isinstance(v, dict)), {})

    for preset_name, preset_value in list(sanitized_presets.items()):
        if not isinstance(preset_value, dict):
            warnings.append(f"Preset '{preset_name}' is not an object and was dropped.")
            sanitized_presets.pop(preset_name, None)
            continue
        _sanitize_sampler(sanitized_presets, preset_name, preset_value, fallback_default, warnings)
        _sanitize_numeric_bounds(sanitized_presets, preset_name, preset_value, fallback_default, warnings)

    if not sanitized_presets:
        warnings.append("All presets were invalid after validation; using built-in defaults.")
        return WorkflowPresetValidationResult(payload=fallback, warnings=warnings)

    default_preset = str(sanitized.get("default_preset") or next(iter(sanitized_presets)))
    if default_preset not in sanitized_presets:
        replacement = next(iter(sanitized_presets))
        warnings.append(
            f"default_preset '{default_preset}' is not defined; using '{replacement}'."
        )
        default_preset = replacement

    return WorkflowPresetValidationResult(
        payload={"default_preset": default_preset, "presets": sanitized_presets},
        warnings=warnings,
    )


def _sanitize_sampler(
    presets: dict[str, Any],
    preset_name: str,
    preset: dict[str, Any],
    fallback_default: dict[str, Any],
    warnings: list[str],
) -> None:
    optimization = preset.get("optimization")
    if not isinstance(optimization, dict):
        return
    sampler = str(optimization.get("sampler", "")).strip().lower()
    if not sampler:
        return
    if sampler in LEGACY_SAMPLER_MIGRATIONS:
        migrated = LEGACY_SAMPLER_MIGRATIONS[sampler]
        optimization["sampler"] = migrated
        warnings.append(
            f"Preset '{preset_name}' migrated optimization.sampler '{sampler}' -> '{migrated}'."
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
        f"Preset '{preset_name}' has unsupported optimization.sampler '{sampler}'; using '{fallback_sampler}'."
    )


def _sanitize_numeric_bounds(
    presets: dict[str, Any],
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
        warnings.append(f"Preset '{preset_name}' has non-numeric {section}.{key}; using default.")
        return

    min_ok = value >= min_value if inclusive_min else value > min_value
    max_ok = True if max_value is None else value <= max_value
    if not (min_ok and max_ok):
        _replace_with_fallback(container, fallback_container, key)
        warnings.append(f"Preset '{preset_name}' has out-of-range {section}.{key}={raw}; using default.")
        return

    container[key] = int(round(value)) if as_int else float(value)


def _replace_with_fallback(container: dict[str, Any], fallback_container: dict[str, Any], key: str) -> None:
    if key in fallback_container:
        container[key] = fallback_container[key]
    else:
        container.pop(key, None)

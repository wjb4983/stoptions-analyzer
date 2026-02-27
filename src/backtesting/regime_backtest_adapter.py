from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from backtesting.schema_contracts import (
    BACKTEST_HYDRATION_PAYLOAD_CONTRACT,
    EXPORT_BUNDLE_MANIFEST_CONTRACT,
    REGIME_TRAINING_MANIFEST_CONTRACT,
)

class RegimeBundleCompatibilityError(ValueError):
    """Raised when a regime bundle/manifest cannot be safely consumed."""


@dataclass(frozen=True)
class RegimeBacktestOption:
    option_id: str
    label: str
    source: str
    manifest_path: str


@dataclass(frozen=True)
class RegimeBacktestContract:
    option_id: str
    regime_name: str
    source: str
    manifest_path: str
    defaults: dict[str, object]
    execution_artifacts: dict[str, object]


def discover_regime_backtest_options(
    regime_training_runs: list[dict[str, object]] | None,
    *,
    regime_exports_root: str | Path = Path("data/regime_exports"),
) -> list[RegimeBacktestOption]:
    options: list[RegimeBacktestOption] = []
    seen_paths: set[str] = set()

    for run in regime_training_runs or []:
        if not isinstance(run, dict):
            continue
        run_id = str(run.get("run_id", "")).strip()
        artifact_path = str(run.get("artifact_path", "")).strip()
        if not run_id or not artifact_path:
            continue
        norm_path = str(Path(artifact_path))
        if norm_path in seen_paths:
            continue
        seen_paths.add(norm_path)
        regime_name = str(run.get("regime_name", "")).strip() or str(run.get("summary", "")).strip() or run_id
        options.append(
            RegimeBacktestOption(
                option_id=f"training:{run_id}",
                label=f"{regime_name} [{run_id}] (training run)",
                source="training_run",
                manifest_path=norm_path,
            )
        )

    exports_root = Path(regime_exports_root)
    if exports_root.exists():
        for manifest_path in sorted(exports_root.glob("*/bundle_manifest.json")):
            norm_path = str(manifest_path)
            if norm_path in seen_paths:
                continue
            try:
                payload = json.loads(manifest_path.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, OSError):
                continue
            if not isinstance(payload, dict):
                continue
            bundle_id = str(payload.get("bundle_id", manifest_path.parent.name)).strip() or manifest_path.parent.name
            run_id = str(payload.get("run_id", "")).strip() or "unknown"
            options.append(
                RegimeBacktestOption(
                    option_id=f"bundle:{bundle_id}",
                    label=f"{bundle_id} [{run_id}] (export bundle)",
                    source="bundle",
                    manifest_path=norm_path,
                )
            )
            seen_paths.add(norm_path)

    options.sort(key=lambda item: item.label.lower())
    return options


def load_regime_backtest_contract(option: RegimeBacktestOption) -> RegimeBacktestContract:
    payload = _read_manifest_payload(Path(option.manifest_path), source=option.source)
    training_payload = _normalize_training_manifest_payload(payload)
    request = training_payload.get("request", {}) if isinstance(training_payload.get("request"), dict) else {}
    training_window = request.get("training_window", {}) if isinstance(request.get("training_window"), dict) else {}
    risk_limits = request.get("risk_limits", {}) if isinstance(request.get("risk_limits"), dict) else {}
    training_data_settings = (
        request.get("training_data_settings", {})
        if isinstance(request.get("training_data_settings"), dict)
        else {}
    )

    scenario_settings = training_data_settings.get("scenario_settings", [])
    scenario_names = [
        str(item.get("name", "")).strip()
        for item in scenario_settings
        if isinstance(item, dict) and str(item.get("name", "")).strip()
    ]
    strategy = _strategy_from_model_choice(str(request.get("model_choice", "")))
    max_net = float(risk_limits.get("max_net_exposure", 0.5) or 0.5)
    defaults: dict[str, object] = {
        "strategy": strategy,
        "lookback_days": str(int(training_window.get("lookback_days", 90) or 90)),
        "skip_days": str(int(training_window.get("retrain_frequency_days", 5) or 5)),
        "portfolio_max_gross_exposure": f"{float(risk_limits.get('max_gross_exposure', 1.0) or 1.0):.2f}",
        "portfolio_max_net_exposure": f"{max_net:.2f}",
        "portfolio_min_net_exposure": f"{-abs(max_net):.2f}",
        "portfolio_max_symbol_weight": f"{float(risk_limits.get('max_position_weight', 0.10) or 0.10):.2f}",
        "portfolio_max_sector_weight": f"{float(risk_limits.get('max_sector_weight', 0.25) or 0.25):.2f}",
        "governance_min_stability_score": f"{float(risk_limits.get('confidence_min_assignment_confidence', 0.60) or 0.60):.2f}",
        "governance_expected_signal_agreement": f"{float(risk_limits.get('confidence_alert_confidence', 0.70) or 0.70):.2f}",
        "governance_max_signal_agreement_drift": f"{max(0.0, 1.0 - float(risk_limits.get('confidence_min_transition_confidence', 0.55) or 0.55)):.2f}",
        "stress_enable_historical_replay_regimes": bool(scenario_names),
        "selected_scenario_packs": ",".join(scenario_names),
    }
    defaults["hydration_schema_version"] = BACKTEST_HYDRATION_PAYLOAD_CONTRACT.current_version
    regime_name = str(request.get("regime_name", training_payload.get("run_id", option.option_id))).strip() or option.option_id
    execution_artifacts = _extract_execution_artifacts(training_payload)
    return RegimeBacktestContract(
        option_id=option.option_id,
        regime_name=regime_name,
        source=option.source,
        manifest_path=option.manifest_path,
        defaults=defaults,
        execution_artifacts=execution_artifacts,
    )


def _extract_execution_artifacts(training_payload: dict[str, Any]) -> dict[str, object]:
    metadata = training_payload.get("metadata", {}) if isinstance(training_payload.get("metadata"), dict) else {}
    artifact_paths = training_payload.get("artifact_paths", {}) if isinstance(training_payload.get("artifact_paths"), dict) else {}
    request = training_payload.get("request", {}) if isinstance(training_payload.get("request"), dict) else {}
    legs = request.get("legs", []) if isinstance(request.get("legs"), list) else []

    champion_by_leg = {
        str(leg_name): str(model_id)
        for leg_name, model_id in (metadata.get("champion_by_leg", {}) or {}).items()
        if str(leg_name).strip() and str(model_id).strip()
    }

    feature_expectations: dict[str, list[str]] = {}
    for raw_leg in legs:
        if not isinstance(raw_leg, dict):
            continue
        leg_name = str(raw_leg.get("name", "")).strip()
        if not leg_name:
            continue
        controls = raw_leg.get("controls", {}) if isinstance(raw_leg.get("controls"), dict) else {}
        feature_columns = controls.get("feature_columns")
        if isinstance(feature_columns, list):
            cleaned = [str(item).strip() for item in feature_columns if str(item).strip()]
            if cleaned:
                feature_expectations[leg_name] = cleaned

    model_paths = {
        key.removesuffix("_model_weights"): str(path)
        for key, path in artifact_paths.items()
        if isinstance(key, str) and key.endswith("_model_weights") and str(path).strip()
    }
    calibration_paths = {
        key.removesuffix("_calibration_object"): str(path)
        for key, path in artifact_paths.items()
        if isinstance(key, str) and key.endswith("_calibration_object") and str(path).strip()
    }

    return {
        "run_id": str(training_payload.get("run_id", "")).strip(),
        "manifest_schema_version": str(training_payload.get("manifest_schema_version", "")).strip(),
        "champion_model_ids": champion_by_leg,
        "model_paths": model_paths,
        "calibration_paths": calibration_paths,
        "feature_expectations": feature_expectations,
    }


def _read_manifest_payload(manifest_path: Path, *, source: str) -> dict[str, Any]:
    try:
        payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise RegimeBundleCompatibilityError(
            f"Manifest not found at '{manifest_path}'. Re-export the regime bundle or retrain the regime."
        ) from exc
    except json.JSONDecodeError as exc:
        raise RegimeBundleCompatibilityError(
            f"Manifest at '{manifest_path}' is invalid JSON (line {exc.lineno}, column {exc.colno})."
        ) from exc

    if not isinstance(payload, dict):
        raise RegimeBundleCompatibilityError(f"Manifest at '{manifest_path}' must be a JSON object.")

    if source != "bundle":
        _ensure_training_manifest_compatible(payload)
        return payload

    bundle_version = str(payload.get("manifest_schema_version", payload.get("bundle_version", ""))).strip()
    if not bundle_version:
        raise RegimeBundleCompatibilityError(
            "Bundle manifest missing 'bundle_version'. Re-export using the latest regime export workflow."
        )
    if not EXPORT_BUNDLE_MANIFEST_CONTRACT.is_compatible(bundle_version):
        raise RegimeBundleCompatibilityError(
            f"Bundle schema v{bundle_version} is not supported by this UI. "
            f"Supported version range: {EXPORT_BUNDLE_MANIFEST_CONTRACT.minimum_compatible_version}"
            f"-{EXPORT_BUNDLE_MANIFEST_CONTRACT.current_version}. "
            "Re-export with a compatible app version."
        )

    contents = payload.get("contents", {}) if isinstance(payload.get("contents"), dict) else {}
    training_manifest_path = str(contents.get("training_manifest", "")).strip()
    if not training_manifest_path:
        raise RegimeBundleCompatibilityError(
            "Bundle manifest missing contents.training_manifest. Re-export the bundle from a full training manifest."
        )
    return _read_manifest_payload(Path(training_manifest_path), source="training_run")


def _ensure_training_manifest_compatible(payload: dict[str, Any]) -> None:
    version = str(payload.get("manifest_schema_version", "")).strip()
    if version:
        if not REGIME_TRAINING_MANIFEST_CONTRACT.is_compatible(version):
            raise RegimeBundleCompatibilityError(
                f"Training manifest schema v{version} is not compatible with this loader. "
                f"Supported version range: {REGIME_TRAINING_MANIFEST_CONTRACT.minimum_compatible_version}"
                f"-{REGIME_TRAINING_MANIFEST_CONTRACT.current_version}."
            )
        return

    # N-1 adapter for legacy manifests (pre-contract-field). Treat request schema >=2 as 2.0.0.
    legacy_request = payload.get("request", {}) if isinstance(payload.get("request"), dict) else {}
    legacy_schema = int(legacy_request.get("schema_version", 0) or 0)
    if legacy_schema < 2:
        raise RegimeBundleCompatibilityError(
            "Legacy training manifest is too old for hydration. Expected request.schema_version >= 2."
        )


def _normalize_training_manifest_payload(payload: dict[str, Any]) -> dict[str, Any]:
    if payload.get("manifest_schema_version"):
        return payload
    normalized = dict(payload)
    normalized["manifest_schema_version"] = REGIME_TRAINING_MANIFEST_CONTRACT.minimum_compatible_version
    normalized.setdefault("metadata", {})
    return normalized


def _strategy_from_model_choice(model_choice: str) -> str:
    normalized = model_choice.strip().lower()
    if normalized in {"cross_sectional", "xsmom", "cross_sectional_momentum"}:
        return "xsmom"
    return "momentum"

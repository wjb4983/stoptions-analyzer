import json
from dataclasses import dataclass, field

from config import (
    DEFAULT_BACKTEST_SETTINGS,
    DEFAULT_GENERAL_ANALYSIS_SETTINGS,
    DEFAULT_REGIME_CONFIDENCE_THRESHOLDS,
    DEFAULT_REGIME_GLOBAL_RISK_LIMITS,
    DEFAULT_REMOTE_EXECUTION_SETTINGS,
    DEFAULT_REGIME_TRAINING_WINDOW,
    DEFAULT_REGIME_TRAINING_DATA_SETTINGS,
    STATE_PATH,
    merged_remote_execution_settings,
)

REGIME_DEFINITION_SCHEMA_VERSION = 3


_LEG_TYPE_ALIASES: dict[str, str] = {
    "Trend Following": "timeseries_momentum",
    "Mean Reversion": "cheap_vol_buying",
    "Volatility Breakout": "volatility_risk_premium_selling",
    "Regime Change": "regime_change_detection",
    "Volatility Clustering": "volatility_clustering",
    "IV/EV Spread": "iv_ev_spread_term_structure",
    "Event Intensity": "self_exciting_event_intensity",
    "Vol Surface": "vol_surface_calibration",
    "Cross-Asset Macro": "cross_asset_macro_conditioned",
    "Meta-Label Ensemble": "meta_label_regime_ensemble",
    "iv_ev_spread": "iv_ev_spread_term_structure",
}


def _normalize_leg_model_type(raw_model_type: object) -> str:
    model_type = str(raw_model_type or "").strip()
    return _LEG_TYPE_ALIASES.get(model_type, model_type)


def _baseline_regime_definitions() -> dict[str, dict[str, object]]:
    return {
        "baseline": {
            "label": "Baseline",
            "schema_version": REGIME_DEFINITION_SCHEMA_VERSION,
            "global_risk_limits": dict(DEFAULT_REGIME_GLOBAL_RISK_LIMITS),
            "training_window": dict(DEFAULT_REGIME_TRAINING_WINDOW),
            "confidence_thresholds": dict(DEFAULT_REGIME_CONFIDENCE_THRESHOLDS),
            "training_data_settings": dict(DEFAULT_REGIME_TRAINING_DATA_SETTINGS),
            "legs": [],
        }
    }


def _migrate_leg_payload(raw_leg: object) -> dict[str, object] | None:
    if not isinstance(raw_leg, dict):
        return None
    migrated = dict(raw_leg)
    migrated["model_type"] = _normalize_leg_model_type(migrated.get("model_type"))
    selected_model_id = str(migrated.get("selected_model_id", "")).strip()
    model_id = str(migrated.get("model_id", "")).strip() or selected_model_id
    migrated["model_id"] = model_id
    migrated["selected_model_id"] = selected_model_id or model_id
    hyperparams = migrated.get("hyperparameters")
    if not isinstance(hyperparams, dict):
        migrated["hyperparameters"] = {}
    for key in ("architecture_spec", "calibration_spec", "event_process_spec"):
        value = migrated.get(key)
        migrated[key] = value if isinstance(value, dict) else None
    return migrated




def _migrate_remote_synced_runs(payload: object) -> dict[str, str]:
    if not isinstance(payload, dict):
        return {}
    migrated: dict[str, str] = {}
    for key, value in payload.items():
        if isinstance(key, str) and isinstance(value, str) and key.strip() and value.strip():
            migrated[key.strip()] = value.strip()
    return migrated

def _migrate_active_jobs(payload: object) -> dict[str, dict[str, object]]:
    if not isinstance(payload, dict):
        return {}
    migrated: dict[str, dict[str, object]] = {}
    for raw_job_id, raw_meta in payload.items():
        if not isinstance(raw_job_id, str) or not raw_job_id.strip() or not isinstance(raw_meta, dict):
            continue
        job_id = raw_job_id.strip()
        migrated[job_id] = {
            "job_id": job_id,
            "job_type": str(raw_meta.get("job_type", "")).strip() or "unknown",
            "source_page": str(raw_meta.get("source_page", "")).strip() or "unknown",
            "status": str(raw_meta.get("status", "queued")).strip() or "queued",
            "submitted_at": str(raw_meta.get("submitted_at", "")).strip() or None,
            "started_at": str(raw_meta.get("started_at", "")).strip() or None,
            "ended_at": str(raw_meta.get("ended_at", "")).strip() or None,
            "server_hostname": str(raw_meta.get("server_hostname", "")).strip() or "unknown",
            "artifact_sync_status": str(raw_meta.get("artifact_sync_status", "not_started")).strip() or "not_started",
            "poll_interval_seconds": float(raw_meta.get("poll_interval_seconds", 0.8) or 0.8),
            "transport_retries": int(raw_meta.get("transport_retries", 0) or 0),
            "max_transport_retries": int(raw_meta.get("max_transport_retries", 4) or 4),
            "retryable_transport_failure": bool(raw_meta.get("retryable_transport_failure", False)),
            "error_kind": str(raw_meta.get("error_kind", "")).strip() or None,
            "error_message": str(raw_meta.get("error_message", "")).strip() or None,
        }
    return migrated


def _migrate_regime_definitions(payload: object) -> dict[str, dict[str, object]]:
    if not isinstance(payload, dict) or not payload:
        return _baseline_regime_definitions()

    migrated_defs: dict[str, dict[str, object]] = {}
    for regime_id, raw_definition in payload.items():
        if not isinstance(regime_id, str) or not isinstance(raw_definition, dict):
            continue

        definition = dict(raw_definition)
        definition["schema_version"] = REGIME_DEFINITION_SCHEMA_VERSION
        if not isinstance(definition.get("training_data_settings"), dict):
            definition["training_data_settings"] = dict(DEFAULT_REGIME_TRAINING_DATA_SETTINGS)

        raw_legs = definition.get("legs")
        migrated_legs: list[dict[str, object]] = []
        if isinstance(raw_legs, list):
            for raw_leg in raw_legs:
                migrated_leg = _migrate_leg_payload(raw_leg)
                if migrated_leg is not None:
                    migrated_legs.append(migrated_leg)
        definition["legs"] = migrated_legs

        migrated_defs[regime_id] = definition

    return migrated_defs or _baseline_regime_definitions()


@dataclass
class AppState:
    tickers: list[str] = field(default_factory=list)
    selected_ticker: str | None = None
    analysis_mode: str = "Stock Analysis"
    option_strategy: str = "Naked Call"
    general_analysis_settings: dict[str, object] = field(
        default_factory=lambda: dict(DEFAULT_GENERAL_ANALYSIS_SETTINGS)
    )
    backtest_settings: dict[str, object] = field(
        default_factory=lambda: dict(DEFAULT_BACKTEST_SETTINGS)
    )
    backtest_templates: dict[str, dict[str, object]] = field(default_factory=dict)
    regime_definitions: dict[str, dict[str, object]] = field(default_factory=_baseline_regime_definitions)
    regime_training_runs: list[dict[str, object]] = field(default_factory=list)
    active_regime_id: str | None = None
    remote_synced_runs: dict[str, str] = field(default_factory=dict)
    remote_execution_settings: dict[str, object] = field(
        default_factory=lambda: dict(DEFAULT_REMOTE_EXECUTION_SETTINGS)
    )
    active_jobs: dict[str, dict[str, object]] = field(default_factory=dict)

    def save(self) -> None:
        payload = {
            "tickers": self.tickers,
            "selected_ticker": self.selected_ticker,
            "analysis_mode": self.analysis_mode,
            "option_strategy": self.option_strategy,
            "general_analysis_settings": self.general_analysis_settings,
            "backtest_settings": self.backtest_settings,
            "backtest_templates": self.backtest_templates,
            "regime_definitions": _migrate_regime_definitions(self.regime_definitions),
            "regime_training_runs": self.regime_training_runs,
            "active_regime_id": self.active_regime_id,
            "remote_synced_runs": _migrate_remote_synced_runs(self.remote_synced_runs),
            "remote_execution_settings": merged_remote_execution_settings(self.remote_execution_settings),
            "active_jobs": _migrate_active_jobs(self.active_jobs),
        }
        STATE_PATH.write_text(json.dumps(payload, indent=2))

    @classmethod
    def load(cls) -> "AppState":
        if not STATE_PATH.exists():
            return cls()
        try:
            payload = json.loads(STATE_PATH.read_text())
        except json.JSONDecodeError:
            return cls()
        return cls(
            tickers=payload.get("tickers", []),
            selected_ticker=payload.get("selected_ticker"),
            analysis_mode=payload.get("analysis_mode", payload.get("analysis_type", "Stock Analysis")),
            option_strategy=payload.get("option_strategy", "Naked Call"),
            general_analysis_settings=payload.get(
                "general_analysis_settings", dict(DEFAULT_GENERAL_ANALYSIS_SETTINGS)
            ),
            backtest_settings=payload.get(
                "backtest_settings", dict(DEFAULT_BACKTEST_SETTINGS)
            ),
            backtest_templates=payload.get("backtest_templates", {}),
            regime_definitions=_migrate_regime_definitions(payload.get("regime_definitions")),
            regime_training_runs=payload.get("regime_training_runs", []),
            active_regime_id=payload.get("active_regime_id"),
            remote_synced_runs=_migrate_remote_synced_runs(payload.get("remote_synced_runs", {})),
            remote_execution_settings=merged_remote_execution_settings(payload.get("remote_execution_settings", {})),
            active_jobs=_migrate_active_jobs(payload.get("active_jobs", {})),
        )

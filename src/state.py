import json
from dataclasses import dataclass, field

from config import (
    DEFAULT_BACKTEST_SETTINGS,
    DEFAULT_GENERAL_ANALYSIS_SETTINGS,
    DEFAULT_REGIME_CONFIDENCE_THRESHOLDS,
    DEFAULT_REGIME_GLOBAL_RISK_LIMITS,
    DEFAULT_REGIME_TRAINING_WINDOW,
    STATE_PATH,
)

REGIME_DEFINITION_SCHEMA_VERSION = 2


def _baseline_regime_definitions() -> dict[str, dict[str, object]]:
    return {
        "baseline": {
            "label": "Baseline",
            "schema_version": REGIME_DEFINITION_SCHEMA_VERSION,
            "global_risk_limits": dict(DEFAULT_REGIME_GLOBAL_RISK_LIMITS),
            "training_window": dict(DEFAULT_REGIME_TRAINING_WINDOW),
            "confidence_thresholds": dict(DEFAULT_REGIME_CONFIDENCE_THRESHOLDS),
            "legs": [],
        }
    }


def _migrate_leg_payload(raw_leg: object) -> dict[str, object] | None:
    if not isinstance(raw_leg, dict):
        return None
    migrated = dict(raw_leg)
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


def _migrate_regime_definitions(payload: object) -> dict[str, dict[str, object]]:
    if not isinstance(payload, dict) or not payload:
        return _baseline_regime_definitions()

    migrated_defs: dict[str, dict[str, object]] = {}
    for regime_id, raw_definition in payload.items():
        if not isinstance(regime_id, str) or not isinstance(raw_definition, dict):
            continue

        definition = dict(raw_definition)
        definition["schema_version"] = REGIME_DEFINITION_SCHEMA_VERSION

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
        )

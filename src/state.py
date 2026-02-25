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
    regime_definitions: dict[str, dict[str, object]] = field(
        default_factory=lambda: {
            "baseline": {
                "label": "Baseline",
                "global_risk_limits": dict(DEFAULT_REGIME_GLOBAL_RISK_LIMITS),
                "training_window": dict(DEFAULT_REGIME_TRAINING_WINDOW),
                "confidence_thresholds": dict(DEFAULT_REGIME_CONFIDENCE_THRESHOLDS),
            }
        }
    )
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
            "regime_definitions": self.regime_definitions,
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
            regime_definitions=payload.get(
                "regime_definitions",
                {
                    "baseline": {
                        "label": "Baseline",
                        "global_risk_limits": dict(DEFAULT_REGIME_GLOBAL_RISK_LIMITS),
                        "training_window": dict(DEFAULT_REGIME_TRAINING_WINDOW),
                        "confidence_thresholds": dict(DEFAULT_REGIME_CONFIDENCE_THRESHOLDS),
                    }
                },
            ),
            regime_training_runs=payload.get("regime_training_runs", []),
            active_regime_id=payload.get("active_regime_id"),
        )

import json
from dataclasses import dataclass, field

from config import DEFAULT_BACKTEST_SETTINGS, DEFAULT_GENERAL_ANALYSIS_SETTINGS, STATE_PATH


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

    def save(self) -> None:
        payload = {
            "tickers": self.tickers,
            "selected_ticker": self.selected_ticker,
            "analysis_mode": self.analysis_mode,
            "option_strategy": self.option_strategy,
            "general_analysis_settings": self.general_analysis_settings,
            "backtest_settings": self.backtest_settings,
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
        )

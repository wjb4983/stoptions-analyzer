from __future__ import annotations

import json
import threading
from pathlib import Path
from datetime import date, datetime, timedelta, timezone
import tkinter as tk
from tkinter import messagebox, ttk
from typing import TYPE_CHECKING, Any

from backtesting.regime_training_pipeline import (
    RegimeLegTrainingConfig,
    RegimeTrainingRequest,
    RegimeTrainingResult,
    execute_regime_training_pipeline,
)
from backtesting.cache_runner import run_backtest_cache
from data_access.cache_audit import audit_universe_history
from backtesting.regime_export_service import export_regime_training_bundle
from models.regime_catalog import ModelDescriptor, get_model_descriptor, list_models_for_leg
from ui.calibration_spec_designer_page import CalibrationSpecDesignerPage
from ui.event_process_designer_page import EventProcessDesignerPage
from ui.regime_mapping import to_regime_leg_spec
from ui.neural_network_designer_page import NeuralNetworkDesignerPage
from config import (
    DEFAULT_REGIME_CONFIDENCE_THRESHOLDS,
    DEFAULT_REGIME_GLOBAL_RISK_LIMITS,
    DEFAULT_REGIME_TRAINING_WINDOW,
    DEFAULT_REGIME_TRAINING_DATA_SETTINGS,
    BACKTEST_CACHE_DIR,
)

if TYPE_CHECKING:
    from main import StoptionsApp


LEG_CONTROL_GROUPS: dict[str, dict[str, object]] = {
    "Trend Following": {
        "signal_parameters": [
            {
                "key": "lookback_days",
                "label": "Lookback (days)",
                "default": 90,
                "min": 20,
                "max": 252,
                "tooltip": "Window used to estimate trend direction.",
            },
            {
                "key": "entry_zscore",
                "label": "Entry z-score",
                "default": 1.2,
                "min": 0.5,
                "max": 3.0,
                "tooltip": "Signal strength threshold to open trades.",
            },
        ],
        "sizing_risk_caps": [
            {
                "key": "max_position_pct",
                "label": "Max position %",
                "default": 0.08,
                "min": 0.01,
                "max": 0.2,
                "tooltip": "Maximum single-name capital allocation.",
            },
            {
                "key": "max_drawdown_stop",
                "label": "Max drawdown stop",
                "default": 0.12,
                "min": 0.03,
                "max": 0.3,
                "tooltip": "Circuit breaker threshold for this leg.",
            },
        ],
        "turnover_liquidity_assumptions": [
            {
                "key": "turnover_limit",
                "label": "Turnover limit",
                "default": 0.25,
                "min": 0.05,
                "max": 1.0,
                "tooltip": "Maximum daily turnover as fraction of NAV.",
            },
            {
                "key": "slippage_bps",
                "label": "Slippage (bps)",
                "default": 6,
                "min": 1,
                "max": 50,
                "tooltip": "Estimated one-way implementation cost.",
            },
        ],
        "confidence_thresholds": [
            {
                "key": "model_confidence_min",
                "label": "Model confidence min",
                "default": 0.65,
                "min": 0.5,
                "max": 0.95,
                "tooltip": "Minimum confidence needed for active exposure.",
            },
            {
                "key": "regime_stability_min",
                "label": "Regime stability min",
                "default": 0.55,
                "min": 0.3,
                "max": 0.95,
                "tooltip": "Minimum persistence confidence before trading.",
            },
        ],
    },
    "Mean Reversion": {
        "signal_parameters": [
            {
                "key": "lookback_days",
                "label": "Lookback (days)",
                "default": 30,
                "min": 10,
                "max": 126,
                "tooltip": "Window used to estimate short-term dislocations.",
            },
            {
                "key": "entry_zscore",
                "label": "Entry z-score",
                "default": 2.0,
                "min": 0.8,
                "max": 4.0,
                "tooltip": "Distance from fair-value required to enter.",
            },
        ],
        "sizing_risk_caps": [
            {
                "key": "max_position_pct",
                "label": "Max position %",
                "default": 0.05,
                "min": 0.01,
                "max": 0.15,
                "tooltip": "Maximum gross exposure on one instrument.",
            },
            {
                "key": "max_drawdown_stop",
                "label": "Max drawdown stop",
                "default": 0.09,
                "min": 0.03,
                "max": 0.25,
                "tooltip": "Risk stop for adverse momentum continuation.",
            },
        ],
        "turnover_liquidity_assumptions": [
            {
                "key": "turnover_limit",
                "label": "Turnover limit",
                "default": 0.45,
                "min": 0.1,
                "max": 1.5,
                "tooltip": "Expected rebalancing intensity.",
            },
            {
                "key": "slippage_bps",
                "label": "Slippage (bps)",
                "default": 10,
                "min": 1,
                "max": 60,
                "tooltip": "Execution cost estimate for fast-turn strategies.",
            },
        ],
        "confidence_thresholds": [
            {
                "key": "model_confidence_min",
                "label": "Model confidence min",
                "default": 0.6,
                "min": 0.4,
                "max": 0.95,
                "tooltip": "Confidence floor before entering reversion trades.",
            },
            {
                "key": "regime_stability_min",
                "label": "Regime stability min",
                "default": 0.5,
                "min": 0.3,
                "max": 0.95,
                "tooltip": "Controls sensitivity to changing market regimes.",
            },
        ],
    },
    "Volatility Breakout": {
        "signal_parameters": [
            {
                "key": "lookback_days",
                "label": "Lookback (days)",
                "default": 45,
                "min": 10,
                "max": 180,
                "tooltip": "Volatility regime estimation horizon.",
            },
            {
                "key": "entry_zscore",
                "label": "Entry z-score",
                "default": 1.6,
                "min": 0.5,
                "max": 4.0,
                "tooltip": "Vol breakout threshold before activating hedges.",
            },
        ],
        "sizing_risk_caps": [
            {
                "key": "max_position_pct",
                "label": "Max position %",
                "default": 0.06,
                "min": 0.01,
                "max": 0.2,
                "tooltip": "Cap for long-vol or defensive allocations.",
            },
            {
                "key": "max_drawdown_stop",
                "label": "Max drawdown stop",
                "default": 0.1,
                "min": 0.03,
                "max": 0.35,
                "tooltip": "Stop level when hedges fail to offset losses.",
            },
        ],
        "turnover_liquidity_assumptions": [
            {
                "key": "turnover_limit",
                "label": "Turnover limit",
                "default": 0.35,
                "min": 0.05,
                "max": 1.2,
                "tooltip": "Rotation expected around volatility shocks.",
            },
            {
                "key": "slippage_bps",
                "label": "Slippage (bps)",
                "default": 12,
                "min": 1,
                "max": 80,
                "tooltip": "Execution drag assumption during stress windows.",
            },
        ],
        "confidence_thresholds": [
            {
                "key": "model_confidence_min",
                "label": "Model confidence min",
                "default": 0.7,
                "min": 0.5,
                "max": 0.98,
                "tooltip": "Higher floor reduces false positive breakout calls.",
            },
            {
                "key": "regime_stability_min",
                "label": "Regime stability min",
                "default": 0.6,
                "min": 0.3,
                "max": 0.95,
                "tooltip": "Persistence filter for volatile environments.",
            },
        ],
    },
    "Regime Change": {
        "signal_parameters": [
            {
                "key": "lookback_days",
                "label": "Lookback (days)",
                "default": 60,
                "min": 20,
                "max": 252,
                "tooltip": "Window used for regime-shift detection signals.",
            },
            {
                "key": "detection_threshold",
                "label": "Detection threshold",
                "default": 1.6,
                "min": 0.5,
                "max": 4.0,
                "tooltip": "Score threshold used to flag likely regime transitions.",
            },
        ],
        "sizing_risk_caps": [
            {
                "key": "max_position_pct",
                "label": "Max position %",
                "default": 0.05,
                "min": 0.01,
                "max": 0.15,
                "tooltip": "Cap for transition trades during state changes.",
            },
            {
                "key": "max_drawdown_stop",
                "label": "Max drawdown stop",
                "default": 0.08,
                "min": 0.03,
                "max": 0.25,
                "tooltip": "Risk stop for false-positive transition calls.",
            },
        ],
        "turnover_liquidity_assumptions": [
            {
                "key": "turnover_limit",
                "label": "Turnover limit",
                "default": 0.3,
                "min": 0.05,
                "max": 1.2,
                "tooltip": "Expected position rotation around transition windows.",
            },
            {
                "key": "slippage_bps",
                "label": "Slippage (bps)",
                "default": 10,
                "min": 1,
                "max": 60,
                "tooltip": "Execution cost estimate when regimes reprice quickly.",
            },
        ],
        "confidence_thresholds": [
            {
                "key": "model_confidence_min",
                "label": "Model confidence min",
                "default": 0.62,
                "min": 0.4,
                "max": 0.98,
                "tooltip": "Confidence floor required to activate state-change trades.",
            },
            {
                "key": "regime_stability_min",
                "label": "Regime stability min",
                "default": 0.5,
                "min": 0.3,
                "max": 0.95,
                "tooltip": "Persistence filter before switching state exposure.",
            },
        ],
    },
    "Volatility Clustering": {
        "signal_parameters": [
            {"key": "lookback_days", "label": "Lookback (days)", "default": 63, "min": 10, "max": 252, "tooltip": "Vol clustering estimation window."},
            {"key": "entry_zscore", "label": "Cluster activation threshold", "default": 1.4, "min": 0.4, "max": 4.0, "tooltip": "Threshold to activate clustering overlays."},
        ],
        "sizing_risk_caps": [
            {"key": "max_position_pct", "label": "Max position %", "default": 0.05, "min": 0.01, "max": 0.15, "tooltip": "Cap for clustering overlays."},
            {"key": "max_drawdown_stop", "label": "Max drawdown stop", "default": 0.1, "min": 0.03, "max": 0.25, "tooltip": "Stop for clustering false positives."},
        ],
        "turnover_liquidity_assumptions": [
            {"key": "turnover_limit", "label": "Turnover limit", "default": 0.3, "min": 0.05, "max": 1.2, "tooltip": "Turnover expectation in clustered volatility."},
            {"key": "slippage_bps", "label": "Slippage (bps)", "default": 11, "min": 1, "max": 70, "tooltip": "Execution drag during clustered volatility."},
        ],
        "confidence_thresholds": [
            {"key": "model_confidence_min", "label": "Model confidence min", "default": 0.64, "min": 0.4, "max": 0.98, "tooltip": "Confidence floor for clustering models."},
            {"key": "regime_stability_min", "label": "Regime stability min", "default": 0.56, "min": 0.3, "max": 0.95, "tooltip": "Persistence filter in high-vol clusters."},
        ],
    },
    "IV/EV Spread": {
        "signal_parameters": [
            {"key": "lookback_days", "label": "Lookback (days)", "default": 45, "min": 10, "max": 180, "tooltip": "Window for IV/EV term spread estimation."},
            {"key": "entry_zscore", "label": "Spread trigger", "default": 1.3, "min": 0.4, "max": 4.0, "tooltip": "IV/EV spread trigger threshold."},
        ],
        "sizing_risk_caps": [
            {"key": "max_position_pct", "label": "Max position %", "default": 0.05, "min": 0.01, "max": 0.15, "tooltip": "Cap for term-structure carry positions."},
            {"key": "max_drawdown_stop", "label": "Max drawdown stop", "default": 0.09, "min": 0.03, "max": 0.25, "tooltip": "Risk stop for spread blowouts."},
        ],
        "turnover_liquidity_assumptions": [
            {"key": "turnover_limit", "label": "Turnover limit", "default": 0.27, "min": 0.05, "max": 1.0, "tooltip": "Expected turnover for term spread rotations."},
            {"key": "slippage_bps", "label": "Slippage (bps)", "default": 8, "min": 1, "max": 60, "tooltip": "Execution drag in spread rebalance windows."},
        ],
        "confidence_thresholds": [
            {"key": "model_confidence_min", "label": "Model confidence min", "default": 0.6, "min": 0.4, "max": 0.98, "tooltip": "Confidence floor before deploying spread legs."},
            {"key": "regime_stability_min", "label": "Regime stability min", "default": 0.52, "min": 0.3, "max": 0.95, "tooltip": "Persistence filter for curve shape regimes."},
        ],
    },
    "Event Intensity": {
        "signal_parameters": [
            {"key": "lookback_days", "label": "Lookback (days)", "default": 15, "min": 5, "max": 90, "tooltip": "Window used to estimate event intensity."},
            {"key": "entry_zscore", "label": "Intensity trigger", "default": 1.4, "min": 0.4, "max": 4.0, "tooltip": "Event-intensity trigger for activation."},
        ],
        "sizing_risk_caps": [
            {"key": "max_position_pct", "label": "Max position %", "default": 0.03, "min": 0.005, "max": 0.12, "tooltip": "Cap for event-driven exposure."},
            {"key": "max_drawdown_stop", "label": "Max drawdown stop", "default": 0.07, "min": 0.02, "max": 0.2, "tooltip": "Stop for event clustering overshoots."},
        ],
        "turnover_liquidity_assumptions": [
            {"key": "turnover_limit", "label": "Turnover limit", "default": 0.5, "min": 0.1, "max": 2.0, "tooltip": "Expected turnover around bursts of events."},
            {"key": "slippage_bps", "label": "Slippage (bps)", "default": 14, "min": 1, "max": 90, "tooltip": "Execution drag around high-intensity windows."},
        ],
        "confidence_thresholds": [
            {"key": "model_confidence_min", "label": "Model confidence min", "default": 0.58, "min": 0.4, "max": 0.98, "tooltip": "Confidence floor for event-intensity activation."},
            {"key": "regime_stability_min", "label": "Regime stability min", "default": 0.48, "min": 0.3, "max": 0.95, "tooltip": "Persistence filter for event regime state."},
        ],
    },
    "Vol Surface": {
        "signal_parameters": [
            {"key": "lookback_days", "label": "Lookback (days)", "default": 30, "min": 10, "max": 180, "tooltip": "Window for surface calibration stability checks."},
            {"key": "entry_zscore", "label": "Calibration stress trigger", "default": 1.2, "min": 0.4, "max": 4.0, "tooltip": "Surface stress threshold to activate model."},
        ],
        "sizing_risk_caps": [
            {"key": "max_position_pct", "label": "Max position %", "default": 0.04, "min": 0.01, "max": 0.12, "tooltip": "Cap for surface-calibrated exposure."},
            {"key": "max_drawdown_stop", "label": "Max drawdown stop", "default": 0.08, "min": 0.03, "max": 0.2, "tooltip": "Risk stop for calibration drift."},
        ],
        "turnover_liquidity_assumptions": [
            {"key": "turnover_limit", "label": "Turnover limit", "default": 0.26, "min": 0.05, "max": 1.0, "tooltip": "Expected rebalance cadence from surface recalibration."},
            {"key": "slippage_bps", "label": "Slippage (bps)", "default": 9, "min": 1, "max": 60, "tooltip": "Execution drag estimate for surface-driven trades."},
        ],
        "confidence_thresholds": [
            {"key": "model_confidence_min", "label": "Model confidence min", "default": 0.62, "min": 0.4, "max": 0.98, "tooltip": "Confidence floor for calibrated surfaces."},
            {"key": "regime_stability_min", "label": "Regime stability min", "default": 0.54, "min": 0.3, "max": 0.95, "tooltip": "Stability filter for surface regime consistency."},
        ],
    },
    "Cross-Asset Macro": {
        "signal_parameters": [
            {
                "key": "lookback_days",
                "label": "Lookback (days)",
                "default": 84,
                "min": 20,
                "max": 252,
                "tooltip": "Macro conditioning lookback horizon.",
            },
            {
                "key": "entry_zscore",
                "label": "Macro shock threshold",
                "default": 1.35,
                "min": 0.4,
                "max": 4.0,
                "tooltip": "Cross-asset/macro trigger threshold.",
            },
        ],
        "sizing_risk_caps": [
            {
                "key": "max_position_pct",
                "label": "Max position %",
                "default": 0.04,
                "min": 0.01,
                "max": 0.15,
                "tooltip": "Cap for macro-conditioned exposure.",
            },
            {
                "key": "max_drawdown_stop",
                "label": "Max drawdown stop",
                "default": 0.08,
                "min": 0.03,
                "max": 0.25,
                "tooltip": "Risk stop for cross-asset dislocations.",
            },
        ],
        "turnover_liquidity_assumptions": [
            {
                "key": "turnover_limit",
                "label": "Turnover limit",
                "default": 0.28,
                "min": 0.05,
                "max": 1.0,
                "tooltip": "Expected turnover for macro rotation.",
            },
            {
                "key": "slippage_bps",
                "label": "Slippage (bps)",
                "default": 9,
                "min": 1,
                "max": 60,
                "tooltip": "Execution drag in stressed macro windows.",
            },
        ],
        "confidence_thresholds": [
            {
                "key": "model_confidence_min",
                "label": "Model confidence min",
                "default": 0.58,
                "min": 0.4,
                "max": 0.98,
                "tooltip": "Confidence floor for macro-conditioned calls.",
            },
            {
                "key": "regime_stability_min",
                "label": "Regime stability min",
                "default": 0.52,
                "min": 0.3,
                "max": 0.95,
                "tooltip": "Persistence filter across asset regimes.",
            },
        ],
    },
    "Meta-Label Ensemble": {
        "signal_parameters": [
            {
                "key": "lookback_days",
                "label": "Lookback (days)",
                "default": 63,
                "min": 20,
                "max": 252,
                "tooltip": "Window used to form ensemble meta-labels.",
            },
            {
                "key": "entry_zscore",
                "label": "Ensemble gating threshold",
                "default": 1.25,
                "min": 0.4,
                "max": 4.0,
                "tooltip": "Threshold for activating stacked learners.",
            },
        ],
        "sizing_risk_caps": [
            {
                "key": "max_position_pct",
                "label": "Max position %",
                "default": 0.04,
                "min": 0.01,
                "max": 0.15,
                "tooltip": "Exposure cap for ensemble signals.",
            },
            {
                "key": "max_drawdown_stop",
                "label": "Max drawdown stop",
                "default": 0.09,
                "min": 0.03,
                "max": 0.25,
                "tooltip": "Circuit breaker for ensemble drift.",
            },
        ],
        "turnover_liquidity_assumptions": [
            {
                "key": "turnover_limit",
                "label": "Turnover limit",
                "default": 0.32,
                "min": 0.05,
                "max": 1.2,
                "tooltip": "Expected rotation from stacked model votes.",
            },
            {
                "key": "slippage_bps",
                "label": "Slippage (bps)",
                "default": 8,
                "min": 1,
                "max": 60,
                "tooltip": "Execution cost estimate for ensemble routing.",
            },
        ],
        "confidence_thresholds": [
            {
                "key": "model_confidence_min",
                "label": "Model confidence min",
                "default": 0.66,
                "min": 0.4,
                "max": 0.98,
                "tooltip": "Meta-label confidence floor before trade activation.",
            },
            {
                "key": "regime_stability_min",
                "label": "Regime stability min",
                "default": 0.57,
                "min": 0.3,
                "max": 0.95,
                "tooltip": "Persistence requirement for ensemble consensus.",
            },
        ],
    },
}

GROUP_TITLES = {
    "signal_parameters": "Signal parameters",
    "sizing_risk_caps": "Sizing / risk caps",
    "turnover_liquidity_assumptions": "Turnover / liquidity assumptions",
    "confidence_thresholds": "Confidence thresholds",
}

STATUS_COLORS = {"green": "#1b8f3a", "yellow": "#b07d00", "red": "#b12704"}
TRAINING_MODE_CHOICES = {
    "single_model": "Single model",
    "ensemble": "Ensemble",
    "auto_model_search": "Auto-model-search",
}

QUICK_PRESETS: list[dict[str, object]] = [
    {
        "name": "Balanced",
        "description": "Steady defaults for broad market conditions.",
        "controls": {
            "lookback_days": 63,
            "entry_zscore": 1.4,
            "max_position_pct": 0.06,
            "max_drawdown_stop": 0.1,
            "turnover_limit": 0.3,
            "slippage_bps": 8,
            "model_confidence_min": 0.65,
            "regime_stability_min": 0.55,
        },
    },
    {
        "name": "Fast React",
        "description": "Higher sensitivity and turnover for tactical workflows.",
        "controls": {
            "lookback_days": 30,
            "entry_zscore": 1.1,
            "max_position_pct": 0.05,
            "max_drawdown_stop": 0.09,
            "turnover_limit": 0.5,
            "slippage_bps": 12,
            "model_confidence_min": 0.6,
            "regime_stability_min": 0.5,
        },
    },
    {
        "name": "Defensive",
        "description": "Lower risk and stricter confidence for stress periods.",
        "controls": {
            "lookback_days": 90,
            "entry_zscore": 1.8,
            "max_position_pct": 0.04,
            "max_drawdown_stop": 0.08,
            "turnover_limit": 0.2,
            "slippage_bps": 7,
            "model_confidence_min": 0.75,
            "regime_stability_min": 0.65,
        },
    },
]


class CreateRegimePage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller

        if not self.controller.state.regime_definitions:
            self.controller.state.regime_definitions = {
                "baseline": {
                    "label": "Baseline",
                    "description": "Default regime profile",
                }
            }
        if self.controller.state.active_regime_id is None:
            self.controller.state.active_regime_id = next(iter(self.controller.state.regime_definitions), None)

        self.regime_legs: list[dict[str, object]] = [self._build_default_leg("Trend Following")]
        self.training_mode: str = "auto_model_search"
        self.selected_leg_index = 0
        self.leg_control_vars: dict[str, tk.Variable] = {}
        self.hyperparameter_vars: dict[str, tk.Variable] = {}
        self._model_selection_by_label: dict[str, ModelDescriptor] = {}
        self._is_training = False
        self._nn_designer_window: NeuralNetworkDesignerPage | None = None
        self._calibration_designer_window: CalibrationSpecDesignerPage | None = None
        self._event_process_designer_window: EventProcessDesignerPage | None = None
        self._training_preview_warnings: list[str] = []

        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        ttk.Label(self, text="Create Regime", font=("Arial", 18, "bold")).grid(row=0, column=0, sticky="w", padx=12, pady=(10, 6))

        self.main_pane = ttk.Panedwindow(self, orient=tk.VERTICAL)
        self.main_pane.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 10))

        top_pane = ttk.Panedwindow(self.main_pane, orient=tk.HORIZONTAL)
        self.main_pane.add(top_pane, weight=5)

        self.legs_panel = ttk.Frame(top_pane, padding=8)
        self.config_panel = ttk.Frame(top_pane, padding=12)
        top_pane.add(self.legs_panel, weight=2)
        top_pane.add(self.config_panel, weight=7)

        self.bottom_panel = ttk.Frame(self.main_pane, padding=8)
        self.main_pane.add(self.bottom_panel, weight=2)

        self._build_legs_panel()
        self._build_config_panel()
        self._build_summary_panel()
        self._build_bottom_panel()

        self._load_editor_state_from_definition()

        self._refresh_legs_list()
        self._load_selected_leg_into_form()
        self._update_validation_and_actions()

    def _theme_background_color(self) -> str:
        """Return a Tk-compatible background color for labels embedded in ttk frames."""
        style = ttk.Style()
        background = style.lookup("TFrame", "background")
        if background:
            return str(background)
        parent_bg = self.master.cget("bg") if self.master is not None else ""
        return str(parent_bg or "#f0f0f0")

    def _build_default_leg(self, leg_type: str) -> dict[str, object]:
        controls = {}
        schema = LEG_CONTROL_GROUPS[leg_type]
        for group in schema.values():
            for control in group:
                controls[control["key"]] = control["default"]
        selected_model_id, hyperparameters = self._default_model_config_for_leg_type(leg_type)
        return {
            "name": f"{leg_type} leg",
            "model_type": leg_type,
            "controls": controls,
            "model_id": selected_model_id,
            "selected_model_id": selected_model_id,
            "hyperparameters": hyperparameters,
            "architecture_spec": None,
            "calibration_spec": None,
            "event_process_spec": None,
            "nn_custom_presets": {},
        }

    def _default_model_config_for_leg_type(self, leg_type: str) -> tuple[str, dict[str, Any]]:
        leg = {
            "name": f"{leg_type} leg",
            "model_type": leg_type,
            "controls": {},
        }
        mapped = to_regime_leg_spec(leg)
        descriptors = list_models_for_leg(mapped.leg_spec.leg_family)
        if not descriptors:
            return "", {}
        descriptor = descriptors[0]
        return descriptor.model_name, dict(descriptor.hyperparameter_template)

    def _load_editor_state_from_definition(self) -> None:
        definition = self._active_regime_definition()
        raw_legs = definition.get("legs") if isinstance(definition.get("legs"), list) else None
        if raw_legs:
            self.regime_legs = [self._normalize_leg_payload(item) for item in raw_legs]
        else:
            self.regime_legs = [self._build_default_leg("Trend Following")]
        self.training_mode = str(definition.get("training_mode", "auto_model_search"))
        if self.training_mode not in TRAINING_MODE_CHOICES:
            self.training_mode = "auto_model_search"
        self.selected_leg_index = max(0, min(self.selected_leg_index, len(self.regime_legs) - 1))

    def _normalize_leg_payload(self, raw_leg: object) -> dict[str, object]:
        if not isinstance(raw_leg, dict):
            return self._build_default_leg("Trend Following")
        leg_type = str(raw_leg.get("model_type", "Trend Following"))
        if leg_type not in LEG_CONTROL_GROUPS:
            leg_type = "Trend Following"
        leg = self._build_default_leg(leg_type)
        leg["name"] = str(raw_leg.get("name", leg["name"]))
        controls = raw_leg.get("controls")
        if isinstance(controls, dict):
            for key, value in controls.items():
                if key in leg["controls"]:
                    try:
                        leg["controls"][key] = float(value)
                    except (TypeError, ValueError):
                        pass
        model_id = raw_leg.get("model_id")
        if isinstance(model_id, str):
            leg["model_id"] = model_id
        selected_model_id = raw_leg.get("selected_model_id")
        if isinstance(selected_model_id, str):
            leg["selected_model_id"] = selected_model_id
        if not isinstance(leg.get("model_id"), str) or not str(leg.get("model_id", "")).strip():
            leg["model_id"] = str(leg.get("selected_model_id", ""))
        hyperparameters = raw_leg.get("hyperparameters")
        if isinstance(hyperparameters, dict):
            leg["hyperparameters"] = dict(hyperparameters)
        for optional_key in ("architecture_spec", "calibration_spec", "event_process_spec"):
            raw_value = raw_leg.get(optional_key)
            if raw_value is None or isinstance(raw_value, dict):
                leg[optional_key] = raw_value
        self._ensure_leg_model_defaults(leg)
        return leg

    def _persist_editor_state(self) -> None:
        controller = getattr(self, "controller", None)
        if controller is None or not hasattr(controller, "state"):
            return
        regime_id = controller.state.active_regime_id
        if not regime_id:
            return
        definition = controller.state.regime_definitions.get(regime_id)
        if not isinstance(definition, dict):
            return
        definition["legs"] = [self._serialize_leg(leg) for leg in self.regime_legs]
        definition["training_mode"] = getattr(self, "training_mode", "auto_model_search")
        controller.persist_state()

    @staticmethod
    def _serialize_leg(leg: dict[str, object]) -> dict[str, object]:
        return {
            "name": str(leg.get("name", "Unnamed leg")),
            "model_type": str(leg.get("model_type", "Trend Following")),
            "controls": dict(leg.get("controls", {})),
            "model_id": str(leg.get("model_id", leg.get("selected_model_id", ""))),
            "selected_model_id": str(leg.get("selected_model_id", "")),
            "hyperparameters": dict(leg.get("hyperparameters", {})),
            "architecture_spec": leg.get("architecture_spec"),
            "calibration_spec": leg.get("calibration_spec"),
            "event_process_spec": leg.get("event_process_spec"),
            "nn_custom_presets": leg.get("nn_custom_presets", {}),
        }

    def _allowed_models_for_leg(self, leg: dict[str, object]) -> list[ModelDescriptor]:
        mapped_leg = to_regime_leg_spec(leg)
        return list_models_for_leg(mapped_leg.leg_spec.leg_family)

    def _selected_model_descriptor(self, leg: dict[str, object]) -> ModelDescriptor | None:
        selected_model_id = str(leg.get("selected_model_id", "")).strip()
        if not selected_model_id:
            return None
        mapped_leg = to_regime_leg_spec(leg)
        return get_model_descriptor(mapped_leg.leg_spec.leg_family, selected_model_id)

    @staticmethod
    def _spec_requirements_for_descriptor(descriptor: ModelDescriptor | None) -> dict[str, str]:
        requirements = {
            "architecture_spec": "unsupported",
            "calibration_spec": "unsupported",
            "event_process_spec": "unsupported",
        }
        if descriptor is None:
            return requirements

        tags = descriptor.capability_tags
        requirements["architecture_spec"] = "required" if "needs_architecture_spec" in tags else "optional" if "supports_architecture_spec" in tags else "unsupported"
        requirements["calibration_spec"] = "required" if "needs_calibration_spec" in tags else "optional" if "supports_calibration_spec" in tags else "unsupported"
        requirements["event_process_spec"] = "required" if "needs_event_process_spec" in tags else "optional" if "supports_event_process_spec" in tags else "unsupported"
        return requirements

    def _required_spec_blockers(self, leg: dict[str, object]) -> list[str]:
        try:
            descriptor = self._selected_model_descriptor(leg)
        except ValueError:
            return []
        requirements = self._spec_requirements_for_descriptor(descriptor)
        label_map = {
            "architecture_spec": "Architecture spec",
            "calibration_spec": "Calibration spec",
            "event_process_spec": "Event process spec",
        }
        blockers: list[str] = []
        for key, requirement in requirements.items():
            if requirement == "required" and not isinstance(leg.get(key), dict):
                blockers.append(label_map[key])
        return blockers

    def _ensure_leg_model_defaults(self, leg: dict[str, object]) -> None:
        descriptors = self._allowed_models_for_leg(leg)
        if not descriptors:
            leg["selected_model_id"] = ""
            leg["hyperparameters"] = {}
            return
        selected = str(leg.get("selected_model_id", "")).strip()
        descriptor_by_model = {item.model_name: item for item in descriptors}
        if selected not in descriptor_by_model:
            selected = descriptors[0].model_name
            leg["selected_model_id"] = selected
        leg["model_id"] = selected
        selected_descriptor = descriptor_by_model[selected]
        current_hyperparams = leg.get("hyperparameters")
        if not isinstance(current_hyperparams, dict):
            current_hyperparams = {}
        merged_hyperparams = dict(selected_descriptor.hyperparameter_template)
        merged_hyperparams.update(current_hyperparams)
        leg["hyperparameters"] = merged_hyperparams

    def _build_legs_panel(self) -> None:
        ttk.Label(self.legs_panel, text="Regime legs", font=("Arial", 12, "bold")).pack(anchor="w")
        self.legs_listbox = tk.Listbox(self.legs_panel, height=12, exportselection=False)
        self.legs_listbox.pack(fill="both", expand=True, pady=6)
        self.legs_listbox.bind("<<ListboxSelect>>", self._on_leg_selected)

        controls = ttk.Frame(self.legs_panel)
        controls.pack(fill="x")
        ttk.Button(controls, text="Add", command=self.add_leg).grid(row=0, column=0, padx=2)
        ttk.Button(controls, text="Remove", command=self.remove_selected_leg).grid(row=0, column=1, padx=2)
        ttk.Button(controls, text="↑", width=3, command=lambda: self.move_selected_leg(-1)).grid(row=0, column=2, padx=2)
        ttk.Button(controls, text="↓", width=3, command=lambda: self.move_selected_leg(1)).grid(row=0, column=3, padx=2)

    def _build_config_panel(self) -> None:
        header = ttk.Frame(self.config_panel)
        header.pack(fill="x")
        ttk.Label(header, text="Selected leg configuration", font=("Arial", 12, "bold")).pack(anchor="w")
        ttk.Label(
            header,
            text="Tune essentials first, then move into advanced controls when needed.",
            foreground="#555555",
        ).pack(anchor="w", pady=(2, 8))

        preset_row = ttk.Frame(header)
        preset_row.pack(fill="x", pady=(0, 8))
        ttk.Label(preset_row, text="Quick presets").pack(side="left", padx=(0, 8))
        for preset in QUICK_PRESETS:
            button = ttk.Button(
                preset_row,
                text=str(preset["name"]),
                command=lambda payload=preset: self._apply_quick_preset(payload),
                takefocus=True,
            )
            button.pack(side="left", padx=(0, 6))
            self._attach_tooltip(button, str(preset["description"]))

        self.leg_type_var = tk.StringVar(value="Trend Following")
        self.leg_type_combo = ttk.Combobox(
            header,
            textvariable=self.leg_type_var,
            values=list(LEG_CONTROL_GROUPS),
            state="readonly",
            width=24,
            takefocus=True,
        )
        self.leg_type_combo.pack(anchor="w", pady=(0, 0))
        self.leg_type_combo.bind("<<ComboboxSelected>>", self._on_leg_type_selected)

        self.form_notebook = ttk.Notebook(self.config_panel, takefocus=True)
        self.form_notebook.pack(fill="both", expand=True, pady=(10, 0))

        self.basics_tab = ttk.Frame(self.form_notebook, padding=8)
        self.advanced_tab = ttk.Frame(self.form_notebook, padding=8)
        self.validation_tab = ttk.Frame(self.form_notebook, padding=8)
        self.form_notebook.add(self.basics_tab, text="Basics")
        self.form_notebook.add(self.advanced_tab, text="Advanced")
        self.form_notebook.add(self.validation_tab, text="Validation & summary")

        self.basics_form_container = ttk.Frame(self.basics_tab)
        self.basics_form_container.pack(fill="both", expand=True)
        self.advanced_form_container = ttk.Frame(self.advanced_tab)
        self.advanced_form_container.pack(fill="both", expand=True)

    def _build_summary_panel(self) -> None:
        ttk.Label(self.validation_tab, text="Validation summary", font=("Arial", 12, "bold")).pack(anchor="w")
        ttk.Label(
            self.validation_tab,
            text="Use badges and recommendations to verify readiness before running training.",
            foreground="#555555",
        ).pack(anchor="w", pady=(2, 8))

        self.validation_badge_vars = {
            "data_sufficiency": tk.StringVar(),
            "overfit_risk": tk.StringVar(),
            "execution_realism": tk.StringVar(),
        }
        self.validation_badges: dict[str, ttk.Label] = {}

        badge_frame = ttk.Frame(self.validation_tab)
        badge_frame.pack(fill="x", pady=(6, 8))
        for idx, (key, var) in enumerate(self.validation_badge_vars.items()):
            label = ttk.Label(badge_frame, textvariable=var)
            label.grid(row=idx, column=0, sticky="w", pady=1)
            self.validation_badges[key] = label

        self.risk_summary_var = tk.StringVar()
        ttk.Label(self.validation_tab, textvariable=self.risk_summary_var, justify="left", wraplength=560).pack(fill="x", pady=(4, 8))

        self.pros_cons_text = tk.Text(self.validation_tab, height=14, wrap="word")
        self.pros_cons_text.pack(fill="both", expand=True)
        self.pros_cons_text.configure(state="disabled")

    def _attach_tooltip(self, widget: tk.Widget, text: str) -> None:
        tooltip_window: tk.Toplevel | None = None

        def show_tooltip(_event: object) -> None:
            nonlocal tooltip_window
            if tooltip_window is not None:
                return
            tooltip_window = tk.Toplevel(self)
            tooltip_window.wm_overrideredirect(True)
            x = widget.winfo_rootx() + 14
            y = widget.winfo_rooty() + 14
            tooltip_window.wm_geometry(f"+{x}+{y}")
            tk.Label(
                tooltip_window,
                text=text,
                bg="#fffbe8",
                relief="solid",
                borderwidth=1,
                padx=6,
                pady=4,
                wraplength=360,
                justify="left",
            ).pack()

        def hide_tooltip(_event: object) -> None:
            nonlocal tooltip_window
            if tooltip_window is not None:
                tooltip_window.destroy()
                tooltip_window = None

        widget.bind("<Enter>", show_tooltip)
        widget.bind("<Leave>", hide_tooltip)

    def _build_control_row(
        self,
        section: ttk.LabelFrame,
        *,
        row: int,
        control: dict[str, object],
        value: object,
        width: int,
    ) -> None:
        ttk.Label(section, text=str(control["label"])).grid(row=row, column=0, sticky="w", padx=(0, 6), pady=2)
        var = tk.StringVar(value=str(value))
        entry = ttk.Entry(section, textvariable=var, width=width, takefocus=True)
        entry.grid(row=row, column=1, sticky="w", padx=(0, 4), pady=2)
        info = ttk.Label(section, text="ⓘ", foreground="#4b6584")
        info.grid(row=row, column=2, sticky="w", pady=2)
        self._attach_tooltip(info, str(control["tooltip"]))
        entry.bind("<FocusOut>", lambda _e, key=str(control["key"]): self._on_control_edited(key))
        self.leg_control_vars[str(control["key"])] = var

    def _apply_quick_preset(self, preset: dict[str, object]) -> None:
        leg = self._selected_leg()
        controls = leg.get("controls")
        if not isinstance(controls, dict):
            controls = {}
            leg["controls"] = controls
        preset_controls = preset.get("controls")
        if isinstance(preset_controls, dict):
            controls.update(preset_controls)
        self._persist_editor_state()
        self._load_selected_leg_into_form()

    def _build_bottom_panel(self) -> None:
        action_row = ttk.Frame(self.bottom_panel)
        action_row.pack(fill="x")
        ttk.Label(action_row, text="Train / export", font=("Arial", 12, "bold")).pack(side="left")

        training_mode_frame = ttk.Frame(action_row)
        training_mode_frame.pack(side="left", padx=(12, 8))
        ttk.Label(training_mode_frame, text="Training mode").pack(side="left")
        self.training_mode_var = tk.StringVar(value=TRAINING_MODE_CHOICES.get(self.training_mode, "Auto-model-search"))
        self.training_mode_combo = ttk.Combobox(
            training_mode_frame,
            textvariable=self.training_mode_var,
            values=list(TRAINING_MODE_CHOICES.values()),
            state="readonly",
            width=20,
        )
        self.training_mode_combo.pack(side="left", padx=(6, 0))
        self.training_mode_combo.bind("<<ComboboxSelected>>", self._on_training_mode_selected)

        self.train_button = ttk.Button(action_row, text="Train", command=self._run_train)
        self.train_button.pack(side="left", padx=(12, 4))
        self.export_button = ttk.Button(action_row, text="Export", command=self._run_export)
        self.export_button.pack(side="left", padx=4)

        self.validation_message_var = tk.StringVar()
        ttk.Label(self.bottom_panel, textvariable=self.validation_message_var, wraplength=760).pack(anchor="w", pady=(6, 4))

        ttk.Label(self.bottom_panel, text="Training execution preview", font=("Arial", 10, "bold")).pack(anchor="w", pady=(4, 2))
        self.training_preview_text = tk.Text(self.bottom_panel, height=9, wrap="word")
        self.training_preview_text.pack(fill="x", expand=False, pady=(0, 6))
        self.training_preview_text.configure(state="disabled")

        ttk.Label(self.bottom_panel, text="Run logs", font=("Arial", 10, "bold")).pack(anchor="w")
        self.run_logs = tk.Text(self.bottom_panel, height=7, wrap="word")
        self.run_logs.pack(fill="both", expand=True, pady=(2, 0))

    def _selected_leg(self) -> dict[str, object]:
        if not self.regime_legs:
            self.regime_legs.append(self._build_default_leg("Trend Following"))
            self.selected_leg_index = 0
        if self.selected_leg_index is None:
            self.selected_leg_index = 0
        self.selected_leg_index = max(0, min(self.selected_leg_index, len(self.regime_legs) - 1))
        return self.regime_legs[self.selected_leg_index]

    def _refresh_legs_list(self) -> None:
        if not hasattr(self, "legs_listbox"):
            return
        self.legs_listbox.delete(0, tk.END)
        for idx, leg in enumerate(self.regime_legs, start=1):
            self.legs_listbox.insert(tk.END, f"{idx}. {leg['name']} ({leg['model_type']})")
        if self.regime_legs:
            self.legs_listbox.selection_set(self.selected_leg_index)

    def _on_leg_selected(self, _event=None) -> None:
        if not hasattr(self, "legs_listbox"):
            return
        selection = self.legs_listbox.curselection()
        if not selection:
            return
        self.selected_leg_index = selection[0]
        self._load_selected_leg_into_form()

    def add_leg(self) -> None:
        self.regime_legs.append(self._build_default_leg("Trend Following"))
        self.selected_leg_index = len(self.regime_legs) - 1
        self._persist_editor_state()
        self._refresh_legs_list()
        self._load_selected_leg_into_form()

    def remove_selected_leg(self) -> None:
        if len(self.regime_legs) <= 1:
            return
        self.regime_legs.pop(self.selected_leg_index)
        self.selected_leg_index = max(0, self.selected_leg_index - 1)
        self._persist_editor_state()
        self._refresh_legs_list()
        self._load_selected_leg_into_form()

    def move_selected_leg(self, offset: int) -> None:
        target = self.selected_leg_index + offset
        if target < 0 or target >= len(self.regime_legs):
            return
        self.regime_legs[self.selected_leg_index], self.regime_legs[target] = (
            self.regime_legs[target],
            self.regime_legs[self.selected_leg_index],
        )
        self.selected_leg_index = target
        self._persist_editor_state()
        self._refresh_legs_list()

    def _on_leg_type_selected(self, _event=None) -> None:
        self._apply_leg_type(self.leg_type_var.get())

    def _apply_leg_type(self, leg_type: str) -> None:
        leg = self._selected_leg()
        leg["model_type"] = leg_type
        leg["name"] = f"{leg_type} leg"
        leg["controls"] = {
            control["key"]: control["default"]
            for group in LEG_CONTROL_GROUPS[leg_type].values()
            for control in group
        }
        selected_model_id, hyperparameters = self._default_model_config_for_leg_type(leg_type)
        leg["selected_model_id"] = selected_model_id
        leg["model_id"] = selected_model_id
        leg["hyperparameters"] = hyperparameters
        self._persist_editor_state()
        self._refresh_legs_list()
        self._load_selected_leg_into_form()

    def _load_selected_leg_into_form(self) -> None:
        leg = self._selected_leg()
        self._ensure_leg_model_defaults(leg)
        if hasattr(self, "leg_type_var"):
            self.leg_type_var.set(str(leg["model_type"]))
        if hasattr(self, "training_mode_var"):
            self.training_mode_var.set(TRAINING_MODE_CHOICES.get(self.training_mode, "Auto-model-search"))
        if not hasattr(self, "basics_form_container"):
            self._update_validation_and_actions()
            return

        for widget in self.basics_form_container.winfo_children():
            widget.destroy()
        for widget in self.advanced_form_container.winfo_children():
            widget.destroy()

        selector_section = ttk.LabelFrame(self.basics_form_container, text="Model & profile", padding=8)
        selector_section.pack(fill="x", pady=(0, 8))
        ttk.Label(
            selector_section,
            text="Pick the leg model and required specs.",
            foreground="#555555",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))

        ttk.Label(selector_section, text="Model").grid(row=1, column=0, sticky="w")
        model_descriptors = self._allowed_models_for_leg(leg)
        self._model_selection_by_label = {
            f"{descriptor.display_name} ({descriptor.model_name})": descriptor
            for descriptor in model_descriptors
        }
        if not self._model_selection_by_label:
            self._model_selection_by_label = {"No models available": ModelDescriptor("", "", {})}
        selected_model_id = str(leg.get("selected_model_id", ""))
        default_label = next(
            (
                label
                for label, descriptor in self._model_selection_by_label.items()
                if descriptor.model_name == selected_model_id
            ),
            next(iter(self._model_selection_by_label)),
        )
        self.selected_model_var = tk.StringVar(value=default_label)
        self.model_combo = ttk.Combobox(
            selector_section,
            textvariable=self.selected_model_var,
            values=list(self._model_selection_by_label.keys()),
            state="readonly",
            width=36,
            takefocus=True,
        )
        self.model_combo.grid(row=1, column=1, sticky="w", padx=4)
        self.model_combo.bind("<<ComboboxSelected>>", self._on_leg_model_selected)

        nn_row = ttk.Frame(selector_section)
        nn_row.grid(row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))
        ttk.Button(nn_row, text="Open neural network designer", command=self._open_neural_network_designer).pack(side="left")
        architecture_state = "configured" if isinstance(leg.get("architecture_spec"), dict) else "not set"
        ttk.Label(nn_row, text=f"Architecture: {architecture_state}").pack(side="left", padx=(8, 0))

        selected_descriptor = self._model_selection_by_label.get(self.selected_model_var.get())
        spec_requirements = self._spec_requirements_for_descriptor(selected_descriptor)
        spec_labels = {
            "architecture_spec": "Architecture spec",
            "calibration_spec": "Calibration spec",
            "event_process_spec": "Event process spec",
        }

        chips_row = ttk.Frame(selector_section)
        chips_row.grid(row=3, column=0, columnspan=2, sticky="w", pady=(6, 0))
        for idx, spec_key in enumerate(("architecture_spec", "calibration_spec", "event_process_spec")):
            requirement = spec_requirements[spec_key]
            configured = isinstance(leg.get(spec_key), dict)
            if requirement == "required":
                state_text = "required · configured" if configured else "required · missing"
                fg_color = "#1b8f3a" if configured else "#b12704"
            elif requirement == "optional":
                state_text = "optional · configured" if configured else "optional · not set"
                fg_color = "#1b8f3a" if configured else "#6a6a6a"
            else:
                state_text = "not required"
                fg_color = "#6a6a6a"
            tk.Label(
                chips_row,
                text=f"{spec_labels[spec_key]}: {state_text}",
                foreground=fg_color,
                background=chip_background,
            ).grid(row=0, column=idx, sticky="w", padx=(0, 12))

        configure_row = ttk.Frame(selector_section)
        configure_row.grid(row=4, column=0, columnspan=2, sticky="w", pady=(4, 0))
        required_actions: list[tuple[str, object]] = []
        if spec_requirements["architecture_spec"] == "required":
            required_actions.append(("Configure architecture…", self._open_neural_network_designer))
        if spec_requirements["calibration_spec"] == "required":
            required_actions.append(("Configure calibration…", self._open_calibration_spec_designer))
        if spec_requirements["event_process_spec"] == "required":
            required_actions.append(("Configure event process…", self._open_event_process_designer))

        if required_actions:
            for idx, (label, callback) in enumerate(required_actions):
                ttk.Button(configure_row, text=label, command=callback).grid(row=0, column=idx, sticky="w", padx=(0, 6))
        else:
            ttk.Label(configure_row, text="No required model-specific spec configuration.").grid(row=0, column=0, sticky="w")

        self.leg_control_vars = {}
        schema = LEG_CONTROL_GROUPS[str(leg["model_type"])]
        controls = leg["controls"]

        high_impact_groups = ("signal_parameters", "sizing_risk_caps")
        for group_key in high_impact_groups:
            group_controls = schema.get(group_key, [])
            section = ttk.LabelFrame(self.basics_form_container, text=GROUP_TITLES[group_key], padding=8)
            section.pack(fill="x", pady=(0, 8))
            ttk.Label(
                section,
                text="High-impact knobs used most often.",
                foreground="#555555",
            ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 4))
            for row, control in enumerate(group_controls, start=1):
                self._build_control_row(
                    section,
                    row=row,
                    control=control,
                    value=controls.get(control["key"], control["default"]),
                    width=10,
                )

        self.hyperparameter_vars = {}
        hyperparameter_section = ttk.LabelFrame(self.advanced_form_container, text="Model hyperparameters", padding=8)
        hyperparameter_section.pack(fill="x", pady=(0, 8))
        ttk.Label(
            hyperparameter_section,
            text="Low-level optimizer and model tuning controls.",
            foreground="#555555",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 4))
        hyperparameters = leg.get("hyperparameters", {})
        if isinstance(hyperparameters, dict) and hyperparameters:
            for row, key in enumerate(sorted(hyperparameters), start=1):
                ttk.Label(hyperparameter_section, text=key).grid(row=row, column=0, sticky="w")
                var = tk.StringVar(value=str(hyperparameters[key]))
                entry = ttk.Entry(hyperparameter_section, textvariable=var, width=14, takefocus=True)
                entry.grid(row=row, column=1, sticky="w", padx=4)
                entry.bind("<FocusOut>", lambda _e, hyper_key=key: self._on_hyperparameter_edited(hyper_key))
                self.hyperparameter_vars[key] = var
        else:
            ttk.Label(hyperparameter_section, text="No editable hyperparameters for selected model.").grid(row=1, column=0, sticky="w")

        for group_key, group_controls in schema.items():
            if group_key in high_impact_groups:
                continue
            section = ttk.LabelFrame(self.advanced_form_container, text=GROUP_TITLES[group_key], padding=8)
            section.pack(fill="x", pady=(0, 8))
            ttk.Label(section, text="Advanced controls.", foreground="#555555").grid(
                row=0, column=0, columnspan=3, sticky="w", pady=(0, 4)
            )
            for row, control in enumerate(group_controls, start=1):
                self._build_control_row(
                    section,
                    row=row,
                    control=control,
                    value=controls.get(control["key"], control["default"]),
                    width=10,
                )

        if not self._initial_focus_set:
            self.model_combo.focus_set()
            self._initial_focus_set = True

        self._update_validation_and_actions()

    def _on_training_mode_selected(self, _event=None) -> None:
        selected_label = self.training_mode_var.get()
        for mode, label in TRAINING_MODE_CHOICES.items():
            if label == selected_label:
                self.training_mode = mode
                break
        self._persist_editor_state()
        self._update_validation_and_actions()

    @staticmethod
    def _preview_selected_model_id(leg: dict[str, object]) -> str:
        model_id = str(leg.get("model_id", "")).strip().lower()
        if model_id:
            return model_id
        return str(leg.get("selected_model_id", "")).strip().lower()

    @staticmethod
    def _preview_select_model_id(mode: str, selected_model_id: str, allowed_model_ids: list[str]) -> str:
        if not allowed_model_ids:
            return ""

        if mode in {"single_model", "auto_model_search", "", "auto"}:
            if selected_model_id and selected_model_id in allowed_model_ids:
                return selected_model_id
            return allowed_model_ids[0]

        if mode == "ensemble":
            if "meta_label_classifier" in allowed_model_ids:
                return "meta_label_classifier"
            return allowed_model_ids[0]

        if mode in allowed_model_ids:
            return mode

        return allowed_model_ids[0]

    @staticmethod
    def _preview_candidate_model_ids(mode: str, selected_model_id: str, allowed_model_ids: list[str]) -> list[str]:
        if not allowed_model_ids:
            return []
        if mode in {"auto_model_search", "auto"}:
            return list(allowed_model_ids)
        if mode in {"single_model", "", "placeholder", "dev", "test"}:
            if selected_model_id and selected_model_id in allowed_model_ids:
                return [selected_model_id]
            return [allowed_model_ids[0]]
        if mode == "ensemble":
            return ["meta_label_classifier"] if "meta_label_classifier" in allowed_model_ids else [allowed_model_ids[0]]
        if mode in allowed_model_ids:
            return [mode]
        return [allowed_model_ids[0]]

    @staticmethod
    def _preview_spec_payload_status(payload: object) -> str:
        if not isinstance(payload, dict):
            return "none"
        keys = sorted(str(key) for key in payload.keys())
        if not keys:
            return "empty object"
        return f"configured ({', '.join(keys[:6])}{'…' if len(keys) > 6 else ''})"

    def _training_profile_payload(self) -> dict[str, object]:
        definition = self._active_regime_definition()
        payload = definition.get("training_profile")
        return payload if isinstance(payload, dict) else {}

    @staticmethod
    def _profile_value_for_leg(mapping: object, *, leg_name: str, leg_family: str, leg_index: int) -> object:
        if not isinstance(mapping, dict):
            return None
        for key in (leg_name, leg_family, str(leg_index), "default"):
            if key in mapping:
                return mapping.get(key)
        return None

    def _profile_effects_for_leg(self, profile: dict[str, object], *, leg_name: str, leg_family: str, leg_index: int) -> dict[str, Any]:
        fixed_model = self._profile_value_for_leg(
            profile.get("fixed_model_by_leg", profile.get("fixed_models")),
            leg_name=leg_name,
            leg_family=leg_family,
            leg_index=leg_index,
        )
        if not isinstance(fixed_model, str):
            fixed_model = profile.get("fixed_model_id") if isinstance(profile.get("fixed_model_id"), str) else None

        auto_seed = self._profile_value_for_leg(
            profile.get("auto_search_seed_model_ids_by_leg", profile.get("auto_search_seed_models")),
            leg_name=leg_name,
            leg_family=leg_family,
            leg_index=leg_index,
        )
        if auto_seed is None:
            auto_seed = profile.get("auto_search_seed_model_ids")

        ensemble_members = self._profile_value_for_leg(
            profile.get("ensemble_member_model_ids_by_leg", profile.get("ensemble_members")),
            leg_name=leg_name,
            leg_family=leg_family,
            leg_index=leg_index,
        )
        if ensemble_members is None:
            ensemble_members = profile.get("ensemble_member_model_ids")

        return {
            "fixed_model": fixed_model.strip().lower() if isinstance(fixed_model, str) and fixed_model.strip() else None,
            "auto_seed_models": [str(item).strip().lower() for item in auto_seed] if isinstance(auto_seed, list) else [],
            "ensemble_members": [str(item).strip().lower() for item in ensemble_members] if isinstance(ensemble_members, list) else [],
        }

    def _build_training_execution_preview(self) -> str:
        mode = str(getattr(self, "training_mode", "auto_model_search")).strip().lower()
        profile = self._training_profile_payload()
        profile_label = str(self._active_regime_definition().get("label", "Active profile"))
        warnings: list[str] = []
        lines = [
            f"Profile: {profile_label}",
            f"Training mode: {mode or 'auto_model_search'}",
        ]

        if not self.regime_legs:
            lines.append("No legs configured.")
            self._training_preview_warnings = warnings
            return "\n".join(lines)

        for idx, leg in enumerate(self.regime_legs, start=1):
            try:
                mapped = to_regime_leg_spec(leg)
            except ValueError as exc:
                lines.append("")
                lines.append(f"Leg {idx}: mapping error ({exc})")
                continue
            leg_family = mapped.leg_spec.leg_family
            leg_name = str(leg.get("name", f"Leg {idx}"))
            allowed_model_ids = [descriptor.model_name for descriptor in list_models_for_leg(leg_family)]
            selected_model_id = self._preview_selected_model_id(leg)
            profile_effects = self._profile_effects_for_leg(profile, leg_name=leg_name, leg_family=leg_family, leg_index=idx - 1)

            effective_selected = profile_effects["fixed_model"] or selected_model_id
            candidate_ids = self._preview_candidate_model_ids(mode, effective_selected, allowed_model_ids)
            selected_for_execution = self._preview_select_model_id(mode, effective_selected, allowed_model_ids)

            lines.append("")
            lines.append(f"Leg {idx}: {leg_name} [{leg_family}]")
            lines.append(f"  • Candidate model ids: {', '.join(candidate_ids) if candidate_ids else 'none'}")
            lines.append(f"  • Locked/effective model selection: {selected_for_execution or 'none'}")
            lines.append(
                "  • Profile contributions: "
                + (
                    f"fixed model={profile_effects['fixed_model']}"
                    if profile_effects["fixed_model"]
                    else "fixed model=none"
                )
                + f"; auto-search seeds={', '.join(profile_effects['auto_seed_models']) if profile_effects['auto_seed_models'] else 'none'}"
                + f"; ensemble members={', '.join(profile_effects['ensemble_members']) if profile_effects['ensemble_members'] else 'none'}"
            )
            lines.append(
                "  • Spec payloads: "
                + f"architecture={self._preview_spec_payload_status(leg.get('architecture_spec'))}, "
                + f"calibration={self._preview_spec_payload_status(leg.get('calibration_spec'))}, "
                + f"event={self._preview_spec_payload_status(leg.get('event_process_spec'))}"
            )

            if mode in {"auto_model_search", "auto"} and selected_model_id and not profile_effects["fixed_model"]:
                warnings.append(
                    f"{leg_name}: selected model '{selected_model_id}' is not locked in auto-model-search; all candidates are evaluated."
                )
            if mode == "ensemble" and profile_effects["ensemble_members"] and "meta_label_classifier" not in candidate_ids:
                warnings.append(
                    f"{leg_name}: profile ensemble members are listed, but pipeline may fallback to first catalog model when meta_label_classifier is unavailable."
                )

        if warnings:
            lines.append("")
            lines.append("Warnings:")
            for warning in warnings:
                lines.append(f"  - {warning}")
        self._training_preview_warnings = warnings
        return "\n".join(lines)

    def _update_training_execution_preview(self) -> None:
        if not hasattr(self, "training_preview_text"):
            return
        preview_text = self._build_training_execution_preview()
        self.training_preview_text.configure(state="normal")
        self.training_preview_text.delete("1.0", tk.END)
        self.training_preview_text.insert("1.0", preview_text)
        self.training_preview_text.configure(state="disabled")

    def _on_leg_model_selected(self, _event=None) -> None:
        leg = self._selected_leg()
        selected_descriptor = self._model_selection_by_label.get(self.selected_model_var.get())
        if selected_descriptor is None:
            return
        leg["selected_model_id"] = selected_descriptor.model_name
        leg["model_id"] = selected_descriptor.model_name
        leg["hyperparameters"] = dict(selected_descriptor.hyperparameter_template)
        self._persist_editor_state()
        self._load_selected_leg_into_form()

    def _on_hyperparameter_edited(self, key: str) -> None:
        leg = self._selected_leg()
        hyperparameters = leg.get("hyperparameters")
        if not isinstance(hyperparameters, dict):
            hyperparameters = {}
            leg["hyperparameters"] = hyperparameters
        raw = self.hyperparameter_vars[key].get().strip()
        if raw.lower() in {"true", "false"}:
            parsed: Any = raw.lower() == "true"
        else:
            try:
                parsed = float(raw) if "." in raw else int(raw)
            except ValueError:
                parsed = raw
        hyperparameters[key] = parsed
        self._persist_editor_state()
        self._update_validation_and_actions()

    def _on_control_edited(self, key: str) -> None:
        leg = self._selected_leg()
        raw = self.leg_control_vars[key].get()
        schema_map = {
            control["key"]: control
            for group in LEG_CONTROL_GROUPS[str(leg["model_type"])] .values()
            for control in group
        }
        control = schema_map[key]
        try:
            parsed = float(raw)
        except ValueError:
            parsed = control["default"]
            self.leg_control_vars[key].set(str(parsed))

        leg["controls"][key] = parsed
        self._persist_editor_state()
        self._update_validation_and_actions()

    def _compute_insights_text(self, leg: dict[str, object]) -> str:
        controls = leg["controls"]
        model = str(leg["model_type"])
        lookback = float(controls.get("lookback_days", 30))
        turnover = float(controls.get("turnover_limit", 0.4))
        confidence = float(controls.get("model_confidence_min", 0.6))
        detection_threshold = float(
            controls.get("detection_threshold", controls.get("entry_zscore", 1.0))
        )

        if model == "Trend Following":
            pros = "Pros: Captures persistent moves and scales well in directional markets."
            cons = "Cons: Vulnerable to chop and delayed reversals."
            use_case = "Best use case: Medium-to-long trends with moderate costs."
            failure = "Failure mode: Sideways whipsaws when turnover is high." if turnover > 0.45 else "Failure mode: Fast reversals after prolonged rallies."
        elif model == "Mean Reversion":
            pros = "Pros: Monetizes short-term dislocations and mean-reverting microstructure effects."
            cons = "Cons: Can compound losses when trends persist."
            use_case = "Best use case: Range-bound or oscillatory tape with stable liquidity."
            failure = "Failure mode: Momentum regimes overpower entry thresholds."
        elif model == "Regime Change":
            pros = "Pros: Reacts quickly to state transitions and can reduce exposure before drawdowns deepen."
            cons = "Cons: Sensitive settings can overtrade during noisy consolidations."
            use_case = "Best use case: Macro inflection points and abrupt volatility/dispersion shifts."
            failure = (
                "Failure mode: False transition triggers create churn and slippage drag."
                if detection_threshold < 1.2
                else "Failure mode: Threshold too strict causes late response to real regime breaks."
            )
        else:
            pros = "Pros: Provides convex protection during volatility spikes."
            cons = "Cons: Carry drag in calm markets."
            use_case = "Best use case: Regime shifts, macro events, and stress episodes."
            failure = "Failure mode: Repeated false breakouts with elevated slippage." if confidence < 0.65 else "Failure mode: Late activation after volatility already repriced."

        nuance = "\nKnob readout: "
        nuance += "long-horizon bias" if lookback >= 80 else "short-horizon sensitivity"
        nuance += ", conservative activation" if confidence >= 0.75 else ", opportunistic activation"
        return "\n".join([pros, cons, use_case, failure + nuance])

    def _validation_snapshot(self, leg: dict[str, object] | None = None) -> dict[str, str]:
        leg = leg or self._selected_leg()
        controls = leg["controls"]
        model = str(leg.get("model_type", ""))
        lookback = float(controls.get("lookback_days", 30))
        detection_threshold = float(
            controls.get("detection_threshold", controls.get("entry_zscore", 1.0))
        )
        turnover = float(controls.get("turnover_limit", 0.3))
        slippage = float(controls.get("slippage_bps", 8))

        if model == "Regime Change":
            data_sufficiency = "green" if lookback >= 45 else "red"
            if lookback < 60 and detection_threshold > 2.4:
                overfit_risk = "red"
            elif lookback < 90 and detection_threshold > 2.0:
                overfit_risk = "yellow"
            else:
                overfit_risk = "green"
        else:
            data_sufficiency = "green" if lookback >= 30 else "red"
            if lookback < 50 and detection_threshold > 2.8:
                overfit_risk = "red"
            elif lookback < 80 and detection_threshold > 2.2:
                overfit_risk = "yellow"
            else:
                overfit_risk = "green"

        if turnover > 1.0 or slippage > 35:
            execution_realism = "red"
        elif turnover > 0.6 or slippage > 20:
            execution_realism = "yellow"
        else:
            execution_realism = "green"
        return {
            "data_sufficiency": data_sufficiency,
            "overfit_risk": overfit_risk,
            "execution_realism": execution_realism,
        }

    def _invalid_knob_message(self, leg: dict[str, object]) -> str | None:
        controls = leg["controls"]
        model = str(leg.get("model_type", ""))
        detection_threshold = float(
            controls.get("detection_threshold", controls.get("entry_zscore", 1.0))
        )
        confidence = float(controls.get("model_confidence_min", 0.6))
        drawdown = float(controls.get("max_drawdown_stop", 0.1))
        max_position = float(controls.get("max_position_pct", 0.05))

        if model == "Regime Change":
            if detection_threshold < 0.8:
                return "Detection threshold is too permissive for regime change detection."
            if detection_threshold > 3.5:
                return "Detection threshold is too strict and may miss real regime transitions."
            if drawdown > 0.2:
                return "Regime Change max drawdown stop should be 20% or lower."
            if max_position > 0.12:
                return "Regime Change max position % should be 12% or lower."
        elif detection_threshold < 1.0 and confidence > 0.85:
            return "Entry z-score is too permissive for the selected confidence floor."

        if max_position > drawdown:
            return "Max position % cannot exceed max drawdown stop."
        return None

    def _can_train_export(self) -> tuple[bool, str]:
        leg = self._selected_leg()
        invalid = self._invalid_knob_message(leg)
        if invalid:
            return False, invalid

        try:
            blockers = self._required_spec_blockers(leg)
            if blockers:
                missing = ", ".join(blockers)
                return False, f"Selected model requires missing specs: {missing}."

            for ui_leg in self.regime_legs:
                to_regime_leg_spec(ui_leg)
        except ValueError as exc:
            return False, str(exc)

        statuses = self._validation_snapshot(leg)
        if "red" in statuses.values():
            return False, "One or more validation badges are red. Resolve before training/export."
        if "yellow" in statuses.values():
            return True, "Warning: yellow validation badges indicate moderate model risk."
        return True, "Validation healthy. Train/export ready."

    def _update_validation_and_actions(self) -> None:
        leg = self._selected_leg()
        statuses = self._validation_snapshot(leg)
        self._update_training_execution_preview()

        if hasattr(self, "risk_summary_var"):
            controls = leg["controls"]
            self.risk_summary_var.set(
                f"Model: {leg['model_type']}\n"
                f"Max position: {controls.get('max_position_pct', 0):.2f}\n"
                f"Max drawdown stop: {controls.get('max_drawdown_stop', 0):.2f}\n"
                f"Turnover limit: {controls.get('turnover_limit', 0):.2f}\n"
                f"Slippage: {controls.get('slippage_bps', 0):.1f} bps"
            )

        if hasattr(self, "pros_cons_text"):
            self.pros_cons_text.configure(state="normal")
            self.pros_cons_text.delete("1.0", tk.END)
            self.pros_cons_text.insert("1.0", self._compute_insights_text(leg))
            self.pros_cons_text.configure(state="disabled")

        for key, status in statuses.items():
            if key in self.validation_badge_vars:
                self.validation_badge_vars[key].set(f"{key.replace('_', ' ').title()}: {status.upper()}")
            if key in self.validation_badges:
                self.validation_badges[key].configure(foreground=STATUS_COLORS[status])

        can_run, message = self._can_train_export()
        blockers = self._required_spec_blockers(leg)
        if blockers:
            blocker_text = ", ".join(blockers)
            message = f"Blocked: selected model requires {blocker_text}. Use Configure actions in Model selection."
            can_run = False
        if self._is_training:
            can_run = False
            message = "Training currently running..."
        if hasattr(self, "train_button"):
            self.train_button.configure(state=("normal" if can_run else "disabled"))
        if hasattr(self, "export_button"):
            self.export_button.configure(state=("normal" if can_run else "disabled"))
        if hasattr(self, "validation_message_var"):
            preview_warnings = getattr(self, "_training_preview_warnings", [])
            if preview_warnings:
                warning_summary = preview_warnings[0]
                if len(preview_warnings) > 1:
                    warning_summary += f" (+{len(preview_warnings) - 1} more preview warning(s))"
                self.validation_message_var.set(f"{message} Preview warning: {warning_summary}")
            else:
                self.validation_message_var.set(message)

    def _append_log(self, message: str) -> None:
        if not hasattr(self, "run_logs"):
            return
        self.run_logs.insert(tk.END, message + "\n")
        self.run_logs.see(tk.END)

    def _append_structured_log(self, *, level: str, event: str, details: str) -> None:
        ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
        self._append_log(f"[{ts}] [{level.upper()}] {event}: {details}")

    def _active_regime_definition(self) -> dict[str, object]:
        regime_id = self.controller.state.active_regime_id
        definitions = self.controller.state.regime_definitions
        definition = definitions.get(regime_id) if regime_id else None
        if isinstance(definition, dict):
            return definition
        if definitions:
            return next(iter(definitions.values()))
        return {}

    def _build_regime_training_request(self) -> RegimeTrainingRequest:
        definition = self._active_regime_definition()
        regime_id = self.controller.state.active_regime_id or "baseline"
        regime_label = str(definition.get("label", regime_id))

        training_window = {
            **DEFAULT_REGIME_TRAINING_WINDOW,
            **(definition.get("training_window", {}) if isinstance(definition.get("training_window"), dict) else {}),
        }
        global_risk_limits = {
            **DEFAULT_REGIME_GLOBAL_RISK_LIMITS,
            **(definition.get("global_risk_limits", {}) if isinstance(definition.get("global_risk_limits"), dict) else {}),
        }
        confidence_thresholds = {
            **DEFAULT_REGIME_CONFIDENCE_THRESHOLDS,
            **(definition.get("confidence_thresholds", {}) if isinstance(definition.get("confidence_thresholds"), dict) else {}),
        }
        training_data_settings = {
            **DEFAULT_REGIME_TRAINING_DATA_SETTINGS,
            **(definition.get("training_data_settings", {}) if isinstance(definition.get("training_data_settings"), dict) else {}),
        }

        legs: list[RegimeLegTrainingConfig] = []
        for leg in self.regime_legs:
            controls = leg.get("controls", {})
            cast_controls = {key: float(value) for key, value in controls.items()}
            mapped_leg = to_regime_leg_spec(leg)
            mapped_controls = {key: float(value) for key, value in mapped_leg.leg_spec.knobs.items()}
            selected_model_id = str(leg.get("selected_model_id", "")).strip()
            hyperparameters_raw = leg.get("hyperparameters", {})
            hyperparameters: dict[str, Any] = {}
            if isinstance(hyperparameters_raw, dict):
                hyperparameters = dict(hyperparameters_raw)
            legs.append(
                RegimeLegTrainingConfig(
                    name=str(leg.get("name", "Unnamed leg")),
                    model_type=mapped_leg.leg_spec.leg_family,
                    controls={**cast_controls, **mapped_controls},
                    model_id=selected_model_id,
                    selected_model_id=selected_model_id,
                    hyperparameters=hyperparameters,
                    architecture_spec=leg.get("architecture_spec") if isinstance(leg.get("architecture_spec"), dict) else None,
                    calibration_spec=leg.get("calibration_spec") if isinstance(leg.get("calibration_spec"), dict) else None,
                    event_process_spec=leg.get("event_process_spec") if isinstance(leg.get("event_process_spec"), dict) else None,
                )
            )

        risk_limits = {
            **{key: float(value) for key, value in global_risk_limits.items()},
            **{f"confidence_{key}": float(value) for key, value in confidence_thresholds.items()},
        }

        backtest_root = str(self.controller.state.backtest_settings.get("backtest_data_root", BACKTEST_CACHE_DIR))
        training_data_settings["cache_root"] = backtest_root
        training_data_settings["universe_symbols"] = [item.strip().upper() for item in self.controller.state.tickers if item.strip()]

        return RegimeTrainingRequest(
            schema_version=2,
            regime_id=regime_id,
            regime_name=regime_label,
            model_choice=str(getattr(self, "training_mode", "auto_model_search")),
            training_window={key: int(value) for key, value in training_window.items()},
            risk_limits=risk_limits,
            legs=tuple(legs),
            training_data_settings=training_data_settings,
        )

    def _open_neural_network_designer(self) -> None:
        leg = self._selected_leg()
        if self._nn_designer_window is not None and self._nn_designer_window.winfo_exists():
            self._nn_designer_window.focus_set()
            return

        custom_presets = leg.get("nn_custom_presets", {})
        if not isinstance(custom_presets, dict):
            custom_presets = {}

        def _save(spec: dict[str, Any]) -> None:
            leg["architecture_spec"] = dict(spec)
            if self._nn_designer_window is not None:
                leg["nn_custom_presets"] = self._nn_designer_window.custom_presets
            self._persist_editor_state()
            self._load_selected_leg_into_form()

        self._nn_designer_window = NeuralNetworkDesignerPage(
            self,
            initial_spec=leg.get("architecture_spec") if isinstance(leg.get("architecture_spec"), dict) else None,
            on_save=_save,
            custom_presets=custom_presets,
        )

    def _open_calibration_spec_designer(self) -> None:
        leg = self._selected_leg()
        if self._calibration_designer_window is not None and self._calibration_designer_window.winfo_exists():
            self._calibration_designer_window.focus_set()
            return

        def _save(spec: dict[str, Any]) -> None:
            leg["calibration_spec"] = dict(spec)
            self._persist_editor_state()
            self._load_selected_leg_into_form()

        self._calibration_designer_window = CalibrationSpecDesignerPage(
            self,
            initial_spec=leg.get("calibration_spec") if isinstance(leg.get("calibration_spec"), dict) else None,
            on_save=_save,
        )

    def _open_event_process_designer(self) -> None:
        leg = self._selected_leg()
        if self._event_process_designer_window is not None and self._event_process_designer_window.winfo_exists():
            self._event_process_designer_window.focus_set()
            return

        def _save(spec: dict[str, Any]) -> None:
            leg["event_process_spec"] = dict(spec)
            self._persist_editor_state()
            self._load_selected_leg_into_form()

        self._event_process_designer_window = EventProcessDesignerPage(
            self,
            initial_spec=leg.get("event_process_spec") if isinstance(leg.get("event_process_spec"), dict) else None,
            on_save=_save,
        )

    def _last_successful_training_run(self) -> dict[str, object] | None:
        runs = self.controller.state.regime_training_runs
        for run in reversed(runs):
            if run.get("status") == "success":
                return run
        return None

    def _run_train(self) -> None:
        if self._is_training:
            messagebox.showinfo("Training in progress", "A regime training run is already in progress.")
            return

        can_run, msg = self._can_train_export()
        if not can_run:
            messagebox.showinfo("Validation blocked", msg)
            return

        try:
            request = self._build_regime_training_request()
        except Exception as exc:
            self._append_structured_log(level="error", event="request_build_failed", details=str(exc))
            messagebox.showerror("Training failed", f"Failed to build training request: {exc}")
            return

        self._is_training = True
        self.train_button.configure(state="disabled")
        self.export_button.configure(state="disabled")
        self._append_structured_log(
            level="info",
            event="training_queued",
            details=f"regime={request.regime_label}, legs={len(request.legs)}",
        )

        def _worker() -> None:
            try:
                cache_report = self._prepare_regime_training_cache(request)
                if cache_report.get("note"):
                    self.after(
                        0,
                        lambda note=str(cache_report.get("note")): self._append_structured_log(
                            level="info", event="cache_check", details=note
                        ),
                    )
                if cache_report.get("failing_symbols"):
                    self.after(
                        0,
                        lambda payload=cache_report: self._append_structured_log(
                            level="warning",
                            event="cache_audit_failures",
                            details=(
                                f"failing_symbols={payload.get('failing_symbols', [])}; "
                                f"details={payload.get('failing_details', {})}"
                            ),
                        ),
                    )
                artifact_path = str(cache_report.get("artifact_path", "")).strip()
                if artifact_path:
                    request.training_data_settings["cache_audit_report"] = artifact_path
                result = execute_regime_training_pipeline(request)
                metadata = dict(result.metadata)
                if cache_report.get("artifact_path"):
                    metadata["cache_audit_report"] = str(cache_report["artifact_path"])
                result = RegimeTrainingResult(
                    run_id=result.run_id,
                    status=result.status,
                    metrics=result.metrics,
                    artifact_paths=result.artifact_paths,
                    timestamps=result.timestamps,
                    warnings=result.warnings,
                    errors=result.errors,
                    error_payload=result.error_payload,
                    summary=result.summary,
                    metadata=metadata,
                    logs=result.logs,
                )
            except Exception as exc:
                self.after(0, lambda error=exc: self._on_training_failed(error))
                return
            if result.status == "success":
                self.after(0, lambda training_result=result: self._on_training_succeeded(training_result))
            else:
                self.after(0, lambda training_result=result: self._on_training_result_failed(training_result))

        threading.Thread(target=_worker, daemon=True).start()

    def _prepare_regime_training_cache(self, request: RegimeTrainingRequest) -> dict[str, Any]:
        settings = dict(request.training_data_settings or {})
        symbols = [str(item).strip().upper() for item in settings.get("universe_symbols", []) if str(item).strip()]
        if not symbols:
            return {"note": "no universe symbols configured for cache check", "failing_symbols": []}

        required_years = max(1, int(settings.get("required_history_years", 5)))
        cache_root = Path(str(settings.get("cache_root", BACKTEST_CACHE_DIR))).expanduser()
        strict = bool(settings.get("cache_audit_strict", True))
        min_symbol_coverage_ratio = float(settings.get("min_symbol_coverage_ratio", 1.0))
        min_bars_per_year = max(1, int(settings.get("min_bars_per_year", 1)))

        report = audit_universe_history(
            symbols=symbols,
            cache_root=cache_root,
            min_years=required_years,
            timeframe="1m",
            strict=strict,
            min_symbol_coverage_ratio=min_symbol_coverage_ratio,
            min_bars_per_year=min_bars_per_year,
        )
        report["requested_symbols"] = symbols
        report["run_audit_created_at"] = datetime.now(timezone.utc).isoformat()

        run_root = Path("data/regime_training_runs")
        run_root.mkdir(parents=True, exist_ok=True)
        report_path = run_root / f"run_audit_{datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S%f')}.json"
        report_path.write_text(json.dumps(report, indent=2, sort_keys=True), encoding="utf-8")

        failing_symbols = [str(item).upper() for item in report.get("failing_symbols", [])]
        if not failing_symbols:
            return {
                "note": f"cache audit passed for {len(symbols)} symbol(s)",
                "failing_symbols": [],
                "failing_details": {},
                "artifact_path": str(report_path),
            }

        if not bool(settings.get("enable_cache_backfill", True)):
            return {
                "note": (
                    f"cache audit failed for {len(failing_symbols)} symbol(s), backfill disabled"
                ),
                "failing_symbols": failing_symbols,
                "failing_details": report.get("failing_details", {}),
                "artifact_path": str(report_path),
            }

        if not self.controller.api_key:
            return {
                "note": (
                    f"cache audit failed for {len(failing_symbols)} symbol(s), skipped backfill (missing API key)"
                ),
                "failing_symbols": failing_symbols,
                "failing_details": report.get("failing_details", {}),
                "artifact_path": str(report_path),
            }

        end_date = date.today()
        start_date = end_date - timedelta(days=365 * required_years + 7)
        run_backtest_cache(
            tickers=failing_symbols,
            start_date=start_date,
            end_date=end_date,
            cache_root=cache_root,
            api_key=self.controller.api_key,
        )
        return {
            "note": (
                f"cache backfill requested for failing symbols={failing_symbols} over {required_years} years"
            ),
            "failing_symbols": failing_symbols,
            "failing_details": report.get("failing_details", {}),
            "artifact_path": str(report_path),
        }

    def _on_training_result_failed(self, result: RegimeTrainingResult) -> None:
        self._is_training = False
        self._update_validation_and_actions()
        detail = "; ".join(result.errors) if result.errors else "unknown error"
        self._append_structured_log(level="error", event="training_failed", details=detail)
        if "INSUFFICIENT_REAL_HISTORY" in detail:
            messagebox.showerror(
                "Training blocked: insufficient real history",
                "Regime training requires sufficient real market history. "
                "Synthetic fallback is disabled for this regime, so training cannot proceed.\n"
                f"{detail}",
            )
            return
        messagebox.showerror(
            "Training failed",
            f"Regime training failed with status={result.status}.\n{detail}",
        )

    def _on_training_failed(self, exc: Exception) -> None:
        self._is_training = False
        self._update_validation_and_actions()
        self._append_structured_log(level="error", event="training_failed", details=str(exc))
        messagebox.showerror("Training failed", f"Regime training failed:\n{exc}")

    def _on_training_succeeded(self, result: RegimeTrainingResult) -> None:
        self._is_training = False
        run_record = {
            "run_id": result.run_id,
            "status": result.status,
            "started_at": result.started_at,
            "completed_at": result.completed_at,
            "artifact_path": result.artifact_path,
            "summary": result.summary,
            "metrics": result.metrics,
            "metadata": result.metadata,
            "logs": list(result.logs),
        }
        self.controller.state.regime_training_runs.append(run_record)
        self.controller.persist_state()
        self._append_structured_log(level="info", event="training_completed", details=result.summary)
        self._append_structured_log(level="info", event="training_artifact", details=result.artifact_path)
        messagebox.showinfo("Training completed", result.summary)
        self._update_validation_and_actions()

    def _run_export(self) -> None:
        can_run, msg = self._can_train_export()
        if not can_run:
            messagebox.showinfo("Validation blocked", msg)
            return

        latest = self._last_successful_training_run()
        if latest is None:
            messagebox.showinfo(
                "Export blocked",
                "No successful training run exists yet. Train a regime before exporting.",
            )
            self._append_structured_log(level="warning", event="export_blocked", details="missing_successful_training")
            return

        artifact_path = str(latest.get("artifact_path", ""))
        if not artifact_path.strip():
            self._append_structured_log(level="warning", event="export_blocked", details="missing_manifest_path")
            messagebox.showinfo("Export blocked", "Latest successful training run does not have a manifest path.")
            return

        try:
            bundle = export_regime_training_bundle(artifact_path)
        except Exception as exc:
            self._append_structured_log(level="error", event="export_failed", details=str(exc))
            messagebox.showerror("Export failed", f"Failed to export training bundle:\n{exc}")
            return

        self._append_structured_log(level="info", event="export_completed", details=f"bundle={bundle.bundle_dir}")
        messagebox.showinfo(
            "Export completed",
            "Export bundle created from the latest successful training artifact.\n"
            f"Bundle path: {bundle.bundle_dir}\n"
            f"Bundle manifest: {bundle.bundle_manifest_path}\n"
            f"Deployment version: {bundle.deployment_version}",
        )

from __future__ import annotations

import threading
from datetime import datetime, timezone
import tkinter as tk
from tkinter import messagebox, ttk
from typing import TYPE_CHECKING

from backtesting.regime_training_pipeline import (
    RegimeLegTrainingConfig,
    RegimeTrainingRequest,
    RegimeTrainingResult,
    run_regime_training,
)
from ui.regime_mapping import to_regime_leg_spec
from config import (
    DEFAULT_REGIME_CONFIDENCE_THRESHOLDS,
    DEFAULT_REGIME_GLOBAL_RISK_LIMITS,
    DEFAULT_REGIME_TRAINING_WINDOW,
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
}

GROUP_TITLES = {
    "signal_parameters": "Signal parameters",
    "sizing_risk_caps": "Sizing / risk caps",
    "turnover_liquidity_assumptions": "Turnover / liquidity assumptions",
    "confidence_thresholds": "Confidence thresholds",
}

STATUS_COLORS = {"green": "#1b8f3a", "yellow": "#b07d00", "red": "#b12704"}


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
        self.selected_leg_index = 0
        self.leg_control_vars: dict[str, tk.Variable] = {}
        self._is_training = False

        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        ttk.Label(self, text="Create Regime", font=("Arial", 18, "bold")).grid(row=0, column=0, sticky="w", padx=12, pady=(10, 6))

        self.main_pane = ttk.Panedwindow(self, orient=tk.VERTICAL)
        self.main_pane.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 10))

        top_pane = ttk.Panedwindow(self.main_pane, orient=tk.HORIZONTAL)
        self.main_pane.add(top_pane, weight=5)

        self.legs_panel = ttk.Frame(top_pane, padding=8)
        self.config_panel = ttk.Frame(top_pane, padding=8)
        self.summary_panel = ttk.Frame(top_pane, padding=8)
        top_pane.add(self.legs_panel, weight=2)
        top_pane.add(self.config_panel, weight=4)
        top_pane.add(self.summary_panel, weight=3)

        self.bottom_panel = ttk.Frame(self.main_pane, padding=8)
        self.main_pane.add(self.bottom_panel, weight=2)

        self._build_legs_panel()
        self._build_config_panel()
        self._build_summary_panel()
        self._build_bottom_panel()

        self._refresh_legs_list()
        self._load_selected_leg_into_form()
        self._update_validation_and_actions()

    def _build_default_leg(self, leg_type: str) -> dict[str, object]:
        controls = {}
        schema = LEG_CONTROL_GROUPS[leg_type]
        for group in schema.values():
            for control in group:
                controls[control["key"]] = control["default"]
        return {
            "name": f"{leg_type} leg",
            "model_type": leg_type,
            "controls": controls,
        }

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

        self.leg_type_var = tk.StringVar(value="Trend Following")
        self.leg_type_combo = ttk.Combobox(
            header,
            textvariable=self.leg_type_var,
            values=list(LEG_CONTROL_GROUPS),
            state="readonly",
        )
        self.leg_type_combo.pack(anchor="w", pady=(6, 0))
        self.leg_type_combo.bind("<<ComboboxSelected>>", self._on_leg_type_selected)

        self.form_container = ttk.Frame(self.config_panel)
        self.form_container.pack(fill="both", expand=True, pady=(8, 0))

    def _build_summary_panel(self) -> None:
        ttk.Label(self.summary_panel, text="Risk summary + strategy fit", font=("Arial", 12, "bold")).pack(anchor="w")

        self.validation_badge_vars = {
            "data_sufficiency": tk.StringVar(),
            "overfit_risk": tk.StringVar(),
            "execution_realism": tk.StringVar(),
        }
        self.validation_badges: dict[str, ttk.Label] = {}

        badge_frame = ttk.Frame(self.summary_panel)
        badge_frame.pack(fill="x", pady=(6, 8))
        for idx, (key, var) in enumerate(self.validation_badge_vars.items()):
            label = ttk.Label(badge_frame, textvariable=var)
            label.grid(row=idx, column=0, sticky="w", pady=1)
            self.validation_badges[key] = label

        self.risk_summary_var = tk.StringVar()
        ttk.Label(self.summary_panel, textvariable=self.risk_summary_var, justify="left", wraplength=320).pack(fill="x", pady=(4, 8))

        self.pros_cons_text = tk.Text(self.summary_panel, height=13, wrap="word")
        self.pros_cons_text.pack(fill="both", expand=True)
        self.pros_cons_text.configure(state="disabled")

    def _build_bottom_panel(self) -> None:
        action_row = ttk.Frame(self.bottom_panel)
        action_row.pack(fill="x")
        ttk.Label(action_row, text="Train / export", font=("Arial", 12, "bold")).pack(side="left")

        self.train_button = ttk.Button(action_row, text="Train", command=self._run_train)
        self.train_button.pack(side="left", padx=(12, 4))
        self.export_button = ttk.Button(action_row, text="Export", command=self._run_export)
        self.export_button.pack(side="left", padx=4)

        self.validation_message_var = tk.StringVar()
        ttk.Label(self.bottom_panel, textvariable=self.validation_message_var, wraplength=760).pack(anchor="w", pady=(6, 4))

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
        self._refresh_legs_list()
        self._load_selected_leg_into_form()

    def remove_selected_leg(self) -> None:
        if len(self.regime_legs) <= 1:
            return
        self.regime_legs.pop(self.selected_leg_index)
        self.selected_leg_index = max(0, self.selected_leg_index - 1)
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
        self._refresh_legs_list()
        self._load_selected_leg_into_form()

    def _load_selected_leg_into_form(self) -> None:
        leg = self._selected_leg()
        if hasattr(self, "leg_type_var"):
            self.leg_type_var.set(str(leg["model_type"]))
        if not hasattr(self, "form_container"):
            self._update_validation_and_actions()
            return

        for widget in self.form_container.winfo_children():
            widget.destroy()

        self.leg_control_vars = {}
        schema = LEG_CONTROL_GROUPS[str(leg["model_type"])]
        controls = leg["controls"]

        for group_key, group_controls in schema.items():
            section = ttk.LabelFrame(self.form_container, text=GROUP_TITLES[group_key], padding=6)
            section.pack(fill="x", pady=3)
            for row, control in enumerate(group_controls):
                ttk.Label(section, text=control["label"]).grid(row=row, column=0, sticky="w")
                var = tk.StringVar(value=str(controls.get(control["key"], control["default"])))
                entry = ttk.Entry(section, textvariable=var, width=12)
                entry.grid(row=row, column=1, sticky="w", padx=4)
                ttk.Label(section, text=control["tooltip"], wraplength=360).grid(row=row, column=2, sticky="w")
                entry.bind("<FocusOut>", lambda _e, key=control["key"]: self._on_control_edited(key))
                self.leg_control_vars[control["key"]] = var

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
        if self._is_training:
            can_run = False
            message = "Training currently running..."
        if hasattr(self, "train_button"):
            self.train_button.configure(state=("normal" if can_run else "disabled"))
        if hasattr(self, "export_button"):
            self.export_button.configure(state=("normal" if can_run else "disabled"))
        if hasattr(self, "validation_message_var"):
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

        legs: list[RegimeLegTrainingConfig] = []
        for leg in self.regime_legs:
            controls = leg.get("controls", {})
            cast_controls = {key: float(value) for key, value in controls.items()}
            mapped_leg = to_regime_leg_spec(leg)
            mapped_controls = {key: float(value) for key, value in mapped_leg.leg_spec.knobs.items()}
            legs.append(
                RegimeLegTrainingConfig(
                    name=str(leg.get("name", "Unnamed leg")),
                    model_type=mapped_leg.leg_spec.leg_family,
                    controls={**cast_controls, **mapped_controls},
                )
            )

        return RegimeTrainingRequest(
            regime_id=regime_id,
            regime_label=regime_label,
            requested_at=datetime.now(timezone.utc).isoformat(),
            training_window={key: int(value) for key, value in training_window.items()},
            global_risk_limits={key: float(value) for key, value in global_risk_limits.items()},
            confidence_thresholds={key: float(value) for key, value in confidence_thresholds.items()},
            legs=tuple(legs),
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
                result = run_regime_training(request)
            except Exception as exc:
                self.after(0, lambda error=exc: self._on_training_failed(error))
                return
            self.after(0, lambda training_result=result: self._on_training_succeeded(training_result))

        threading.Thread(target=_worker, daemon=True).start()

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
        summary = str(latest.get("summary", ""))
        self._append_structured_log(level="info", event="export_completed", details=f"artifact={artifact_path}")
        messagebox.showinfo(
            "Export completed",
            "Export uses the latest successful training artifact.\n"
            f"Path: {artifact_path}\n"
            f"Summary: {summary}",
        )

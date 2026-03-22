from __future__ import annotations

import threading
import csv
import json
import webbrowser
from datetime import date, datetime
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

from analysis.prompt_pack import write_prompt_pack
from execution import (
    JOB_BACKTEST_MULTI_SIGNAL,
    JOB_BACKTEST_OPTIMIZATION,
    JOB_BACKTEST_TIME_SERIES,
    JOB_BACKTEST_TRAINED_REGIME,
    JOB_BACKTEST_WALK_FORWARD,
)
from execution.artifact_sync import DEFAULT_REMOTE_NAMESPACE_PREFIX
from backtesting.application_service import (
    BacktestRequestValidationError,
    BacktestingApplicationService,
    ClassicStrategyRunRequest,
    TrainedRegimeReplayRunRequest,
)
from backtesting.regime_backtest_adapter import (
    RegimeBundleCompatibilityError,
    RegimeBacktestContract,
    RegimeBacktestOption,
    discover_regime_backtest_options,
    load_regime_backtest_contract,
)
from backtesting.scenario_toolkit import list_scenario_pack_templates
from config import BACKTEST_CACHE_DIR, BACKTEST_OUTPUT_DIR, BACKTEST_STRATEGY_PRESETS, BACKTEST_TEST_SUITE_PRESETS, DEFAULT_BACKTEST_SETTINGS
from ui.backtesting_insights import (
    aggregate_regime_market_stress,
    build_guardrails,
    build_scenario_comparison,
    compare_manifests,
    compare_robustness_frontiers,
    fold_variance_rows,
    metric_deltas,
    parameter_diffs,
    parse_tags,
    read_experiment_index,
    read_stress_scenarios,
    apply_governance_decision,
)
from ui.option_registry import (
    ENTRY_SIGNALS,
    EXECUTION_MODELS,
    EXIT_SIGNALS,
    OPTIMIZER_SAMPLERS,
    migration_hint_text,
    normalize_supported_option,
    validate_option_values,
)
from utils.parsing import normalize_cache_root, parse_date, parse_float

STRATEGIES = ["momentum", "xsmom"]
TIMEFRAMES = ["1m", "5m", "15m", "30m", "1h", "1d"]
PORTFOLIO_METHODS = ["equal_weight", "vol_target", "inverse_vol", "capped_optimization", "hrp", "herc"]
BACKTEST_WORKFLOW_TYPE_MAP = {"Classic Strategy": "classic_strategy", "Trained Regime": "trained_regime"}
BACKTEST_WORKFLOW_TYPES = tuple(BACKTEST_WORKFLOW_TYPE_MAP.values())
BACKTEST_WORKFLOW_TYPE_LABELS = {value: key for key, value in BACKTEST_WORKFLOW_TYPE_MAP.items()}

STRESS_PROFILES: dict[str, dict[str, float]] = {
    "Mild": {
        "historical_window_fraction": 0.15,
        "historical_replay_window_bars": 12,
        "synthetic_jump_magnitude": 0.01,
        "synthetic_jump_interval": 10,
        "synthetic_vol_cluster_multiplier": 1.2,
        "overlay_spread_multiplier": 1.5,
        "overlay_liquidity_multiplier": 0.7,
    },
    "Base": {
        "historical_window_fraction": 0.20,
        "historical_replay_window_bars": 20,
        "synthetic_jump_magnitude": 0.02,
        "synthetic_jump_interval": 7,
        "synthetic_vol_cluster_multiplier": 1.6,
        "overlay_spread_multiplier": 2.5,
        "overlay_liquidity_multiplier": 0.4,
    },
    "Severe": {
        "historical_window_fraction": 0.30,
        "historical_replay_window_bars": 36,
        "synthetic_jump_magnitude": 0.035,
        "synthetic_jump_interval": 4,
        "synthetic_vol_cluster_multiplier": 2.2,
        "overlay_spread_multiplier": 4.0,
        "overlay_liquidity_multiplier": 0.25,
    },
}

DEFAULT_OPTIMIZER_SEARCH_SPACE = '{"combo_index":{"type":"discrete","values":[0]}}'
DEFAULT_OPTIMIZER_OBJECTIVES = '[{"name":"sharpe","sense":"maximize"},{"name":"turnover_total","sense":"minimize"},{"name":"max_drawdown","sense":"maximize"}]'
DEFAULT_OPTIMIZER_OBJECTIVE_WEIGHTS = '{}'
DEFAULT_OPTIMIZER_OVERFITTING_PENALTY = '{}'

TIMEFRAME_HISTORY_DAYS = {"1m": 14, "5m": 30, "15m": 60, "30m": 120, "1h": 365, "1d": 3650}
GOVERNANCE_PROMOTION_STATES = ["research", "paper", "shadow", "production"]
GOVERNANCE_APPROVAL_STATES = ["pending", "in_review", "approved", "rejected", "waived"]
BACKTEST_SETTINGS_SCHEMA_VERSION = 3

class BacktestingPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller
        self._updating_wf_fractions = False
        self._advanced_widgets: dict[str, tk.Widget] = {}
        self._advanced_tooltips: dict[str, str] = {}
        self._optimizer_json_lint_vars: dict[str, tk.StringVar] = {}
        self._validation_messages: list[str] = []
        self._stale_preset_messages: list[str] = []
        self._updating_risk_limit_controls = False
        self._regime_backtest_options: list[RegimeBacktestOption] = []
        self._regime_option_lookup: dict[str, RegimeBacktestOption] = {}
        self._active_regime_contract: RegimeBacktestContract | None = None
        self._regime_loading_defaults = False
        self._regime_locked_fields = True
        self._regime_loaded_values: dict[str, str] = {}
        self._regime_immutable_values: dict[str, str] = {}
        self._application_service = BacktestingApplicationService(output_dir=BACKTEST_OUTPUT_DIR)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        page_container = ttk.Frame(self)
        page_container.grid(row=0, column=0, sticky="nsew")
        page_container.columnconfigure(0, weight=1)
        page_container.rowconfigure(0, weight=1)

        self.page_canvas = tk.Canvas(page_container, highlightthickness=0)
        self.page_canvas.grid(row=0, column=0, sticky="nsew")
        page_scrollbar = ttk.Scrollbar(page_container, orient="vertical", command=self.page_canvas.yview)
        page_scrollbar.grid(row=0, column=1, sticky="ns")
        self.page_canvas.configure(yscrollcommand=page_scrollbar.set)

        content = ttk.Frame(self.page_canvas)
        self._page_canvas_window = self.page_canvas.create_window((0, 0), window=content, anchor="nw")
        self.page_canvas.bind("<Configure>", self._on_page_canvas_configure)
        content.bind("<Configure>", self._on_page_frame_configure)
        self.page_canvas.bind("<Enter>", self._bind_mousewheel)
        self.page_canvas.bind("<Leave>", self._unbind_mousewheel)

        ttk.Label(content, text="Backtesting Parameters", font=("Arial", 18, "bold")).pack(pady=10)

        intro = (
            "Choose a strategy and configure its parameters. Shared settings stay visible, "
            "and strategy-specific controls appear only for the selected strategy."
        )
        ttk.Label(content, text=intro, wraplength=950, justify="center").pack(pady=5)

        content = ttk.Frame(content)
        content.pack(fill="both", expand=True, padx=30, pady=10)
        content.columnconfigure(0, weight=1)
        content.rowconfigure(0, weight=1)

        self.section_notebook = ttk.Notebook(content)
        self.section_notebook.grid(row=0, column=0, sticky="nsew")

        run_setup_tab = ttk.Frame(self.section_notebook)
        run_setup_tab.columnconfigure(0, weight=1)
        run_setup_tab.rowconfigure(3, weight=1)
        self.section_notebook.add(run_setup_tab, text="Run Setup")

        ttk.Label(
            run_setup_tab,
            text="Workflow Mode & Presets",
            font=("Arial", 12, "bold"),
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(10, 2))
        ttk.Label(
            run_setup_tab,
            text="Choose Basic/Advanced mode and optionally apply a preset before tuning strategy fields.",
            justify="left",
        ).grid(row=1, column=0, sticky="w", padx=12, pady=(0, 6))

        self.ui_mode_var = tk.StringVar(value="basic")
        self.preset_var = tk.StringVar(value="Custom")
        self._preset_display_to_key = {"Custom": "custom"}
        for preset_key, preset_cfg in BACKTEST_STRATEGY_PRESETS.items():
            display = f"{preset_cfg.get('label', preset_key)} ({preset_key})"
            self._preset_display_to_key[display] = preset_key
        self._preset_key_to_display = {value: key for key, value in self._preset_display_to_key.items()}
        self._suite_display_to_key = {"Custom": "custom"}
        for suite_key, suite_cfg in BACKTEST_TEST_SUITE_PRESETS.items():
            suite_display = f"{suite_cfg.get('label', suite_key)} ({suite_key})"
            self._suite_display_to_key[suite_display] = suite_key
        self._suite_key_to_display = {value: key for key, value in self._suite_display_to_key.items()}

        workflow_frame = ttk.LabelFrame(run_setup_tab, text="Workflow Mode & Presets")
        workflow_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(2, 8))
        workflow_frame.columnconfigure(1, weight=1)
        ttk.Label(workflow_frame, text="Mode").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        mode_row = ttk.Frame(workflow_frame)
        mode_row.grid(row=0, column=1, sticky="w", padx=8, pady=6)
        self.show_advanced_controls_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            mode_row,
            text="Show advanced controls",
            variable=self.show_advanced_controls_var,
            command=self._on_show_advanced_controls_toggled,
        ).pack(side="left")

        ttk.Label(workflow_frame, text="Preset").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        preset_values = list(self._preset_display_to_key.keys())
        self.preset_combo = ttk.Combobox(
            workflow_frame,
            textvariable=self.preset_var,
            state="readonly",
            values=preset_values,
        )
        self.preset_combo.grid(row=1, column=1, sticky="ew", padx=8, pady=6)
        self.preset_combo.bind("<<ComboboxSelected>>", self._on_preset_selected)
        ttk.Button(workflow_frame, text="Restore advanced defaults", command=self._restore_advanced_defaults).grid(
            row=2,
            column=1,
            sticky="w",
            padx=8,
            pady=(0, 6),
        )

        ttk.Label(workflow_frame, text="Test Suite").grid(row=3, column=0, sticky="w", padx=8, pady=6)
        self.test_suite_var = tk.StringVar(value="Custom")
        self.test_suite_combo = ttk.Combobox(
            workflow_frame,
            textvariable=self.test_suite_var,
            state="readonly",
            values=list(self._suite_display_to_key.keys()),
        )
        self.test_suite_combo.grid(row=3, column=1, sticky="ew", padx=8, pady=6)
        self.test_suite_combo.bind("<<ComboboxSelected>>", self._on_test_suite_selected)

        ttk.Label(workflow_frame, text="Backtest Type").grid(row=4, column=0, sticky="w", padx=8, pady=6)
        self.backtest_type_var = tk.StringVar(value="Classic Strategy")
        self.backtest_type_combo = ttk.Combobox(
            workflow_frame,
            textvariable=self.backtest_type_var,
            state="readonly",
            values=list(BACKTEST_WORKFLOW_TYPE_MAP.keys()),
        )
        self.backtest_type_combo.grid(row=4, column=1, sticky="ew", padx=8, pady=6)

        self.start_date_var = tk.StringVar()
        self.end_date_var = tk.StringVar()
        self.starting_capital_var = tk.StringVar()

        regime_replay_frame = ttk.LabelFrame(workflow_frame, text="Regime Replay (required)")
        regime_replay_frame.grid(row=5, column=0, columnspan=2, sticky="ew", padx=8, pady=(4, 8))
        regime_replay_frame.columnconfigure(1, weight=1)

        ttk.Label(regime_replay_frame, text="Trained regime artifact").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        self.trained_regime_var = tk.StringVar(value="")
        self.trained_regime_combo = ttk.Combobox(regime_replay_frame, textvariable=self.trained_regime_var, state="readonly", values=[])
        self.trained_regime_combo.grid(row=0, column=1, sticky="ew", padx=6, pady=4)
        self.trained_regime_combo.bind("<<ComboboxSelected>>", self._on_trained_regime_selected)

        ttk.Label(regime_replay_frame, text="Replay date range").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        date_row = ttk.Frame(regime_replay_frame)
        date_row.grid(row=1, column=1, sticky="w", padx=6, pady=4)
        ttk.Entry(date_row, textvariable=self.start_date_var, width=14).pack(side="left")
        ttk.Label(date_row, text=" → ").pack(side="left")
        ttk.Entry(date_row, textvariable=self.end_date_var, width=14).pack(side="left")

        ttk.Label(regime_replay_frame, text="Ticker universe source").grid(row=2, column=0, sticky="w", padx=6, pady=4)
        self.regime_ticker_source_var = tk.StringVar(value="Ticker Entry page (0 symbols)")
        ttk.Label(regime_replay_frame, textvariable=self.regime_ticker_source_var, justify="left").grid(row=2, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(regime_replay_frame, text="Starting capital").grid(row=3, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(regime_replay_frame, textvariable=self.starting_capital_var).grid(row=3, column=1, sticky="ew", padx=6, pady=4)

        quick_actions = ttk.Frame(regime_replay_frame)
        quick_actions.grid(row=4, column=0, columnspan=2, sticky="w", padx=6, pady=(2, 4))
        ttk.Button(quick_actions, text="Run exact training-window replay", command=self._run_exact_training_window_replay).pack(side="left", padx=(0, 6))
        ttk.Button(quick_actions, text="Run full-history replay", command=self._run_full_history_replay).pack(side="left")

        ttk.Button(regime_replay_frame, text="Run regime replay", command=self.run_full_chain).grid(row=5, column=1, sticky="e", padx=6, pady=(4, 6))

        self.regime_overrides_expanded_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            regime_replay_frame,
            text="Advanced overrides",
            variable=self.regime_overrides_expanded_var,
            command=self._toggle_regime_advanced_overrides,
        ).grid(row=6, column=0, columnspan=2, sticky="w", padx=6, pady=(0, 4))

        self.regime_advanced_overrides_frame = ttk.Frame(regime_replay_frame)
        self.regime_advanced_overrides_frame.columnconfigure(0, weight=1)
        self.regime_immutable_reason_var = tk.StringVar(value="Immutable fields: select a trained regime to inspect replay constraints.")
        ttk.Label(
            self.regime_advanced_overrides_frame,
            textvariable=self.regime_immutable_reason_var,
            foreground="#225577",
            wraplength=620,
            justify="left",
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=2, pady=(0, 2))

        lock_row = ttk.Frame(self.regime_advanced_overrides_frame)
        lock_row.grid(row=1, column=0, sticky="w", pady=(0, 4))
        self.regime_lock_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(lock_row, text="Lock override fields", variable=self.regime_lock_var, command=self._on_regime_lock_toggled).pack(side="left")
        ttk.Button(lock_row, text="Compatibility troubleshooting", command=self._show_regime_compatibility_details).pack(side="left", padx=(8, 0))

        self.regime_provenance_var = tk.StringVar(value="No trained regime loaded.")
        self.regime_diff_var = tk.StringVar(value="")
        ttk.Label(self.regime_advanced_overrides_frame, textvariable=self.regime_provenance_var, justify="left", foreground="#225577").grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 2))
        ttk.Label(self.regime_advanced_overrides_frame, textvariable=self.regime_diff_var, justify="left", foreground="#995500", wraplength=640).grid(row=3, column=0, columnspan=2, sticky="w", pady=(0, 4))
        self._toggle_regime_advanced_overrides()

        strategy_frame = ttk.LabelFrame(run_setup_tab, text="Strategy")
        strategy_frame.grid(row=3, column=0, sticky="nsew", padx=10, pady=10)
        strategy_frame.columnconfigure(1, weight=1)

        row = 0
        ttk.Label(strategy_frame, text="Strategy").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.strategy_var = tk.StringVar(value="momentum")
        self.strategy_combo = ttk.Combobox(
            strategy_frame,
            textvariable=self.strategy_var,
            state="readonly",
            values=STRATEGIES,
        )
        self.strategy_combo.grid(row=row, column=1, sticky="ew", padx=8, pady=6)
        self.strategy_combo.bind("<<ComboboxSelected>>", self._on_strategy_changed)

        row += 1
        ttk.Label(strategy_frame, text="Lookback (bars)").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.lookback_days_var = tk.StringVar()
        self.lookback_days_entry = ttk.Entry(strategy_frame, textvariable=self.lookback_days_var)
        self.lookback_days_entry.grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="Skip (bars)").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.skip_days_var = tk.StringVar()
        self.skip_days_entry = ttk.Entry(strategy_frame, textvariable=self.skip_days_var)
        self.skip_days_entry.grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="Costs (bps)").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.costs_bps_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.costs_bps_var).grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="Execution Model").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.execution_model_var = tk.StringVar(value="bps")
        ttk.Combobox(
            strategy_frame,
            textvariable=self.execution_model_var,
            state="readonly",
            values=EXECUTION_MODELS,
        ).grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="Spread bps / Max Participation").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        execution_row = ttk.Frame(strategy_frame)
        execution_row.grid(row=row, column=1, sticky="ew", padx=8, pady=6)
        self.execution_spread_bps_var = tk.StringVar(value="2")
        self.execution_max_participation_var = tk.StringVar(value="1.0")
        ttk.Entry(execution_row, textvariable=self.execution_spread_bps_var, width=10).pack(side="left")
        ttk.Label(execution_row, text=" / ").pack(side="left")
        ttk.Entry(execution_row, textvariable=self.execution_max_participation_var, width=10).pack(side="left")

        row += 1
        ttk.Label(strategy_frame, text="Impact bps / Latency bars / Latency ms").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        execution_row2 = ttk.Frame(strategy_frame)
        execution_row2.grid(row=row, column=1, sticky="ew", padx=8, pady=6)
        self.execution_impact_bps_var = tk.StringVar(value="5")
        self.execution_latency_bars_var = tk.StringVar(value="0")
        self.execution_latency_ms_var = tk.StringVar(value="0")
        ttk.Entry(execution_row2, textvariable=self.execution_impact_bps_var, width=8).pack(side="left")
        ttk.Label(execution_row2, text=" / ").pack(side="left")
        ttk.Entry(execution_row2, textvariable=self.execution_latency_bars_var, width=8).pack(side="left")
        ttk.Label(execution_row2, text=" / ").pack(side="left")
        ttk.Entry(execution_row2, textvariable=self.execution_latency_ms_var, width=10).pack(side="left")

        row += 1
        ttk.Label(strategy_frame, text="Stress: Replay/Jump/Overlay").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        stress_row = ttk.Frame(strategy_frame)
        stress_row.grid(row=row, column=1, sticky="ew", padx=8, pady=6)
        self.stress_enable_historical_replay_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(stress_row, text="Historical replay regimes", variable=self.stress_enable_historical_replay_var).pack(side="left")

        row += 1
        ttk.Label(strategy_frame, text="Stress window frac / replay bars").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        stress_row2 = ttk.Frame(strategy_frame)
        stress_row2.grid(row=row, column=1, sticky="ew", padx=8, pady=6)
        self.stress_historical_window_fraction_var = tk.StringVar(value="0.20")
        self.stress_historical_replay_window_bars_var = tk.StringVar(value="20")
        ttk.Entry(stress_row2, textvariable=self.stress_historical_window_fraction_var, width=8).pack(side="left")
        ttk.Label(stress_row2, text=" / ").pack(side="left")
        ttk.Entry(stress_row2, textvariable=self.stress_historical_replay_window_bars_var, width=8).pack(side="left")

        row += 1
        ttk.Label(strategy_frame, text="Stress jump mag / interval / vol cluster").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        stress_row3 = ttk.Frame(strategy_frame)
        stress_row3.grid(row=row, column=1, sticky="ew", padx=8, pady=6)
        self.stress_synthetic_jump_magnitude_var = tk.StringVar(value="0.02")
        self.stress_synthetic_jump_interval_var = tk.StringVar(value="7")
        self.stress_synthetic_vol_cluster_multiplier_var = tk.StringVar(value="1.6")
        ttk.Entry(stress_row3, textvariable=self.stress_synthetic_jump_magnitude_var, width=8).pack(side="left")
        ttk.Label(stress_row3, text=" / ").pack(side="left")
        ttk.Entry(stress_row3, textvariable=self.stress_synthetic_jump_interval_var, width=8).pack(side="left")
        ttk.Label(stress_row3, text=" / ").pack(side="left")
        ttk.Entry(stress_row3, textvariable=self.stress_synthetic_vol_cluster_multiplier_var, width=8).pack(side="left")

        row += 1
        ttk.Label(strategy_frame, text="Stress overlay spread / liquidity").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        stress_row4 = ttk.Frame(strategy_frame)
        stress_row4.grid(row=row, column=1, sticky="ew", padx=8, pady=6)
        self.stress_overlay_spread_multiplier_var = tk.StringVar(value="2.5")
        self.stress_overlay_liquidity_multiplier_var = tk.StringVar(value="0.4")
        ttk.Entry(stress_row4, textvariable=self.stress_overlay_spread_multiplier_var, width=8).pack(side="left")
        ttk.Label(stress_row4, text=" / ").pack(side="left")
        ttk.Entry(stress_row4, textvariable=self.stress_overlay_liquidity_multiplier_var, width=8).pack(side="left")

        row += 1
        ttk.Label(strategy_frame, text="Scenario packs").grid(row=row, column=0, sticky="nw", padx=8, pady=6)
        self.scenario_pack_listbox = tk.Listbox(strategy_frame, selectmode="multiple", exportselection=False, height=4)
        self.scenario_pack_listbox.grid(row=row, column=1, sticky="ew", padx=8, pady=6)
        self._scenario_pack_options = tuple(list_scenario_pack_templates())
        for option in self._scenario_pack_options:
            self.scenario_pack_listbox.insert("end", option)

        row += 1
        ttk.Label(strategy_frame, text="Stress profile").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        profile_row = ttk.Frame(strategy_frame)
        profile_row.grid(row=row, column=1, sticky="w", padx=8, pady=6)
        self.selected_stress_profile_var = tk.StringVar(value="Base")
        for profile_name in ("Mild", "Base", "Severe"):
            ttk.Button(profile_row, text=profile_name, command=lambda p=profile_name: self._apply_stress_profile(p)).pack(side="left", padx=(0, 6))

        row += 1
        self.run_selection_summary_var = tk.StringVar(value="Run summary will include selected scenario packs and stress profile.")
        ttk.Label(strategy_frame, textvariable=self.run_selection_summary_var, foreground="#2a4f7a", justify="left", wraplength=620).grid(
            row=row,
            column=0,
            columnspan=2,
            sticky="w",
            padx=8,
            pady=(0, 6),
        )

        row += 1
        ttk.Label(strategy_frame, text="Starting Capital").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(strategy_frame, textvariable=self.starting_capital_var).grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="Bet Size Mode").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.bet_sizing_mode_var = tk.StringVar()
        ttk.Combobox(
            strategy_frame,
            textvariable=self.bet_sizing_mode_var,
            state="readonly",
            values=["kelly", "half_kelly", "custom"],
        ).grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="Custom Bet %").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.custom_bet_pct_var = tk.StringVar()
        self.custom_bet_pct_entry = ttk.Entry(strategy_frame, textvariable=self.custom_bet_pct_var)
        self.custom_bet_pct_entry.grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="Resolution").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.timeframe_var = tk.StringVar(value="1m")
        ttk.Combobox(
            strategy_frame,
            textvariable=self.timeframe_var,
            state="readonly",
            values=TIMEFRAMES,
        ).grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="Portfolio Method").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.portfolio_method_var = tk.StringVar(value="equal_weight")
        self.portfolio_method_combo = ttk.Combobox(
            strategy_frame,
            textvariable=self.portfolio_method_var,
            state="readonly",
            values=PORTFOLIO_METHODS,
        )
        self.portfolio_method_combo.grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="Portfolio Vol Lookback").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.portfolio_vol_lookback_var = tk.StringVar(value="20")
        self.portfolio_vol_lookback_entry = ttk.Entry(strategy_frame, textvariable=self.portfolio_vol_lookback_var)
        self.portfolio_vol_lookback_entry.grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="Target Volatility").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.portfolio_target_vol_var = tk.StringVar(value="0.10")
        self.portfolio_target_vol_entry = ttk.Entry(strategy_frame, textvariable=self.portfolio_target_vol_var)
        self.portfolio_target_vol_entry.grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="Max Symbol Weight").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.portfolio_max_symbol_var = tk.StringVar(value="0.25")
        self.portfolio_max_symbol_entry = ttk.Entry(strategy_frame, textvariable=self.portfolio_max_symbol_var)
        self.portfolio_max_symbol_entry.grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="Max Sector Weight").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.portfolio_max_sector_var = tk.StringVar(value="0.60")
        self.portfolio_max_sector_entry = ttk.Entry(strategy_frame, textvariable=self.portfolio_max_sector_var)
        self.portfolio_max_sector_entry.grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="Rebalance Frequency (bars)").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.portfolio_rebalance_frequency_var = tk.StringVar(value="1")
        self.portfolio_rebalance_frequency_entry = ttk.Entry(strategy_frame, textvariable=self.portfolio_rebalance_frequency_var)
        self.portfolio_rebalance_frequency_entry.grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="Cluster Linkage").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.portfolio_clustering_linkage_var = tk.StringVar(value="single")
        self.portfolio_clustering_linkage_combo = ttk.Combobox(
            strategy_frame,
            textvariable=self.portfolio_clustering_linkage_var,
            state="readonly",
            values=["single", "complete", "average", "ward"],
        )
        self.portfolio_clustering_linkage_combo.grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="Covariance Shrinkage Strength").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.portfolio_covariance_shrinkage_var = tk.StringVar(value="0.15")
        self.portfolio_covariance_shrinkage_entry = ttk.Entry(strategy_frame, textvariable=self.portfolio_covariance_shrinkage_var)
        self.portfolio_covariance_shrinkage_entry.grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="Max Gross Exposure").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.portfolio_max_gross_var = tk.StringVar(value="1.0")
        self.portfolio_max_gross_entry = ttk.Entry(strategy_frame, textvariable=self.portfolio_max_gross_var)
        self.portfolio_max_gross_entry.grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="Min Net Exposure").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.portfolio_min_net_var = tk.StringVar(value="-1.0")
        self.portfolio_min_net_entry = ttk.Entry(strategy_frame, textvariable=self.portfolio_min_net_var)
        self.portfolio_min_net_entry.grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="Max Net Exposure").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.portfolio_max_net_var = tk.StringVar(value="1.0")
        self.portfolio_max_net_entry = ttk.Entry(strategy_frame, textvariable=self.portfolio_max_net_var)
        self.portfolio_max_net_entry.grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        self.portfolio_max_net_gamma_var = tk.StringVar(value="")
        self.portfolio_max_abs_vega_bucket_var = tk.StringVar(value="")
        self.portfolio_max_abs_delta_underlying_var = tk.StringVar(value="")
        self.portfolio_max_participation_rate_var = tk.StringVar(value="")

        risk_limits_card = ttk.LabelFrame(strategy_frame, text="Options Risk Limits")
        risk_limits_card.grid(row=row, column=0, columnspan=2, sticky="ew", padx=8, pady=6)
        risk_limits_card.columnconfigure(1, weight=1)

        self._risk_limit_slider_vars: dict[str, tk.DoubleVar] = {
            "portfolio_max_net_gamma": tk.DoubleVar(value=0.0),
            "portfolio_max_abs_vega_bucket": tk.DoubleVar(value=0.0),
            "portfolio_max_abs_delta_per_underlying": tk.DoubleVar(value=0.0),
            "max_participation_rate": tk.DoubleVar(value=0.0),
        }
        self._risk_limit_vars: dict[str, tk.StringVar] = {
            "portfolio_max_net_gamma": self.portfolio_max_net_gamma_var,
            "portfolio_max_abs_vega_bucket": self.portfolio_max_abs_vega_bucket_var,
            "portfolio_max_abs_delta_per_underlying": self.portfolio_max_abs_delta_underlying_var,
            "max_participation_rate": self.portfolio_max_participation_rate_var,
        }
        self._risk_limit_ranges: dict[str, tuple[float, float, float]] = {
            "portfolio_max_net_gamma": (0.0, 5.0, 0.05),
            "portfolio_max_abs_vega_bucket": (0.0, 50_000.0, 250.0),
            "portfolio_max_abs_delta_per_underlying": (0.0, 10_000.0, 50.0),
            "max_participation_rate": (0.0, 1.0, 0.01),
        }

        self._build_risk_limit_slider_row(risk_limits_card, 0, "Max Net Gamma", "portfolio_max_net_gamma")
        self._build_risk_limit_slider_row(risk_limits_card, 1, "Max Abs Vega Bucket", "portfolio_max_abs_vega_bucket")
        self._build_risk_limit_slider_row(risk_limits_card, 2, "Max Abs Delta / Underlying", "portfolio_max_abs_delta_per_underlying")
        self._build_risk_limit_slider_row(risk_limits_card, 3, "Max Participation / Capacity Threshold", "max_participation_rate")

        presets_row = ttk.Frame(risk_limits_card)
        presets_row.grid(row=4, column=0, columnspan=3, sticky="w", padx=8, pady=(2, 6))
        ttk.Label(presets_row, text="Quick presets:").pack(side="left")
        ttk.Button(presets_row, text="Conservative", command=lambda: self._apply_options_risk_limit_preset("conservative")).pack(side="left", padx=(6, 0))
        ttk.Button(presets_row, text="Balanced", command=lambda: self._apply_options_risk_limit_preset("balanced")).pack(side="left", padx=(6, 0))
        ttk.Button(presets_row, text="Aggressive", command=lambda: self._apply_options_risk_limit_preset("aggressive")).pack(side="left", padx=(6, 0))

        row += 1
        self.use_walk_forward_var = tk.BooleanVar(value=False)
        self.use_optimizer_var = tk.BooleanVar(value=False)
        self.use_walk_forward_check = ttk.Checkbutton(
            strategy_frame,
            text="Use Walk-Forward (Momentum)",
            variable=self.use_walk_forward_var,
            command=self._update_validation_hint,
        )
        self.use_walk_forward_check.grid(row=row, column=0, columnspan=2, sticky="w", padx=8, pady=6)

        row += 1
        self.use_optimizer_check = ttk.Checkbutton(
            strategy_frame,
            text="Run multi-objective optimizer (Momentum)",
            variable=self.use_optimizer_var,
            command=self._update_validation_hint,
        )
        self.use_optimizer_check.grid(row=row, column=0, columnspan=2, sticky="w", padx=8, pady=6)

        row += 1
        ttk.Label(
            strategy_frame,
            text="Walk-forward tunes on train+validation, then evaluates only on out-of-sample test folds. Optimizer runs constrained multi-objective search with configurable samplers and pruning.",
            wraplength=520,
            justify="left",
        ).grid(row=row, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 6))

        row += 1
        self.optimizer_sampler_var = tk.StringVar(value="tpe")
        self.optimizer_trials_var = tk.StringVar(value="20")
        ttk.Label(strategy_frame, text="Optimizer sampler / trials").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        optimizer_row = ttk.Frame(strategy_frame)
        optimizer_row.grid(row=row, column=1, sticky="w", padx=8, pady=6)
        ttk.Combobox(optimizer_row, textvariable=self.optimizer_sampler_var, state="readonly", values=OPTIMIZER_SAMPLERS, width=10).pack(side="left")
        ttk.Label(optimizer_row, text=" / ").pack(side="left")
        ttk.Entry(optimizer_row, textvariable=self.optimizer_trials_var, width=8).pack(side="left")

        row += 1
        self.optimizer_enable_pruning_var = tk.BooleanVar(value=True)
        self.optimizer_prune_constraint_var = tk.BooleanVar(value=True)
        self.optimizer_prune_lcb_var = tk.BooleanVar(value=True)
        self.optimizer_min_completed_var = tk.StringVar(value="5")
        ttk.Label(strategy_frame, text="Pruning controls").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        prune_row = ttk.Frame(strategy_frame)
        prune_row.grid(row=row, column=1, sticky="w", padx=8, pady=6)
        ttk.Checkbutton(prune_row, text="Enable", variable=self.optimizer_enable_pruning_var).pack(side="left")
        ttk.Checkbutton(prune_row, text="Constraint", variable=self.optimizer_prune_constraint_var).pack(side="left", padx=(8,0))
        ttk.Checkbutton(prune_row, text="LCB", variable=self.optimizer_prune_lcb_var).pack(side="left", padx=(8,0))
        ttk.Label(prune_row, text="min done").pack(side="left", padx=(8,0))
        ttk.Entry(prune_row, textvariable=self.optimizer_min_completed_var, width=5).pack(side="left")

        row += 1
        self.optimizer_staged_budgets_var = tk.StringVar(value='[{"label":"coarse","n_trials":12,"sampler":"random","partial_period_fractions":[0.33,0.66]},{"label":"fine","n_trials":20,"sampler":"tpe","partial_period_fractions":[0.5,1.0]}]')
        ttk.Label(strategy_frame, text="Staged budgets (JSON)").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(strategy_frame, textvariable=self.optimizer_staged_budgets_var).grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        self.show_advanced_optimization_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            strategy_frame,
            text="Advanced Optimization",
            variable=self.show_advanced_optimization_var,
            command=self._toggle_advanced_optimization,
        ).grid(row=row, column=0, columnspan=2, sticky="w", padx=8, pady=6)

        row += 1
        self.advanced_optimization_frame = ttk.LabelFrame(strategy_frame, text="Advanced Optimization")
        self.advanced_optimization_frame.grid(row=row, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 6))
        self.advanced_optimization_frame.columnconfigure(1, weight=1)

        self.optimizer_search_space_var = tk.StringVar(value=DEFAULT_OPTIMIZER_SEARCH_SPACE)
        self.optimizer_objectives_var = tk.StringVar(value=DEFAULT_OPTIMIZER_OBJECTIVES)
        self.optimizer_max_turnover_var = tk.StringVar(value="")
        self.optimizer_max_drawdown_floor_var = tk.StringVar(value="")
        self.optimizer_min_trades_var = tk.StringVar(value="")
        self.optimizer_objective_weights_var = tk.StringVar(value=DEFAULT_OPTIMIZER_OBJECTIVE_WEIGHTS)
        self.optimizer_overfitting_penalty_var = tk.StringVar(value=DEFAULT_OPTIMIZER_OVERFITTING_PENALTY)

        self._build_optimizer_json_input(self.advanced_optimization_frame, 0, "Search space (JSON)", self.optimizer_search_space_var, DEFAULT_OPTIMIZER_SEARCH_SPACE, "search_space")
        self._build_optimizer_json_input(self.advanced_optimization_frame, 2, "Objectives (JSON)", self.optimizer_objectives_var, DEFAULT_OPTIMIZER_OBJECTIVES, "objectives")

        ttk.Label(self.advanced_optimization_frame, text="Max turnover / drawdown floor / min trades").grid(row=4, column=0, sticky="w", padx=8, pady=6)
        optimizer_constraints_row = ttk.Frame(self.advanced_optimization_frame)
        optimizer_constraints_row.grid(row=4, column=1, sticky="w", padx=8, pady=6)
        ttk.Entry(optimizer_constraints_row, textvariable=self.optimizer_max_turnover_var, width=8).pack(side="left")
        ttk.Label(optimizer_constraints_row, text=" / ").pack(side="left")
        ttk.Entry(optimizer_constraints_row, textvariable=self.optimizer_max_drawdown_floor_var, width=8).pack(side="left")
        ttk.Label(optimizer_constraints_row, text=" / ").pack(side="left")
        ttk.Entry(optimizer_constraints_row, textvariable=self.optimizer_min_trades_var, width=8).pack(side="left")

        self._build_optimizer_json_input(self.advanced_optimization_frame, 5, "Objective weights (JSON)", self.optimizer_objective_weights_var, DEFAULT_OPTIMIZER_OBJECTIVE_WEIGHTS, "objective_weights")
        self._build_optimizer_json_input(self.advanced_optimization_frame, 7, "Overfitting penalty (JSON)", self.optimizer_overfitting_penalty_var, DEFAULT_OPTIMIZER_OVERFITTING_PENALTY, "overfitting_penalty")
        self._toggle_advanced_optimization()

        row += 1
        self.walk_forward_frame = ttk.LabelFrame(strategy_frame, text="Walk-Forward Windows (fractions or bars)")
        self.walk_forward_frame.grid(row=row, column=0, columnspan=2, sticky="ew", padx=8, pady=6)
        self.walk_forward_frame.columnconfigure(1, weight=1)

        ttk.Label(self.walk_forward_frame, text="Train Fraction").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        self.wf_train_fraction_var = tk.DoubleVar(value=0.70)
        self.wf_train_fraction_label = ttk.Label(self.walk_forward_frame, text="0.70")
        self.wf_train_scale = tk.Scale(
            self.walk_forward_frame,
            from_=0.05,
            to=0.90,
            resolution=0.01,
            orient="horizontal",
            variable=self.wf_train_fraction_var,
            command=lambda _value: self._on_wf_fraction_changed("train"),
        )
        self.wf_train_scale.grid(row=0, column=1, sticky="ew", padx=8, pady=6)
        self.wf_train_fraction_label.grid(row=0, column=2, sticky="e", padx=8, pady=6)

        ttk.Label(self.walk_forward_frame, text="Validation Fraction").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        self.wf_validation_fraction_var = tk.DoubleVar(value=0.15)
        self.wf_validation_fraction_label = ttk.Label(self.walk_forward_frame, text="0.15")
        self.wf_validation_scale = tk.Scale(
            self.walk_forward_frame,
            from_=0.05,
            to=0.90,
            resolution=0.01,
            orient="horizontal",
            variable=self.wf_validation_fraction_var,
            command=lambda _value: self._on_wf_fraction_changed("validation"),
        )
        self.wf_validation_scale.grid(row=1, column=1, sticky="ew", padx=8, pady=6)
        self.wf_validation_fraction_label.grid(row=1, column=2, sticky="e", padx=8, pady=6)

        ttk.Label(self.walk_forward_frame, text="Test Fraction").grid(row=2, column=0, sticky="w", padx=8, pady=6)
        self.wf_test_fraction_var = tk.DoubleVar(value=0.15)
        self.wf_test_fraction_label = ttk.Label(self.walk_forward_frame, text="0.15")
        self.wf_test_scale = tk.Scale(
            self.walk_forward_frame,
            from_=0.05,
            to=0.90,
            resolution=0.01,
            orient="horizontal",
            variable=self.wf_test_fraction_var,
            command=lambda _value: self._on_wf_fraction_changed("test"),
        )
        self.wf_test_scale.grid(row=2, column=1, sticky="ew", padx=8, pady=6)
        self.wf_test_fraction_label.grid(row=2, column=2, sticky="e", padx=8, pady=6)

        ttk.Label(self.walk_forward_frame, text="Step Fraction").grid(row=3, column=0, sticky="w", padx=8, pady=6)
        self.wf_step_fraction_var = tk.DoubleVar(value=0.15)
        self.wf_step_fraction_label = ttk.Label(self.walk_forward_frame, text="0.15")
        self.wf_step_scale = tk.Scale(
            self.walk_forward_frame,
            from_=0.05,
            to=1.00,
            resolution=0.01,
            orient="horizontal",
            variable=self.wf_step_fraction_var,
            command=lambda _value: self._refresh_wf_fraction_labels(),
        )
        self.wf_step_scale.grid(row=3, column=1, sticky="ew", padx=8, pady=6)
        self.wf_step_fraction_label.grid(row=3, column=2, sticky="e", padx=8, pady=6)

        ttk.Label(self.walk_forward_frame, text="Train/Validation/Test/Step Bars").grid(row=4, column=0, sticky="w", padx=8, pady=6)
        bars_row = ttk.Frame(self.walk_forward_frame)
        bars_row.grid(row=4, column=1, sticky="ew", padx=8, pady=6)
        self.wf_train_bars_var = tk.StringVar(value="")
        self.wf_validation_bars_var = tk.StringVar(value="")
        self.wf_test_bars_var = tk.StringVar(value="")
        self.wf_step_bars_var = tk.StringVar(value="")
        ttk.Entry(bars_row, textvariable=self.wf_train_bars_var, width=7).pack(side="left")
        ttk.Label(bars_row, text=" / ").pack(side="left")
        ttk.Entry(bars_row, textvariable=self.wf_validation_bars_var, width=7).pack(side="left")
        ttk.Label(bars_row, text=" / ").pack(side="left")
        ttk.Entry(bars_row, textvariable=self.wf_test_bars_var, width=7).pack(side="left")
        ttk.Label(bars_row, text=" / ").pack(side="left")
        ttk.Entry(bars_row, textvariable=self.wf_step_bars_var, width=7).pack(side="left")

        self.wf_mode_hint_var = tk.StringVar(value="Use either fractions OR bars; leave the other mode blank.")
        ttk.Label(self.walk_forward_frame, textvariable=self.wf_mode_hint_var, foreground="#995500", justify="left").grid(row=5, column=1, sticky="w", padx=8, pady=(0, 6))

        ttk.Label(self.walk_forward_frame, text="CV Scheme").grid(row=6, column=0, sticky="w", padx=8, pady=6)
        self.wf_cv_scheme_var = tk.StringVar(value="walk_forward")
        self.wf_cv_scheme_combo = ttk.Combobox(self.walk_forward_frame, textvariable=self.wf_cv_scheme_var, state="readonly", values=["walk_forward", "cpcv"])
        self.wf_cv_scheme_combo.grid(row=6, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(self.walk_forward_frame, text="Purge/Embargo Bars").grid(row=7, column=0, sticky="w", padx=8, pady=6)
        pe_row = ttk.Frame(self.walk_forward_frame)
        pe_row.grid(row=7, column=1, sticky="ew", padx=8, pady=6)
        self.wf_purge_bars_var = tk.StringVar(value="0")
        self.wf_embargo_bars_var = tk.StringVar(value="0")
        ttk.Entry(pe_row, textvariable=self.wf_purge_bars_var, width=8).pack(side="left")
        ttk.Label(pe_row, text=" / ").pack(side="left")
        ttk.Entry(pe_row, textvariable=self.wf_embargo_bars_var, width=8).pack(side="left")

        ttk.Label(self.walk_forward_frame, text="CPCV Groups/Test Groups").grid(row=8, column=0, sticky="w", padx=8, pady=6)
        cpcv_row = ttk.Frame(self.walk_forward_frame)
        cpcv_row.grid(row=8, column=1, sticky="ew", padx=8, pady=6)
        self.wf_cpcv_groups_var = tk.StringVar(value="6")
        self.wf_cpcv_test_groups_var = tk.StringVar(value="2")
        ttk.Entry(cpcv_row, textvariable=self.wf_cpcv_groups_var, width=8).pack(side="left")
        ttk.Label(cpcv_row, text=" / ").pack(side="left")
        ttk.Entry(cpcv_row, textvariable=self.wf_cpcv_test_groups_var, width=8).pack(side="left")

        ttk.Label(self.walk_forward_frame, text="CV Seed").grid(row=9, column=0, sticky="w", padx=8, pady=6)
        self.wf_cv_seed_var = tk.StringVar(value="42")
        ttk.Entry(self.walk_forward_frame, textvariable=self.wf_cv_seed_var).grid(row=9, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(self.walk_forward_frame, text="Label Horizon Bars").grid(row=10, column=0, sticky="w", padx=8, pady=6)
        self.wf_label_horizon_bars_var = tk.StringVar(value="1")
        ttk.Entry(self.walk_forward_frame, textvariable=self.wf_label_horizon_bars_var).grid(row=10, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(self.walk_forward_frame, text="Nested Optimization / Inner Train Fraction").grid(row=11, column=0, sticky="w", padx=8, pady=6)
        nested_row = ttk.Frame(self.walk_forward_frame)
        nested_row.grid(row=11, column=1, sticky="w", padx=8, pady=6)
        self.wf_nested_optimization_var = tk.BooleanVar(value=False)
        self.wf_inner_train_fraction_var = tk.StringVar(value="0.70")
        ttk.Checkbutton(nested_row, text="Enable", variable=self.wf_nested_optimization_var).pack(side="left")
        ttk.Label(nested_row, text=" / ").pack(side="left")
        ttk.Entry(nested_row, textvariable=self.wf_inner_train_fraction_var, width=8).pack(side="left")

        self.wf_objective_weights_var = tk.StringVar(value="")
        self.wf_overfitting_penalty_var = tk.StringVar(value="")
        ttk.Label(self.walk_forward_frame, text="WF Objective Weights (optional JSON)").grid(row=12, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(self.walk_forward_frame, textvariable=self.wf_objective_weights_var).grid(row=12, column=1, sticky="ew", padx=8, pady=6)
        ttk.Label(self.walk_forward_frame, text="WF Overfitting Penalty (optional JSON)").grid(row=13, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(self.walk_forward_frame, textvariable=self.wf_overfitting_penalty_var).grid(row=13, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(self.walk_forward_frame, text="Strategy Key / Prior Strategy Keys (CSV)").grid(row=14, column=0, sticky="w", padx=8, pady=6)
        lineage_row = ttk.Frame(self.walk_forward_frame)
        lineage_row.grid(row=14, column=1, sticky="ew", padx=8, pady=6)
        self.wf_strategy_key_var = tk.StringVar(value="")
        self.wf_prior_strategy_keys_var = tk.StringVar(value="")
        ttk.Entry(lineage_row, textvariable=self.wf_strategy_key_var, width=18).pack(side="left")
        ttk.Label(lineage_row, text=" / ").pack(side="left")
        ttk.Entry(lineage_row, textvariable=self.wf_prior_strategy_keys_var, width=22).pack(side="left")

        row += 1
        self.strategy_specific_container = ttk.Frame(strategy_frame)
        self.strategy_specific_container.grid(row=row, column=0, columnspan=2, sticky="ew", padx=0, pady=0)
        self.strategy_specific_container.columnconfigure(0, weight=1)

        self._build_momentum_options(self.strategy_specific_container)
        self._build_xsmom_options(self.strategy_specific_container)

        row += 1
        ttk.Label(strategy_frame, text="Start Date (YYYY-MM-DD)").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(strategy_frame, textvariable=self.start_date_var).grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="End Date (YYYY-MM-DD)").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(strategy_frame, textvariable=self.end_date_var).grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="Backtest Data Root").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.backtest_root_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.backtest_root_var).grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        self.validation_hint_var = tk.StringVar(value="")
        self.validation_hint_label = ttk.Label(
            strategy_frame,
            textvariable=self.validation_hint_var,
            foreground="#995500",
            wraplength=620,
            justify="left",
        )
        self.validation_hint_label.grid(row=row, column=0, columnspan=2, sticky="w", padx=8, pady=(4, 8))

        row += 1
        template_row = ttk.Frame(strategy_frame)
        template_row.grid(row=row, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 6))
        template_row.columnconfigure(1, weight=1)
        ttk.Label(template_row, text="Experiment Template").grid(row=0, column=0, sticky="w")
        self.template_var = tk.StringVar(value="")
        self.template_combo = ttk.Combobox(template_row, textvariable=self.template_var, state="normal", values=[])
        self.template_combo.grid(row=0, column=1, sticky="ew", padx=(8, 8))
        ttk.Button(template_row, text="Load Template", command=self.load_template).grid(row=0, column=2, padx=(0, 6))
        ttk.Button(template_row, text="Save as Experiment Template", command=self.save_template).grid(row=0, column=3)

        governance_frame = ttk.LabelFrame(run_setup_tab, text="Governance & Promotion")
        governance_frame.grid(row=4, column=0, sticky="nsew", padx=10, pady=(4, 8))
        governance_frame.columnconfigure(1, weight=1)

        g_row = 0
        ttk.Label(governance_frame, text="Hypothesis ID").grid(row=g_row, column=0, sticky="w", padx=8, pady=4)
        self.gov_hypothesis_id_var = tk.StringVar(value="")
        ttk.Entry(governance_frame, textvariable=self.gov_hypothesis_id_var).grid(row=g_row, column=1, sticky="ew", padx=8, pady=4)

        g_row += 1
        ttk.Label(governance_frame, text="Experiment ID (required for deployment)").grid(row=g_row, column=0, sticky="w", padx=8, pady=4)
        self.gov_experiment_id_var = tk.StringVar(value="")
        ttk.Entry(governance_frame, textvariable=self.gov_experiment_id_var).grid(row=g_row, column=1, sticky="ew", padx=8, pady=4)

        g_row += 1
        ttk.Label(governance_frame, text="Owner").grid(row=g_row, column=0, sticky="w", padx=8, pady=4)
        self.gov_owner_var = tk.StringVar(value="")
        ttk.Entry(governance_frame, textvariable=self.gov_owner_var).grid(row=g_row, column=1, sticky="ew", padx=8, pady=4)

        g_row += 1
        ttk.Label(governance_frame, text="Dataset Snapshot Lock").grid(row=g_row, column=0, sticky="w", padx=8, pady=4)
        self.gov_dataset_lock_var = tk.StringVar(value="")
        ttk.Entry(governance_frame, textvariable=self.gov_dataset_lock_var).grid(row=g_row, column=1, sticky="ew", padx=8, pady=4)

        g_row += 1
        ttk.Label(governance_frame, text="Acceptance Criteria").grid(row=g_row, column=0, sticky="nw", padx=8, pady=4)
        self.gov_acceptance_text = tk.Text(governance_frame, height=3)
        self.gov_acceptance_text.grid(row=g_row, column=1, sticky="ew", padx=8, pady=4)

        g_row += 1
        ttk.Label(governance_frame, text="Promotion State").grid(row=g_row, column=0, sticky="w", padx=8, pady=4)
        self.gov_promotion_state_var = tk.StringVar(value="research")
        ttk.Combobox(governance_frame, textvariable=self.gov_promotion_state_var, state="readonly", values=GOVERNANCE_PROMOTION_STATES).grid(row=g_row, column=1, sticky="ew", padx=8, pady=4)

        g_row += 1
        ttk.Label(governance_frame, text="Approval Status").grid(row=g_row, column=0, sticky="w", padx=8, pady=4)
        self.gov_approval_status_var = tk.StringVar(value="pending")
        ttk.Combobox(governance_frame, textvariable=self.gov_approval_status_var, state="readonly", values=GOVERNANCE_APPROVAL_STATES).grid(row=g_row, column=1, sticky="ew", padx=8, pady=4)

        g_row += 1
        note_row = ttk.Frame(governance_frame)
        note_row.grid(row=g_row, column=0, columnspan=2, sticky="ew", padx=8, pady=4)
        note_row.columnconfigure(3, weight=1)
        ttk.Label(note_row, text="Gov note owner").grid(row=0, column=0, sticky="w")
        self.gov_note_owner_var = tk.StringVar(value="")
        ttk.Entry(note_row, textvariable=self.gov_note_owner_var, width=16).grid(row=0, column=1, sticky="w", padx=(6, 10))
        ttk.Label(note_row, text="Gov note").grid(row=0, column=2, sticky="w")
        self.gov_note_text_var = tk.StringVar(value="")
        ttk.Entry(note_row, textvariable=self.gov_note_text_var).grid(row=0, column=3, sticky="ew", padx=(6, 10))
        ttk.Button(note_row, text="Append Governance Note", command=self._append_governance_note).grid(row=0, column=4, sticky="e")

        g_row += 1
        gate_row = ttk.Frame(governance_frame)
        gate_row.grid(row=g_row, column=0, columnspan=2, sticky="ew", padx=8, pady=4)
        for idx, (label, default) in enumerate([
            ("Min OOS Periods", "3"),
            ("Min Stability", "0.55"),
            ("Max Turnover", "4.0"),
            ("Min Capacity", "0.5"),
            ("Max Signal Drift", "0.10"),
            ("Max Fill Drift (bps)", "5.0"),
            ("Max PnL Div", "0.15"),
        ]):
            ttk.Label(gate_row, text=label).grid(row=0, column=idx * 2, sticky="w", padx=(0, 4))
            var = tk.StringVar(value=default)
            entry = ttk.Entry(gate_row, textvariable=var, width=8)
            entry.grid(row=0, column=idx * 2 + 1, sticky="w", padx=(0, 10))
            if label == "Min OOS Periods":
                self.gov_min_oos_periods_var = var
            elif label == "Min Stability":
                self.gov_min_stability_var = var
                self.gov_min_stability_entry = entry
            elif label == "Max Turnover":
                self.gov_max_turnover_var = var
            elif label == "Min Capacity":
                self.gov_min_capacity_var = var
            elif label == "Max Signal Drift":
                self.gov_max_signal_agreement_drift_var = var
                self.gov_max_signal_agreement_drift_entry = entry
            elif label == "Max Fill Drift (bps)":
                self.gov_max_fill_slippage_drift_bps_var = var
            else:
                self.gov_max_pnl_attribution_divergence_var = var

        g_row += 1
        ttk.Label(governance_frame, text="Expected signal/fill/pnl").grid(row=g_row, column=0, sticky="w", padx=8, pady=4)
        expected_row = ttk.Frame(governance_frame)
        expected_row.grid(row=g_row, column=1, sticky="ew", padx=8, pady=4)
        self.gov_expected_signal_agreement_var = tk.StringVar(value="1.0")
        self.gov_expected_fill_slippage_bps_var = tk.StringVar(value="0.0")
        self.gov_expected_pnl_attribution_var = tk.StringVar(value="1.0")
        self.gov_expected_signal_agreement_entry = ttk.Entry(expected_row, textvariable=self.gov_expected_signal_agreement_var, width=8)
        self.gov_expected_signal_agreement_entry.pack(side="left")
        ttk.Label(expected_row, text=" / ").pack(side="left")
        ttk.Entry(expected_row, textvariable=self.gov_expected_fill_slippage_bps_var, width=8).pack(side="left")
        ttk.Label(expected_row, text=" / ").pack(side="left")
        ttk.Entry(expected_row, textvariable=self.gov_expected_pnl_attribution_var, width=8).pack(side="left")

        g_row += 1
        ttk.Label(governance_frame, text="Observed signal/fill/pnl").grid(row=g_row, column=0, sticky="w", padx=8, pady=4)
        observed_row = ttk.Frame(governance_frame)
        observed_row.grid(row=g_row, column=1, sticky="ew", padx=8, pady=4)
        self.gov_observed_signal_agreement_var = tk.StringVar(value="1.0")
        self.gov_observed_fill_slippage_bps_var = tk.StringVar(value="0.0")
        self.gov_observed_pnl_attribution_var = tk.StringVar(value="1.0")
        ttk.Entry(observed_row, textvariable=self.gov_observed_signal_agreement_var, width=8).pack(side="left")
        ttk.Label(observed_row, text=" / ").pack(side="left")
        ttk.Entry(observed_row, textvariable=self.gov_observed_fill_slippage_bps_var, width=8).pack(side="left")
        ttk.Label(observed_row, text=" / ").pack(side="left")
        ttk.Entry(observed_row, textvariable=self.gov_observed_pnl_attribution_var, width=8).pack(side="left")

        self._governance_comments: list[dict[str, str]] = []
        self._governance_review_actions: list[dict[str, str]] = []
        self._governance_decision_log: list[dict[str, str]] = []

        notes_frame = ttk.LabelFrame(run_setup_tab, text="Run Notes")
        notes_frame.grid(row=5, column=0, sticky="nsew", padx=10, pady=10)
        notes_frame.columnconfigure(0, weight=1)
        self.notes_text = tk.Text(notes_frame, height=10)
        self.notes_text.grid(row=0, column=0, sticky="nsew", padx=8, pady=6)
        self.governance_log_text = tk.Text(notes_frame, height=6)
        self.governance_log_text.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 6))
        self.governance_log_text.configure(state="disabled")

        button_row = ttk.Frame(run_setup_tab)
        button_row.grid(row=6, column=0, sticky="ew", padx=10, pady=(4, 10))
        button_row.columnconfigure(0, weight=1)
        button_row.columnconfigure(1, weight=1)
        button_row.columnconfigure(2, weight=1)
        button_row.columnconfigure(3, weight=1)
        button_row.columnconfigure(4, weight=1)

        ttk.Button(button_row, text="Save Parameters", command=self.save_settings).grid(row=0, column=0, padx=10, pady=4)
        self.run_stress_only_button = ttk.Button(button_row, text="Run stress only", command=self.run_stress_only)
        self.run_stress_only_button.grid(row=0, column=1, padx=10, pady=4)
        self.run_full_chain_button = ttk.Button(button_row, text="Run full chain", command=self.run_full_chain)
        self.run_full_chain_button.grid(row=0, column=2, padx=10, pady=4)
        self.run_button = self.run_full_chain_button
        ttk.Button(button_row, text="Export Prompt Pack", command=self.export_prompt_pack).grid(row=0, column=3, padx=10, pady=4)
        ttk.Button(
            button_row,
            text="Back to Main Menu",
            command=lambda: controller.show_frame("MainMenu"),
        ).grid(row=0, column=4, padx=10, pady=4)

        self._register_advanced_widgets()
        self._bind_validation_watchers()
        self._bind_regime_override_watchers()
        self._build_results_tabs()
        self._on_strategy_changed()
        self._on_mode_changed()
        self._refresh_template_choices()
        self._refresh_trained_regime_choices()
        self._update_validation_hint()

    def _append_governance_note(self) -> None:
        note = self.gov_note_text_var.get().strip()
        if not note:
            return
        owner = self.gov_note_owner_var.get().strip() or self.gov_owner_var.get().strip() or "research_lab_ui"
        timestamp = datetime.now().isoformat(timespec="seconds")
        comment = {"owner": owner, "note": note, "timestamp": timestamp}
        self._governance_comments.append(comment)
        self._governance_review_actions.append(
            {"owner": owner, "action": "note_appended", "status": "recorded", "timestamp": timestamp}
        )
        self.gov_note_text_var.set("")
        self._render_governance_log()

    def _render_governance_log(self) -> None:
        if not hasattr(self, "governance_log_text"):
            return
        self.governance_log_text.configure(state="normal")
        self.governance_log_text.delete("1.0", tk.END)
        for row in self._governance_comments:
            self.governance_log_text.insert(
                tk.END,
                f"[{row.get('timestamp', '')}] {row.get('owner', 'research_lab_ui')}: {row.get('note', '')}\n",
            )
        self.governance_log_text.configure(state="disabled")

    def _on_page_canvas_configure(self, event: tk.Event) -> None:
        self.page_canvas.itemconfigure(self._page_canvas_window, width=event.width)

    def _on_page_frame_configure(self, _event: tk.Event) -> None:
        self.page_canvas.configure(scrollregion=self.page_canvas.bbox("all"))

    def _bind_mousewheel(self, _event: tk.Event) -> None:
        self.page_canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.page_canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.page_canvas.bind_all("<Button-5>", self._on_mousewheel)

    def _unbind_mousewheel(self, _event: tk.Event) -> None:
        self.page_canvas.unbind_all("<MouseWheel>")
        self.page_canvas.unbind_all("<Button-4>")
        self.page_canvas.unbind_all("<Button-5>")

    def _on_mousewheel(self, event: tk.Event) -> None:
        if getattr(event, "num", None) == 4:
            self.page_canvas.yview_scroll(-1, "units")
            return
        if getattr(event, "num", None) == 5:
            self.page_canvas.yview_scroll(1, "units")
            return
        delta = getattr(event, "delta", 0)
        if delta:
            self.page_canvas.yview_scroll(int(-delta / 120), "units")

    def _build_results_tabs(self) -> None:
        self.current_run_dirs: list[Path] = []
        self.current_fold_dirs: list[Path] = []
        self._run_dir_to_label: dict[Path, str] = {}
        self._last_output_text = ""
        self._run_manifest_cache: dict[Path, dict[str, object]] = {}
        self._run_metrics_cache: dict[Path, dict[str, float]] = {}

        leaderboard_tab = ttk.Frame(self.section_notebook)
        leaderboard_tab.columnconfigure(0, weight=1)
        leaderboard_tab.rowconfigure(3, weight=1)
        self.section_notebook.add(leaderboard_tab, text="Results Leaderboard")

        controls = ttk.Frame(leaderboard_tab)
        controls.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 4))
        controls.columnconfigure(2, weight=1)
        ttk.Label(controls, text="Detected Run Artifacts").grid(row=0, column=0, sticky="w")
        ttk.Button(controls, text="Refresh", command=self._refresh_artifacts_view).grid(row=0, column=1, padx=(8, 0))

        self.run_listbox = tk.Listbox(leaderboard_tab, selectmode="extended", exportselection=False, height=5)
        self.run_listbox.grid(row=1, column=0, sticky="ew", padx=10, pady=4)
        self._selection_job: str | None = None
        self.run_listbox.bind("<<ListboxSelect>>", self._on_run_selection_event)

        compare_row = ttk.Frame(leaderboard_tab)
        compare_row.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 4))
        ttk.Button(compare_row, text="Compare Selected Runs", command=self._on_compare_selection_changed).pack(side="left")
        self.guardrail_frame = ttk.Frame(compare_row)
        self.guardrail_frame.pack(side="right", fill="x", expand=True)

        self.leaderboard_tree = ttk.Treeview(leaderboard_tab, show="headings", height=10)
        self.leaderboard_tree.grid(row=3, column=0, sticky="nsew", padx=10, pady=4)
        self.leaderboard_tree_scroll = ttk.Scrollbar(leaderboard_tab, orient="vertical", command=self.leaderboard_tree.yview)
        self.leaderboard_tree_scroll.grid(row=3, column=1, sticky="ns", pady=4)
        self.leaderboard_tree.configure(yscrollcommand=self.leaderboard_tree_scroll.set)
        leaderboard_tab.rowconfigure(3, weight=1)

        self.delta_summary_var = tk.StringVar(value="Select at least two runs for delta metrics.")
        ttk.Label(leaderboard_tab, textvariable=self.delta_summary_var, justify="left").grid(
            row=4, column=0, sticky="ew", padx=10, pady=(4, 8)
        )

        inplace_compare_frame = ttk.Labelframe(leaderboard_tab, text="Selected Pair Delta Snapshot")
        inplace_compare_frame.grid(row=5, column=0, sticky="nsew", padx=10, pady=(0, 8))
        inplace_compare_frame.columnconfigure(0, weight=1)
        inplace_compare_frame.rowconfigure(0, weight=1)
        self.inplace_compare_tree = ttk.Treeview(inplace_compare_frame, show="headings", height=6)
        self.inplace_compare_tree.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)

        browser_tab = ttk.Frame(self.section_notebook)
        browser_tab.columnconfigure(0, weight=1)
        browser_tab.rowconfigure(1, weight=1)
        self.section_notebook.add(browser_tab, text="Experiment Browser")
        ttk.Button(browser_tab, text="Refresh Browser", command=self._refresh_experiment_browser).grid(row=0, column=0, sticky="w", padx=10, pady=(8, 4))
        ttk.Button(browser_tab, text="Export Notebook Bundle", command=self._export_selected_review_packet).grid(row=0, column=0, sticky="e", padx=10, pady=(8, 4))
        filter_row = ttk.Frame(browser_tab)
        filter_row.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 4))
        ttk.Label(filter_row, text="Min Sharpe").pack(side="left")
        self.browser_min_sharpe_var = tk.StringVar(value="")
        ttk.Entry(filter_row, textvariable=self.browser_min_sharpe_var, width=8).pack(side="left", padx=(6, 10))
        ttk.Label(filter_row, text="Approval").pack(side="left")
        self.browser_approval_filter_var = tk.StringVar(value="all")
        ttk.Combobox(filter_row, textvariable=self.browser_approval_filter_var, state="readonly", values=["all", *GOVERNANCE_APPROVAL_STATES], width=12).pack(side="left", padx=(6, 10))
        ttk.Label(filter_row, text="Tag contains").pack(side="left")
        self.browser_tag_filter_var = tk.StringVar(value="")
        ttk.Entry(filter_row, textvariable=self.browser_tag_filter_var, width=16).pack(side="left", padx=(6, 10))
        ttk.Label(filter_row, text="Rank by").pack(side="left")
        self.browser_rank_metric_var = tk.StringVar(value="sharpe")
        ttk.Combobox(filter_row, textvariable=self.browser_rank_metric_var, state="readonly", values=["sharpe", "cagr", "sortino"], width=10).pack(side="left", padx=(6, 10))
        ttk.Button(filter_row, text="Apply", command=self._refresh_experiment_browser).pack(side="left")
        self.experiment_tree = ttk.Treeview(browser_tab, show="headings", height=10)
        self.experiment_tree.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 4))
        self.experiment_tree.bind("<<TreeviewSelect>>", lambda _e: self._on_experiment_tree_selected())

        workflow_row = ttk.Frame(browser_tab)
        workflow_row.grid(row=3, column=0, sticky="ew", padx=10, pady=(2, 4))
        ttk.Label(workflow_row, text="Workflow reason").pack(side="left")
        self.workflow_reason_var = tk.StringVar(value="")
        ttk.Entry(workflow_row, textvariable=self.workflow_reason_var, width=64).pack(side="left", padx=(6, 10), fill="x", expand=True)
        ttk.Button(workflow_row, text="Promote", command=lambda: self._apply_experiment_workflow("promote")).pack(side="left")
        ttk.Button(workflow_row, text="Reject", command=lambda: self._apply_experiment_workflow("reject")).pack(side="left", padx=(6, 0))
        ttk.Button(workflow_row, text="Waive", command=lambda: self._apply_experiment_workflow("waive")).pack(side="left", padx=(6, 0))

        self.experiment_detail_var = tk.StringVar(value="Select an experiment run to inspect tags, metrics, and reproducibility.")
        ttk.Label(browser_tab, textvariable=self.experiment_detail_var, justify="left").grid(row=4, column=0, sticky="ew", padx=10, pady=(2, 8))

        compare_tab = ttk.Frame(self.section_notebook)
        compare_tab.columnconfigure(0, weight=1)
        compare_tab.rowconfigure(2, weight=1)
        self.section_notebook.add(compare_tab, text="Run Comparison")
        selectors = ttk.Frame(compare_tab)
        selectors.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 4))
        ttk.Label(selectors, text="Base Run").pack(side="left")
        self.compare_base_var = tk.StringVar(value="")
        self.compare_base_combo = ttk.Combobox(selectors, textvariable=self.compare_base_var, state="readonly", values=[])
        self.compare_base_combo.pack(side="left", padx=(6, 12))
        ttk.Label(selectors, text="Compare Run").pack(side="left")
        self.compare_other_var = tk.StringVar(value="")
        self.compare_other_combo = ttk.Combobox(selectors, textvariable=self.compare_other_var, state="readonly", values=[])
        self.compare_other_combo.pack(side="left", padx=(6, 12))
        ttk.Button(selectors, text="Diff", command=self._render_selected_run_comparison).pack(side="left")

        cmp_pane = ttk.Panedwindow(compare_tab, orient="horizontal")
        cmp_pane.grid(row=1, column=0, rowspan=2, sticky="nsew", padx=10, pady=4)
        metrics_frame = ttk.Labelframe(cmp_pane, text="Metric Deltas")
        params_frame = ttk.Labelframe(cmp_pane, text="Parameter Diffs")
        variance_frame = ttk.Labelframe(cmp_pane, text="Fold-by-fold WF Variance")
        scenario_frame = ttk.Labelframe(cmp_pane, text="Scenario Comparison")
        frontier_frame = ttk.Labelframe(cmp_pane, text="Robustness Frontier")
        manifest_frame = ttk.Labelframe(cmp_pane, text="Trained-regime A/B Manifest + Provenance")
        cmp_pane.add(metrics_frame, weight=1)
        cmp_pane.add(params_frame, weight=1)
        cmp_pane.add(variance_frame, weight=1)
        cmp_pane.add(scenario_frame, weight=1)
        cmp_pane.add(frontier_frame, weight=1)
        cmp_pane.add(manifest_frame, weight=1)
        metrics_frame.columnconfigure(0, weight=1)
        metrics_frame.rowconfigure(0, weight=1)
        params_frame.columnconfigure(0, weight=1)
        params_frame.rowconfigure(0, weight=1)
        variance_frame.columnconfigure(0, weight=1)
        variance_frame.rowconfigure(0, weight=1)
        self.metric_delta_tree = ttk.Treeview(metrics_frame, show="headings", height=12)
        self.metric_delta_tree.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        self.param_diff_tree = ttk.Treeview(params_frame, show="headings", height=12)
        self.param_diff_tree.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        self.fold_variance_tree = ttk.Treeview(variance_frame, show="headings", height=12)
        self.fold_variance_tree.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        scenario_frame.columnconfigure(0, weight=1)
        scenario_frame.rowconfigure(0, weight=1)
        self.scenario_compare_tree = ttk.Treeview(scenario_frame, show="headings", height=12)
        self.scenario_compare_tree.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        frontier_frame.columnconfigure(0, weight=1)
        frontier_frame.rowconfigure(0, weight=1)
        self.frontier_compare_tree = ttk.Treeview(frontier_frame, show="headings", height=12)
        self.frontier_compare_tree.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        manifest_frame.columnconfigure(0, weight=1)
        manifest_frame.rowconfigure(1, weight=1)
        self.compare_manifest_summary_var = tk.StringVar(value="Manifest and provenance deltas appear here for trained-regime comparisons.")
        ttk.Label(manifest_frame, textvariable=self.compare_manifest_summary_var, justify="left", wraplength=360).grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 2))
        self.manifest_diff_tree = ttk.Treeview(manifest_frame, show="headings", height=10)
        self.manifest_diff_tree.grid(row=1, column=0, sticky="nsew", padx=6, pady=6)

        heatmap_tab = ttk.Frame(self.section_notebook)
        heatmap_tab.columnconfigure(0, weight=1)
        heatmap_tab.rowconfigure(1, weight=1)
        self.section_notebook.add(heatmap_tab, text="Parameter Stability")
        ttk.Label(heatmap_tab, text="Sharpe vs parameter knobs (interactive scatter heatmap) and fold consistency.").grid(row=0, column=0, sticky="w", padx=10, pady=(8, 2))
        self.heatmap_canvas = tk.Canvas(heatmap_tab, height=280, bg="#ffffff", highlightthickness=1, highlightbackground="#d0d0d0")
        self.heatmap_canvas.grid(row=1, column=0, sticky="nsew", padx=10, pady=6)
        self.stability_canvas = tk.Canvas(heatmap_tab, height=180, bg="#ffffff", highlightthickness=1, highlightbackground="#d0d0d0")
        self.stability_canvas.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 8))

        equity_tab = ttk.Frame(self.section_notebook)
        equity_tab.columnconfigure(0, weight=1)
        equity_tab.rowconfigure(1, weight=1)
        self.section_notebook.add(equity_tab, text="Equity/Drawdown charts")

        self.chart_status_var = tk.StringVar(value="Select one or more runs to visualize equity and drawdown.")
        ttk.Label(equity_tab, textvariable=self.chart_status_var, justify="left").grid(row=0, column=0, sticky="ew", padx=10, pady=(6, 0))

        self.equity_canvas = tk.Canvas(equity_tab, height=240, bg="#ffffff", highlightthickness=1, highlightbackground="#d0d0d0")
        self.equity_canvas.grid(row=1, column=0, sticky="nsew", padx=10, pady=6)
        self.equity_canvas.bind("<Configure>", lambda _event: self._update_equity_overlap(self._last_equity_run_dirs))

        self.drawdown_tree = ttk.Treeview(equity_tab, show="headings", height=12)
        self.drawdown_tree.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 8))
        equity_tab.rowconfigure(2, weight=1)
        self._last_equity_run_dirs: list[Path] = []

        run_results_tab = ttk.Frame(self.section_notebook)
        run_results_tab.columnconfigure(0, weight=1)
        run_results_tab.rowconfigure(2, weight=1)
        self.section_notebook.add(run_results_tab, text="Run Results")
        load_row = ttk.Frame(run_results_tab)
        load_row.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 4))
        ttk.Label(load_row, text="Completed run directory").pack(side="left")
        self.results_run_var = tk.StringVar(value="")
        self.results_run_combo = ttk.Combobox(load_row, textvariable=self.results_run_var, state="readonly", values=[])
        self.results_run_combo.pack(side="left", padx=(6, 8), fill="x", expand=True)
        ttk.Button(load_row, text="Load", command=self._load_results_run).pack(side="left", padx=(0, 6))
        ttk.Button(load_row, text="Export cards JSON", command=lambda: self._export_results_payload("results_cards", self._last_results_cards_payload)).pack(side="left")

        cards_row = ttk.Frame(run_results_tab)
        cards_row.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 4))
        self.results_summary_var = tk.StringVar(value="Load a completed run to render timelines, alpha/IR, and diagnostics cards.")
        ttk.Label(cards_row, textvariable=self.results_summary_var, justify="left").pack(side="left", fill="x", expand=True)

        results_notebook = ttk.Notebook(run_results_tab)
        results_notebook.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 8))

        timelines_tab = ttk.Frame(results_notebook)
        timelines_tab.columnconfigure(0, weight=1)
        results_notebook.add(timelines_tab, text="Timelines")
        self.results_equity_canvas = tk.Canvas(timelines_tab, height=150, bg="#fff", highlightthickness=1, highlightbackground="#d0d0d0")
        self.results_equity_canvas.grid(row=0, column=0, sticky="ew", pady=(6, 4))
        self.results_regime_canvas = tk.Canvas(timelines_tab, height=120, bg="#fff", highlightthickness=1, highlightbackground="#d0d0d0")
        self.results_regime_canvas.grid(row=1, column=0, sticky="ew", pady=4)
        self.results_turnover_canvas = tk.Canvas(timelines_tab, height=120, bg="#fff", highlightthickness=1, highlightbackground="#d0d0d0")
        self.results_turnover_canvas.grid(row=2, column=0, sticky="ew", pady=(4, 6))

        alpha_tab = ttk.Frame(results_notebook)
        alpha_tab.columnconfigure(0, weight=1)
        alpha_tab.rowconfigure(0, weight=1)
        results_notebook.add(alpha_tab, text="Benchmark-relative alpha/IR")
        self.alpha_ir_tree = ttk.Treeview(alpha_tab, show="headings", height=10)
        self.alpha_ir_tree.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        ttk.Button(alpha_tab, text="Export alpha/IR CSV", command=lambda: self._export_tree_rows("alpha_ir", self.alpha_ir_tree)).grid(row=1, column=0, sticky="e", padx=6, pady=(0, 6))

        drilldown_tab = ttk.Frame(results_notebook)
        drilldown_tab.columnconfigure(0, weight=1)
        drilldown_tab.rowconfigure(1, weight=1)
        results_notebook.add(drilldown_tab, text="Drill-down")
        drill_btn_row = ttk.Frame(drilldown_tab)
        drill_btn_row.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 0))
        ttk.Button(drill_btn_row, text="Export trade explainability CSV", command=lambda: self._export_tree_rows("trade_explainability", self.trade_explain_tree)).pack(side="left")
        ttk.Button(drill_btn_row, text="Export cost breakdown CSV", command=lambda: self._export_tree_rows("cost_breakdown", self.cost_breakdown_tree)).pack(side="left", padx=(6, 0))
        ttk.Button(drill_btn_row, text="Export failure diagnostics CSV", command=lambda: self._export_tree_rows("failure_windows", self.failure_window_tree)).pack(side="left", padx=(6, 0))

        drill_pane = ttk.Panedwindow(drilldown_tab, orient="horizontal")
        drill_pane.grid(row=1, column=0, sticky="nsew", padx=6, pady=6)
        trade_frame = ttk.Labelframe(drill_pane, text="Trade-level explainability")
        cost_frame = ttk.Labelframe(drill_pane, text="Cost breakdown")
        fail_frame = ttk.Labelframe(drill_pane, text="Failure-window diagnostics")
        drill_pane.add(trade_frame, weight=1)
        drill_pane.add(cost_frame, weight=1)
        drill_pane.add(fail_frame, weight=1)
        for frame in (trade_frame, cost_frame, fail_frame):
            frame.columnconfigure(0, weight=1)
            frame.rowconfigure(0, weight=1)
        self.trade_explain_tree = ttk.Treeview(trade_frame, show="headings", height=10)
        self.trade_explain_tree.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        self.trade_explain_tree.bind("<<TreeviewSelect>>", self._on_trade_drilldown_selected)
        self.cost_breakdown_tree = ttk.Treeview(cost_frame, show="headings", height=10)
        self.cost_breakdown_tree.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        self.failure_window_tree = ttk.Treeview(fail_frame, show="headings", height=10)
        self.failure_window_tree.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        self._trade_drilldown_rows: list[dict[str, object]] = []
        self._cost_drilldown_rows: list[dict[str, object]] = []
        self._failure_drilldown_rows: list[dict[str, object]] = []
        self._last_results_cards_payload: dict[str, object] = {}

        trades_tab = ttk.Frame(self.section_notebook)
        trades_tab.columnconfigure(0, weight=1)
        trades_tab.rowconfigure(1, weight=1)
        self.section_notebook.add(trades_tab, text="Trades/Costs diagnostics")
        self.trades_tree = ttk.Treeview(trades_tab, show="headings", height=14)
        self.trades_tree.grid(row=0, column=0, sticky="nsew", padx=10, pady=(8, 4))

        self.attribution_notebook = ttk.Notebook(trades_tab)
        self.attribution_notebook.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 8))

        turnover_tab = ttk.Frame(self.attribution_notebook)
        turnover_tab.columnconfigure(0, weight=1)
        self.attribution_notebook.add(turnover_tab, text="Turnover Attribution")
        self.turnover_canvas = tk.Canvas(turnover_tab, height=180, bg="#fff", highlightthickness=1, highlightbackground="#d0d0d0")
        self.turnover_canvas.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)

        slippage_tab = ttk.Frame(self.attribution_notebook)
        slippage_tab.columnconfigure(0, weight=1)
        self.attribution_notebook.add(slippage_tab, text="Slippage Attribution")
        self.cost_canvas = tk.Canvas(slippage_tab, height=180, bg="#fff", highlightthickness=1, highlightbackground="#d0d0d0")
        self.cost_canvas.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        self.slippage_summary_var = tk.StringVar(value="Slippage decomposition card will populate after selecting a run.")
        ttk.Label(slippage_tab, textvariable=self.slippage_summary_var, justify="left", foreground="#7a4f00", wraplength=760).grid(row=1, column=0, sticky="ew", padx=6, pady=(0, 6))

        regime_tab = ttk.Frame(self.attribution_notebook)
        regime_tab.columnconfigure(0, weight=1)
        regime_tab.rowconfigure(0, weight=1)
        self.attribution_notebook.add(regime_tab, text="Regime Attribution")
        self.regime_tree = ttk.Treeview(regime_tab, show="headings", height=10)
        self.regime_tree.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)

        drawdown_tab = ttk.Frame(self.attribution_notebook)
        drawdown_tab.columnconfigure(0, weight=1)
        drawdown_tab.rowconfigure(0, weight=1)
        self.attribution_notebook.add(drawdown_tab, text="Drawdown Decomposition")
        self.drawdown_decomp_tree = ttk.Treeview(drawdown_tab, show="headings", height=10)
        self.drawdown_decomp_tree.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)

        capacity_tab = ttk.Frame(self.attribution_notebook)
        capacity_tab.columnconfigure(0, weight=1)
        self.attribution_notebook.add(capacity_tab, text="Capacity Frontier")
        self.capacity_canvas = tk.Canvas(capacity_tab, height=180, bg="#fff", highlightthickness=1, highlightbackground="#d0d0d0")
        self.capacity_canvas.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)

        risk_tab = ttk.Frame(self.attribution_notebook)
        risk_tab.columnconfigure(0, weight=1)
        risk_tab.rowconfigure(1, weight=1)
        self.attribution_notebook.add(risk_tab, text="Risk Control Dashboard")
        self.risk_summary_var = tk.StringVar(value="Risk dashboard and intervention audit trail load per selected run.")
        ttk.Label(risk_tab, textvariable=self.risk_summary_var, justify="left").grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 2))
        self.risk_dashboard_tree = ttk.Treeview(risk_tab, show="headings", height=8)
        self.risk_dashboard_tree.grid(row=1, column=0, sticky="nsew", padx=6, pady=4)
        self.risk_interventions_tree = ttk.Treeview(risk_tab, show="headings", height=8)
        self.risk_interventions_tree.grid(row=2, column=0, sticky="nsew", padx=6, pady=(0, 6))

        wf_tab = ttk.Frame(self.section_notebook)
        wf_tab.columnconfigure(0, weight=1)
        wf_tab.rowconfigure(1, weight=1)
        self.section_notebook.add(wf_tab, text="Fold/WF diagnostics")
        self.wf_status_var = tk.StringVar(value="No walk-forward diagnostics loaded.")
        ttk.Label(wf_tab, textvariable=self.wf_status_var, justify="left").grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 4))
        self.wf_notebook = ttk.Notebook(wf_tab)
        self.wf_notebook.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 8))
        wf_summary = ttk.Frame(self.wf_notebook)
        wf_summary.columnconfigure(0, weight=1)
        wf_summary.rowconfigure(0, weight=1)
        self.wf_notebook.add(wf_summary, text="Fold Summary")
        self.wf_tree = ttk.Treeview(wf_summary, show="headings", height=12)
        self.wf_tree.grid(row=0, column=0, sticky="nsew")
        self.wf_tree.bind("<<TreeviewSelect>>", lambda _e: self._on_fold_selected())

        wf_params = ttk.Frame(self.wf_notebook)
        wf_params.columnconfigure(0, weight=1)
        wf_params.rowconfigure(0, weight=1)
        self.wf_notebook.add(wf_params, text="selected_params")
        self.wf_params_tree = ttk.Treeview(wf_params, show="headings", height=12)
        self.wf_params_tree.grid(row=0, column=0, sticky="nsew")

        wf_equity = ttk.Frame(self.wf_notebook)
        wf_equity.columnconfigure(0, weight=1)
        self.wf_notebook.add(wf_equity, text="OOS equity")
        self.wf_equity_canvas = tk.Canvas(wf_equity, height=220, bg="#fff", highlightthickness=1, highlightbackground="#d0d0d0")
        self.wf_equity_canvas.grid(row=0, column=0, sticky="nsew")

        wf_diag = ttk.Frame(self.wf_notebook)
        wf_diag.columnconfigure(0, weight=1)
        wf_diag.rowconfigure(0, weight=1)
        self.wf_notebook.add(wf_diag, text="Diagnostics")
        self.wf_diag_tree = ttk.Treeview(wf_diag, show="headings", height=12)
        self.wf_diag_tree.grid(row=0, column=0, sticky="nsew")
        self._last_fold_rows: list[dict[str, object]] = []

        logs_tab = ttk.Frame(self.section_notebook)
        logs_tab.columnconfigure(0, weight=1)
        logs_tab.rowconfigure(0, weight=1)
        self.section_notebook.add(logs_tab, text="Logs (Debug)")
        self.logs_text = tk.Text(logs_tab, height=14)
        self.logs_text.grid(row=0, column=0, sticky="nsew", padx=10, pady=8)

        self._load_historical_runs()

    def _set_tree_data(self, tree: ttk.Treeview, rows: list[dict[str, object]]) -> None:
        for item in tree.get_children():
            tree.delete(item)
        if not rows:
            tree["columns"] = ()
            return
        columns = list(rows[0].keys())
        tree["columns"] = columns
        for col in columns:
            tree.heading(col, text=col, command=lambda c=col: self._sort_treeview(tree, c, False))
            tree.column(col, width=130, anchor="w")
        for row in rows:
            values = [self._format_cell(row.get(col)) for col in columns]
            tree.insert("", "end", values=values)

    def _sort_treeview(self, tree: ttk.Treeview, col: str, reverse: bool) -> None:
        data = [(tree.set(item, col), item) for item in tree.get_children("")]
        def sort_key(pair: tuple[str, str]) -> object:
            raw = pair[0]
            try:
                return float(raw)
            except (TypeError, ValueError):
                return str(raw)
        data.sort(key=sort_key, reverse=reverse)
        for index, (_value, item) in enumerate(data):
            tree.move(item, "", index)
        tree.heading(col, command=lambda: self._sort_treeview(tree, col, not reverse))

    def _format_cell(self, value: object) -> str:
        if isinstance(value, float):
            return f"{value:.6f}"
        if value is None:
            return ""
        return str(value)

    def _refresh_artifacts_view(self) -> None:
        if not self.current_run_dirs:
            self.run_listbox.delete(0, tk.END)
            return
        selected = set(self.run_listbox.curselection())
        self.run_listbox.delete(0, tk.END)
        for idx, run_dir in enumerate(self.current_run_dirs):
            self.run_listbox.insert(tk.END, self._run_dir_to_label.get(run_dir, run_dir.name))
            if idx in selected:
                self.run_listbox.selection_set(idx)

    def _on_run_selection_event(self, _event: tk.Event) -> None:
        if self._selection_job is not None:
            self.after_cancel(self._selection_job)
        self._selection_job = self.after(40, self._on_compare_selection_changed)

    def _on_compare_selection_changed(self) -> None:
        self._selection_job = None
        selected = [self.current_run_dirs[idx] for idx in self.run_listbox.curselection() if idx < len(self.current_run_dirs)]
        if len(selected) == 1:
            self._render_single_run(selected[0])
            self.delta_summary_var.set("Select at least two runs for delta metrics.")
            self._set_tree_data(self.inplace_compare_tree, [])
            self._render_guardrails(selected[0])
            return
        if len(selected) < 2:
            self.delta_summary_var.set("Select at least two runs for delta metrics.")
            self._set_tree_data(self.inplace_compare_tree, [])
            return
        metric_maps = [self._load_metric_map(run_dir) for run_dir in selected]
        common_keys = sorted(set.intersection(*(set(metrics.keys()) for metrics in metric_maps if metrics))) if metric_maps else []
        if not common_keys:
            self.delta_summary_var.set("No overlapping metrics found across selected runs.")
            return
        base = metric_maps[0]
        lines = [f"Base run: {selected[0].name}"]
        for run_dir, metrics in zip(selected[1:], metric_maps[1:]):
            deltas = []
            for key in common_keys[:8]:
                delta = float(metrics.get(key, 0.0)) - float(base.get(key, 0.0))
                deltas.append(f"{key}: {delta:+.6f}")
            lines.append(f"vs {run_dir.name} -> " + ", ".join(deltas))
        self.delta_summary_var.set("\n".join(lines))
        self._render_inplace_pair_comparison(selected[0], selected[1])
        self._update_equity_overlap(selected)
        self._populate_run_compare_combos(selected)
        self._render_selected_run_comparison()
        self._render_guardrails(selected[0])

    def _extract_stress_penalty(self, run_dir: Path, metrics: dict[str, float] | None = None) -> float:
        metric_map = metrics if isinstance(metrics, dict) else self._load_metric_map(run_dir)
        if "stress_fragility_index" in metric_map:
            return float(metric_map["stress_fragility_index"])
        stress_payload = read_stress_scenarios(run_dir)
        rows = stress_payload.get("scenario_attribution", []) if isinstance(stress_payload, dict) else []
        if not isinstance(rows, list):
            return 0.0
        penalties: list[float] = []
        for row in rows:
            if not isinstance(row, dict):
                continue
            sharpe_drag = max(0.0, -float(row.get("delta_sharpe", 0.0)))
            drawdown_drag = max(0.0, -float(row.get("delta_max_drawdown", 0.0)))
            penalties.append((0.6 * sharpe_drag) + (0.4 * drawdown_drag))
        return float(sum(penalties) / len(penalties)) if penalties else 0.0

    def _extract_governance_readiness(self, run_dir: Path) -> float:
        manifest = self._read_json(run_dir / "manifest.json")
        if not isinstance(manifest, dict):
            return 0.0
        governance = manifest.get("governance", {}) if isinstance(manifest.get("governance"), dict) else {}
        gate_checks = governance.get("gate_checks", {}) if isinstance(governance.get("gate_checks"), dict) else {}
        if gate_checks:
            total = len(gate_checks)
            passed = sum(1 for value in gate_checks.values() if bool(value))
            return float(passed / total) if total else 0.0
        if "is_promotion_ready" in governance:
            return 1.0 if bool(governance.get("is_promotion_ready", False)) else 0.0
        return 0.0

    def _render_inplace_pair_comparison(self, base_run: Path, other_run: Path) -> None:
        base_metrics = self._load_metric_map(base_run)
        other_metrics = self._load_metric_map(other_run)
        metric_rows = metric_deltas(base_metrics, other_metrics)
        metric_lookup = {str(row.get("metric", "")): row for row in metric_rows if isinstance(row, dict)}

        drawdown_row = metric_lookup.get("max_drawdown") or metric_lookup.get("rolling_drawdown_worst")
        turnover_row = metric_lookup.get("turnover_total") or metric_lookup.get("turnover")
        sharpe_row = metric_lookup.get("sharpe")

        base_stress_penalty = self._extract_stress_penalty(base_run, base_metrics)
        other_stress_penalty = self._extract_stress_penalty(other_run, other_metrics)
        base_readiness = self._extract_governance_readiness(base_run)
        other_readiness = self._extract_governance_readiness(other_run)

        rows = [
            {
                "metric": "Sharpe",
                "base": float(sharpe_row.get("base", 0.0)) if isinstance(sharpe_row, dict) else 0.0,
                "compare": float(sharpe_row.get("compare", 0.0)) if isinstance(sharpe_row, dict) else 0.0,
                "delta": float(sharpe_row.get("delta", 0.0)) if isinstance(sharpe_row, dict) else 0.0,
            },
            {
                "metric": "Drawdown",
                "base": float(drawdown_row.get("base", 0.0)) if isinstance(drawdown_row, dict) else 0.0,
                "compare": float(drawdown_row.get("compare", 0.0)) if isinstance(drawdown_row, dict) else 0.0,
                "delta": float(drawdown_row.get("delta", 0.0)) if isinstance(drawdown_row, dict) else 0.0,
            },
            {
                "metric": "Turnover",
                "base": float(turnover_row.get("base", 0.0)) if isinstance(turnover_row, dict) else 0.0,
                "compare": float(turnover_row.get("compare", 0.0)) if isinstance(turnover_row, dict) else 0.0,
                "delta": float(turnover_row.get("delta", 0.0)) if isinstance(turnover_row, dict) else 0.0,
            },
            {
                "metric": "Stress penalty",
                "base": base_stress_penalty,
                "compare": other_stress_penalty,
                "delta": other_stress_penalty - base_stress_penalty,
            },
            {
                "metric": "Governance readiness",
                "base": base_readiness,
                "compare": other_readiness,
                "delta": other_readiness - base_readiness,
            },
        ]
        self._set_tree_data(self.inplace_compare_tree, rows)

    def _update_equity_overlap(self, run_dirs: list[Path]) -> None:
        self._last_equity_run_dirs = list(run_dirs)
        self.equity_canvas.delete("all")
        width = max(10, int(self.equity_canvas.winfo_width()))
        height = max(10, int(self.equity_canvas.winfo_height()))
        if width <= 20 or height <= 20:
            return

        margin_left = 50
        margin_right = 12
        margin_top = 12
        margin_bottom = 20
        plot_w = max(10, width - margin_left - margin_right)
        plot_h = max(10, height - margin_top - margin_bottom)

        palette = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd", "#8c564b"]
        curves: list[tuple[str, list[float], str]] = []
        for idx, run_dir in enumerate(run_dirs):
            equity_rows = self._load_rows(run_dir, "equity")
            if not equity_rows:
                continue
            values: list[float] = []
            for row in equity_rows:
                value = self._safe_float(row.get("equity"))
                if value is not None:
                    values.append(value)
            if values:
                curves.append((run_dir.name, values, palette[idx % len(palette)]))

        if not curves:
            self.chart_status_var.set("No equity series available for selected runs.")
            self.equity_canvas.create_text(width // 2, height // 2, text="No equity data", fill="#777")
            return

        all_values = [v for _name, series, _color in curves for v in series]
        min_v = min(all_values)
        max_v = max(all_values)
        if max_v <= min_v:
            max_v = min_v + 1.0

        self.equity_canvas.create_rectangle(
            margin_left,
            margin_top,
            margin_left + plot_w,
            margin_top + plot_h,
            outline="#d0d0d0",
        )

        def map_x(i: int, n: int) -> float:
            if n <= 1:
                return margin_left
            return margin_left + (i / (n - 1)) * plot_w

        def map_y(v: float) -> float:
            return margin_top + (1.0 - (v - min_v) / (max_v - min_v)) * plot_h

        legend_y = margin_top + 8
        for name, series, color in curves[:6]:
            points: list[float] = []
            for i, value in enumerate(series):
                points.extend([map_x(i, len(series)), map_y(value)])
            if len(points) >= 4:
                self.equity_canvas.create_line(*points, fill=color, width=1.8)
            self.equity_canvas.create_rectangle(margin_left + plot_w - 170, legend_y - 6, margin_left + plot_w - 158, legend_y + 6, fill=color, outline=color)
            self.equity_canvas.create_text(margin_left + plot_w - 154, legend_y, anchor="w", text=name[:30], fill="#222")
            legend_y += 16

        self.equity_canvas.create_text(8, margin_top + 2, anchor="nw", text=f"{max_v:.2f}", fill="#666")
        self.equity_canvas.create_text(8, margin_top + plot_h - 10, anchor="nw", text=f"{min_v:.2f}", fill="#666")
        self.chart_status_var.set(f"Showing overlap for {len(curves)} run(s).")

    def _safe_float(self, value: object) -> float | None:
        try:
            if value is None or value == "":
                return None
            return float(value)
        except (TypeError, ValueError):
            return None

    def _load_metric_map(self, run_dir: Path) -> dict[str, float]:
        metrics_path = run_dir / "metrics.json"
        if not metrics_path.exists():
            metrics_path = run_dir / "aggregate_metrics.json"
            if metrics_path.exists():
                parsed = self._read_json(metrics_path)
                if isinstance(parsed, dict):
                    return {str(k): float(v) for k, v in parsed.items() if isinstance(v, (int, float))}
            return {}
        rows = self._read_json(metrics_path)
        if not isinstance(rows, list):
            return {}
        output: dict[str, float] = {}
        for row in rows:
            if not isinstance(row, dict):
                continue
            metric = str(row.get("metric", ""))
            value = row.get("value")
            if metric and isinstance(value, (int, float)):
                output[metric] = float(value)
        return output

    def _load_rows(self, run_dir: Path, stem: str) -> list[dict[str, object]]:
        json_path = run_dir / f"{stem}.json"
        if json_path.exists():
            parsed = self._read_json(json_path)
            if isinstance(parsed, list):
                return [row for row in parsed if isinstance(row, dict)]
        csv_path = run_dir / f"{stem}.csv"
        if csv_path.exists():
            with csv_path.open("r", newline="") as handle:
                reader = csv.DictReader(handle)
                return [dict(row) for row in reader]
        return []

    def _read_json(self, path: Path) -> object:
        try:
            return json.loads(path.read_text())
        except Exception:
            return None

    def _consume_run_outputs(self, output_text: str) -> None:
        self._last_output_text = output_text
        self.logs_text.delete("1.0", tk.END)
        self.logs_text.insert("1.0", output_text)

        run_dirs: list[Path] = []
        for marker in ("Leaderboard outputs:", "Saved outputs to:"):
            idx = output_text.rfind(marker)
            if idx < 0:
                continue
            raw = output_text[idx + len(marker):].strip().splitlines()[0].strip()
            candidate = Path(raw)
            if candidate.exists():
                run_dirs.append(candidate)
        if not run_dirs:
            return

        primary_dir = run_dirs[-1]
        self._register_run_dirs(run_dirs)
        leaderboard_rows = self._load_rows(primary_dir, "leaderboard")
        if leaderboard_rows:
            combo_dirs: list[Path] = []
            for row in leaderboard_rows:
                run_dir_str = str(row.get("run_dir", "")).strip()
                if run_dir_str:
                    candidate = Path(run_dir_str)
                    if candidate.exists():
                        combo_dirs.append(candidate)
            self._register_run_dirs(combo_dirs)
            self._set_tree_data(self.leaderboard_tree, leaderboard_rows)
        else:
            metrics = self._load_metric_map(primary_dir)
            self._set_tree_data(
                self.leaderboard_tree,
                [{"metric": key, "value": value} for key, value in sorted(metrics.items())],
            )

        self._refresh_artifacts_view()
        self._render_single_run(primary_dir)

    def _render_single_run(self, run_dir: Path) -> None:
        drawdown_rows = self._load_rows(run_dir, "drawdown")
        if not drawdown_rows:
            drawdown_rows = self._load_rows(run_dir, "risk_diagnostics")
        self._set_tree_data(self.drawdown_tree, drawdown_rows[:200])

        trade_rows = self._load_rows(run_dir, "trades")
        if not trade_rows:
            trade_rows = self._load_rows(run_dir, "trade_log")
        self._set_tree_data(self.trades_tree, trade_rows[:200])

        fold_summary = self._read_json(run_dir / "fold_summary.json")
        if isinstance(fold_summary, list):
            wf_rows = [row for row in fold_summary if isinstance(row, dict)]
            self._last_fold_rows = wf_rows
            self._set_tree_data(self.wf_tree, wf_rows)
            self.wf_status_var.set(f"Loaded {len(wf_rows)} walk-forward folds from {run_dir.name}.")
        else:
            self._last_fold_rows = []
            self._set_tree_data(self.wf_tree, [])
            self.wf_status_var.set("No walk-forward diagnostics loaded.")

        self._update_equity_overlap([run_dir])
        self._render_parameter_stability(run_dir)
        self._render_cost_attribution(run_dir)
        self._render_risk_controls(run_dir)
        self._render_guardrails(run_dir)
        self._populate_run_compare_combos()
        self._refresh_experiment_browser()

    def _render_risk_controls(self, run_dir: Path) -> None:
        dashboard_rows = self._load_rows(run_dir, "risk_dashboard")
        interventions = self._load_rows(run_dir, "risk_interventions")
        self._set_tree_data(self.risk_dashboard_tree, dashboard_rows[:200])
        self._set_tree_data(self.risk_interventions_tree, interventions[:200])

        if dashboard_rows:
            latest = dashboard_rows[-1]
            regime = latest.get("regime_state", "unknown")
            confidence = self._safe_float(latest.get("model_confidence", 0.0)) or 0.0
            self.risk_summary_var.set(
                f"Latest regime={regime} model_confidence={confidence:.3f} | interventions={len(interventions)}"
            )
        else:
            self.risk_summary_var.set("No risk dashboard artifacts found for selected run.")

    def _load_historical_runs(self) -> None:
        run_dirs = self._scan_backtest_output_runs()
        if not run_dirs:
            return
        self._register_run_dirs(run_dirs)
        self._refresh_artifacts_view()
        self._render_single_run(self.current_run_dirs[0])

    def _scan_backtest_output_runs(self) -> list[Path]:
        candidates: list[Path] = []
        state_mapping = self._remote_synced_run_dirs()
        if BACKTEST_OUTPUT_DIR.exists():
            for path in BACKTEST_OUTPUT_DIR.iterdir():
                if not path.is_dir():
                    continue
                prefix = f"{DEFAULT_REMOTE_NAMESPACE_PREFIX}__"
                if path.name.startswith(prefix):
                    remote_job_id = path.name[len(prefix):].strip()
                    if remote_job_id and state_mapping.get(remote_job_id) != path:
                        state_mapping[remote_job_id] = path
                if (path / "leaderboard.json").exists() or (path / "metrics.json").exists() or (path / "aggregate_metrics.json").exists() or (path / "manifest.json").exists() or (path / "fold_summary.json").exists():
                    candidates.append(path)

        self._persist_remote_sync_mapping(state_mapping)
        remote_mappings = self._remote_synced_run_dirs()
        for mapped_path in remote_mappings.values():
            if mapped_path.exists() and mapped_path.is_dir():
                candidates.append(mapped_path)

        deduped: list[Path] = []
        seen: set[Path] = set()
        for run_dir in candidates:
            resolved = run_dir.resolve()
            if resolved in seen:
                continue
            seen.add(resolved)
            deduped.append(run_dir)
        deduped.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        return deduped


    def _persist_remote_sync_mapping(self, mapping: dict[str, Path]) -> None:
        normalized = {job_id: str(path) for job_id, path in mapping.items() if job_id and str(path).strip()}
        if normalized == self.controller.state.remote_synced_runs:
            return
        self.controller.state.remote_synced_runs = normalized
        self.controller.persist_state()

    def _remote_synced_run_dirs(self) -> dict[str, Path]:
        raw_mapping = self.controller.state.remote_synced_runs if isinstance(self.controller.state.remote_synced_runs, dict) else {}
        mapping: dict[str, Path] = {}
        for remote_job_id, raw_path in raw_mapping.items():
            if not isinstance(remote_job_id, str) or not isinstance(raw_path, str):
                continue
            cleaned_job_id = remote_job_id.strip()
            cleaned_path = raw_path.strip()
            if not cleaned_job_id or not cleaned_path:
                continue
            mapping[cleaned_job_id] = Path(cleaned_path)
        return mapping

    def _register_run_dirs(self, run_dirs: list[Path]) -> None:
        merged: list[Path] = list(self.current_run_dirs)
        seen = {run.resolve() for run in self.current_run_dirs if run.exists()}
        for run_dir in run_dirs:
            if not run_dir.exists():
                continue
            resolved = run_dir.resolve()
            if resolved in seen:
                continue
            seen.add(resolved)
            merged.append(run_dir)
            self._run_dir_to_label[run_dir] = self._format_run_label(run_dir)
        merged.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        self.current_run_dirs = merged

    def _format_run_label(self, run_dir: Path) -> str:
        parts = [run_dir.name]
        remote_mapping = self._remote_synced_run_dirs()
        remote_job_id = next((job_id for job_id, mapped_path in remote_mapping.items() if mapped_path.exists() and mapped_path.resolve() == run_dir.resolve()), None)
        if remote_job_id:
            parts.insert(0, f"{DEFAULT_REMOTE_NAMESPACE_PREFIX}:{remote_job_id}")

        manifest = self._read_json(run_dir / "manifest.json")
        if isinstance(manifest, dict):
            strategy = manifest.get("strategy")
            timeframe = manifest.get("timeframe")
            if isinstance(strategy, str) and strategy:
                parts.append(strategy)
            if isinstance(timeframe, str) and timeframe:
                parts.append(timeframe)
        return " | ".join(parts)

    def _refresh_experiment_browser(self) -> None:
        rows = read_experiment_index(BACKTEST_OUTPUT_DIR)
        rendered: list[dict[str, object]] = []
        min_sharpe = parse_float(self.browser_min_sharpe_var.get()) if hasattr(self, "browser_min_sharpe_var") else None
        approval_filter = (self.browser_approval_filter_var.get().strip().lower() if hasattr(self, "browser_approval_filter_var") else "all")
        tag_filter = (self.browser_tag_filter_var.get().strip().lower() if hasattr(self, "browser_tag_filter_var") else "")
        rank_metric = (self.browser_rank_metric_var.get().strip() if hasattr(self, "browser_rank_metric_var") else "sharpe") or "sharpe"
        for row in reversed(rows[-500:]):
            run_dir = Path(str(row.get("run_dir", "")))
            manifest = self._read_json(run_dir / "manifest.json")
            metrics = self._load_metric_map(run_dir)
            tags = parse_tags(manifest if isinstance(manifest, dict) else None)
            sharpe = self._safe_float(metrics.get("sharpe", row.get("primary_metric_value", "")))
            approval = str((row.get("governance") or {}).get("approval_status", "")).strip().lower()
            if min_sharpe is not None and sharpe is not None and sharpe < float(min_sharpe):
                continue
            if approval_filter != "all" and approval != approval_filter:
                continue
            if tag_filter and not any(tag_filter in str(tag).lower() for tag in tags):
                continue
            rank_value = self._safe_float(metrics.get(rank_metric, row.get("primary_metric_value", 0.0))) or -1e9
            governance = (row.get("governance") or {}) if isinstance(row.get("governance"), dict) else {}
            rendered.append(
                {
                    "rank_metric": rank_metric,
                    "rank_value": rank_value,
                    "timestamp": row.get("timestamp", ""),
                    "run_type": row.get("run_type", ""),
                    "run": run_dir.name,
                    "tags": ", ".join(tags[:4]),
                    "best_sharpe": sharpe if sharpe is not None else "",
                    "approval": approval,
                    "promotion": str(governance.get("promotion_state", "")),
                    "experiment_id": str(governance.get("experiment_id", ""))[:20],
                    "model_artifacts": len(row.get("model_artifacts", []) or []),
                    "plot_artifacts": len(row.get("plot_artifacts", []) or []),
                    "run_id": str(row.get("run_id", ""))[:12],
                    "cfg_hash": str(row.get("config_hash", ""))[:12],
                    "cfg_chk": str(row.get("config_checksum", ""))[:12],
                    "snapshot_chk": str(row.get("data_snapshot_checksum", ""))[:12],
                    "manifest_chk": str(row.get("manifest_checksum", ""))[:12],
                    "fingerprint": str(row.get("reproducibility_fingerprint", ""))[:12],
                }
            )
        rendered.sort(key=lambda item: float(item.get("rank_value", -1e9)), reverse=True)
        for idx, item in enumerate(rendered, start=1):
            item["rank"] = idx
            item.pop("rank_value", None)
            item.pop("rank_metric", None)
        self._set_tree_data(self.experiment_tree, rendered)

    def _on_experiment_tree_selected(self) -> None:
        selected = self.experiment_tree.selection()
        if not selected:
            return
        values = self.experiment_tree.item(selected[0], "values")
        if len(values) < 3:
            return
        run_name = str(values[2])
        run_dir = next((d for d in self.current_run_dirs if d.name == run_name), None)
        if run_dir is None:
            self.experiment_detail_var.set(f"Run {run_name} not found on disk.")
            return
        manifest = self._read_json(run_dir / "manifest.json")
        metrics = self._load_metric_map(run_dir)
        tags = parse_tags(manifest if isinstance(manifest, dict) else None)
        fp = ""
        if isinstance(manifest, dict):
            fp = str(manifest.get("reproducibility_fingerprint", ""))
        run_id = str(manifest.get("run_id", "")) if isinstance(manifest, dict) else ""
        cfg_hash = str(manifest.get("config_hash", "")) if isinstance(manifest, dict) else ""
        cfg_chk = str(manifest.get("config_checksum", "")) if isinstance(manifest, dict) else ""
        snap_chk = str(manifest.get("data_snapshot_checksum", "")) if isinstance(manifest, dict) else ""
        man_chk = str(manifest.get("manifest_checksum", "")) if isinstance(manifest, dict) else ""
        governance = manifest.get("governance", {}) if isinstance(manifest, dict) and isinstance(manifest.get("governance"), dict) else {}
        repro_meta = manifest.get("reproducibility_metadata", {}) if isinstance(manifest, dict) and isinstance(manifest.get("reproducibility_metadata"), dict) else {}
        feature_hashes = repro_meta.get("feature_hashes", {}) if isinstance(repro_meta.get("feature_hashes"), dict) else {}
        drift_monitoring = governance.get("drift_monitoring", {}) if isinstance(governance.get("drift_monitoring"), dict) else {}
        governance_diagnostics = governance.get("governance_diagnostics", {}) if isinstance(governance.get("governance_diagnostics"), dict) else {}
        gate_checks = governance.get("gate_checks", {}) if isinstance(governance.get("gate_checks"), dict) else {}
        missing_checks = governance.get("missing_required_checks", []) if isinstance(governance.get("missing_required_checks"), list) else []
        drift_alerts = drift_monitoring.get("alert_summaries", []) if isinstance(drift_monitoring.get("alert_summaries"), list) else []
        dsr_diag = governance_diagnostics.get("deflated_sharpe_reality_check", {}) if isinstance(governance_diagnostics.get("deflated_sharpe_reality_check"), dict) else {}
        stability_diag = governance_diagnostics.get("parameter_stability", {}) if isinstance(governance_diagnostics.get("parameter_stability"), dict) else {}
        split_drift_diag = governance_diagnostics.get("train_validation_test_drift", {}) if isinstance(governance_diagnostics.get("train_validation_test_drift"), dict) else {}
        msg = [
            f"Run: {run_dir.name}",
            f"Run ID: {run_id[:24] if run_id else 'n/a'}",
            f"Tags: {', '.join(tags) if tags else '-'}",
            f"Sharpe: {metrics.get('sharpe', 'n/a')}",
            f"CAGR: {metrics.get('cagr', 'n/a')}",
            f"Promotion: {governance.get('promotion_state', 'n/a')}",
            f"Approval: {governance.get('approval_status', 'n/a')}",
            f"Experiment ID: {str(governance.get('experiment_id', '')).strip() or 'n/a'}",
            f"Drift monitor: {'OK' if drift_monitoring.get('within_tolerance', False) else 'BREACH'}",
            f"Drift alerts: {len(drift_alerts)}",
            f"Deflated Sharpe / RC: {'OK' if gate_checks.get('deflated_sharpe_reality_check', False) else 'BREACH'} (dsr={float(dsr_diag.get('deflated_sharpe_ratio', 0.0) or 0.0):.3f}, p={float(dsr_diag.get('combined_reality_check_pvalue', 1.0) or 1.0):.3f})",
            f"Parameter stability penalty: {'OK' if gate_checks.get('parameter_stability_penalty', False) else 'BREACH'} ({float(stability_diag.get('parameter_stability_penalty', 1.0) or 1.0):.3f})",
            f"Train/validation/test drift: {'OK' if gate_checks.get('train_validation_test_drift', False) else 'BREACH'} (tv={float(split_drift_diag.get('train_validation_abs_drift', 0.0) or 0.0):.3f}, vt={float(split_drift_diag.get('validation_test_abs_drift', 0.0) or 0.0):.3f})",
            f"Missing required checks: {', '.join(missing_checks) if missing_checks else 'none'}",
            f"Config hash/checksum: {(cfg_hash[:12] if cfg_hash else 'n/a')} / {(cfg_chk[:12] if cfg_chk else 'n/a')}",
            f"Snapshot checksum: {snap_chk[:24] if snap_chk else 'n/a'}",
            f"Manifest checksum: {man_chk[:24] if man_chk else 'n/a'}",
            f"Fingerprint: {fp[:24] if fp else 'n/a'}",
            f"Feature hashes tracked: {len(feature_hashes)}",
        ]
        self.experiment_detail_var.set("\n".join(msg))

    def _apply_experiment_workflow(self, action: str) -> None:
        selected = self.experiment_tree.selection()
        if not selected:
            messagebox.showinfo("Workflow", "Select an experiment run first.")
            return
        reason = self.workflow_reason_var.get().strip()
        if not reason:
            messagebox.showinfo("Workflow", "Enter a reason for promote/reject/waive.")
            return
        values = self.experiment_tree.item(selected[0], "values")
        if len(values) < 3:
            messagebox.showerror("Workflow", "Invalid experiment selection.")
            return
        run_name = str(values[2])
        run_id = ""
        for row in read_experiment_index(BACKTEST_OUTPUT_DIR):
            row_run = Path(str(row.get("run_dir", ""))).name
            if row_run == run_name:
                run_id = str(row.get("run_id", ""))
                if run_id:
                    break
        if not run_id:
            messagebox.showerror("Workflow", f"Could not resolve run ID for {run_name}.")
            return
        if action == "promote":
            run_dir = next((d for d in self.current_run_dirs if d.name == run_name), None)
            if run_dir is not None:
                manifest = self._read_json(run_dir / "manifest.json")
                if isinstance(manifest, dict):
                    governance = manifest.get("governance", {}) if isinstance(manifest.get("governance"), dict) else {}
                    drift_monitoring = governance.get("drift_monitoring", {}) if isinstance(governance.get("drift_monitoring"), dict) else {}
                    if not bool(drift_monitoring.get("within_tolerance", False)):
                        messagebox.showerror("Workflow", "Promotion blocked: drift monitoring exceeds configured tolerances.")
                        return
                    if not str(governance.get("experiment_id", "")).strip():
                        messagebox.showerror("Workflow", "Promotion blocked: experiment ID is required for deployment approvals.")
                        return
        success = apply_governance_decision(
            BACKTEST_OUTPUT_DIR,
            run_id=run_id,
            action=action,
            reason=reason,
            actor=self.gov_owner_var.get().strip() or "ui",
        )
        if not success:
            messagebox.showerror("Workflow", f"Run ID {run_id} not found in experiment registry or promotion blocked by drift tolerances.")
            return
        self.workflow_reason_var.set("")
        self._refresh_experiment_browser()
        self.experiment_detail_var.set(f"Applied {action} to run {run_id[:12]} with reason logged.")

    def export_prompt_pack(self) -> None:
        if not self.save_settings(show_confirmation=False):
            return
        settings_payload = dict(self.controller.state.backtest_settings or {})
        recent_outputs = self._collect_prompt_pack_recent_outputs()
        output_path = write_prompt_pack(
            output_dir=BACKTEST_OUTPUT_DIR / "prompt_packs",
            file_stem="backtesting_prompt_pack",
            title="Backtesting Prompt Pack",
            config=settings_payload,
            recent_outputs=recent_outputs,
        )
        messagebox.showinfo("Prompt pack exported", f"Saved prompt pack to:\n{output_path}")

    def _collect_prompt_pack_recent_outputs(self) -> dict[str, object]:
        run_dir: Path | None = None
        selected = self.experiment_tree.selection()
        if selected:
            values = self.experiment_tree.item(selected[0], "values")
            if len(values) >= 3:
                run_name = str(values[2])
                run_dir = next((d for d in self.current_run_dirs if d.name == run_name), None)
        if run_dir is None and self.current_run_dirs:
            run_dir = self.current_run_dirs[-1]

        payload: dict[str, object] = {
            "last_console_output": self._last_output_text[-6000:],
            "available_run_dirs": [str(path) for path in self.current_run_dirs[-10:]],
        }
        if run_dir is None:
            return payload

        payload["selected_run_dir"] = str(run_dir)
        payload["selected_run_name"] = run_dir.name
        payload["metrics"] = self._load_metric_map(run_dir)
        manifest = self._read_json(run_dir / "manifest.json")
        payload["manifest"] = manifest if isinstance(manifest, dict) else {}
        for stem in ("fold_summary", "stress_scenarios", "risk_diagnostics", "drawdown", "trades", "trade_log"):
            rows = self._load_rows(run_dir, stem)
            if rows:
                payload[f"{stem}_sample"] = rows[:25]
        return payload

    def _export_selected_review_packet(self) -> None:
        selected = self.experiment_tree.selection()
        if not selected:
            messagebox.showinfo("Notebook bundle", "Select an experiment run first.")
            return
        values = self.experiment_tree.item(selected[0], "values")
        if len(values) < 3:
            return
        run_name = str(values[2])
        run_dir = next((d for d in self.current_run_dirs if d.name == run_name), None)
        if run_dir is None:
            messagebox.showinfo("Notebook bundle", f"Run {run_name} not found on disk.")
            return

        bundle_dir = run_dir / "notebook_bundle"
        bundle_dir.mkdir(parents=True, exist_ok=True)
        copied: list[str] = []
        stems = [
            "metrics",
            "equity",
            "returns",
            "trades",
            "trade_log",
            "drawdown",
            "risk_diagnostics",
            "turnover_by_symbol",
            "regime_pnl_attribution",
            "fold_summary",
            "capacity_frontier",
            "stress_scenarios",
            "robustness_frontier",
        ]
        for stem in stems:
            rows = self._load_rows(run_dir, stem)
            if rows:
                csv_path = bundle_dir / f"{stem}.csv"
                with csv_path.open("w", newline="") as handle:
                    writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
                    writer.writeheader()
                    writer.writerows(rows)
                copied.append(csv_path.name)
                try:
                    import pandas as pd  # type: ignore
                    pd.DataFrame(rows).to_parquet(bundle_dir / f"{stem}.parquet", index=False)
                    copied.append(f"{stem}.parquet")
                except Exception:
                    pass
            else:
                src_json = run_dir / f"{stem}.json"
                if src_json.exists():
                    dst_json = bundle_dir / src_json.name
                    dst_json.write_text(src_json.read_text(encoding="utf-8"), encoding="utf-8")
                    copied.append(dst_json.name)

        manifest = self._read_json(run_dir / "manifest.json")
        metrics = self._load_metric_map(run_dir)
        packet = {
            "run": run_dir.name,
            "source_dir": str(run_dir),
            "manifest": manifest if isinstance(manifest, dict) else {},
            "metrics": metrics,
            "bundle_files": sorted(copied),
            "generated_at": date.today().isoformat(),
        }
        packet_path = bundle_dir / "bundle_metadata.json"
        packet_path.write_text(json.dumps(packet, indent=2), encoding="utf-8")
        messagebox.showinfo("Notebook bundle", f"Saved notebook bundle to {bundle_dir}")

    def _populate_run_compare_combos(self, selected: list[Path] | None = None) -> None:
        runs = selected or self.current_run_dirs
        names = [run.name for run in runs]
        self.compare_base_combo.configure(values=names)
        self.compare_other_combo.configure(values=names)
        if hasattr(self, "results_run_combo"):
            self.results_run_combo.configure(values=names)
            if names and not self.results_run_var.get():
                self.results_run_var.set(names[0])
        if names and not self.compare_base_var.get():
            self.compare_base_var.set(names[0])
        if len(names) > 1 and not self.compare_other_var.get():
            self.compare_other_var.set(names[1])

    def _resolve_run_by_name(self, name: str) -> Path | None:
        clean = name.strip()
        if not clean:
            return None
        for run_dir in self.current_run_dirs:
            if run_dir.name == clean:
                return run_dir
        return None

    def _export_tree_rows(self, stem: str, tree: ttk.Treeview) -> None:
        run_dir = self._resolve_run_by_name(self.results_run_var.get()) if hasattr(self, "results_run_var") else None
        if run_dir is None:
            selected = [self.current_run_dirs[idx] for idx in self.run_listbox.curselection() if idx < len(self.current_run_dirs)] if hasattr(self, "run_listbox") else []
            run_dir = selected[0] if selected else (self.current_run_dirs[0] if self.current_run_dirs else None)
        if run_dir is None:
            messagebox.showinfo("Export", "Select a run before exporting.")
            return
        rows = self._tree_rows(tree)
        if not rows:
            messagebox.showinfo("Export", "No rows to export.")
            return
        export_dir = run_dir / "exports"
        export_dir.mkdir(parents=True, exist_ok=True)
        csv_path = export_dir / f"{stem}.csv"
        json_path = export_dir / f"{stem}.json"
        with csv_path.open("w", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
            writer.writeheader()
            writer.writerows(rows)
        json_path.write_text(json.dumps(rows, indent=2), encoding="utf-8")
        messagebox.showinfo("Export", f"Exported {stem} CSV/JSON to {export_dir}")

    def _export_results_payload(self, stem: str, payload: dict[str, object]) -> None:
        run_dir = self._resolve_run_by_name(self.results_run_var.get()) if hasattr(self, "results_run_var") else None
        if run_dir is None:
            messagebox.showinfo("Export", "Load a run in Run Results first.")
            return
        export_dir = run_dir / "exports"
        export_dir.mkdir(parents=True, exist_ok=True)
        path = export_dir / f"{stem}.json"
        path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        messagebox.showinfo("Export", f"Exported {stem} JSON to {path}")

    def _tree_rows(self, tree: ttk.Treeview) -> list[dict[str, object]]:
        columns = [str(c) for c in tree["columns"]]
        rows: list[dict[str, object]] = []
        for item in tree.get_children():
            values = tree.item(item, "values")
            rows.append({col: values[idx] if idx < len(values) else "" for idx, col in enumerate(columns)})
        return rows

    def _load_results_run(self) -> None:
        run_dir = self._resolve_run_by_name(self.results_run_var.get())
        if run_dir is None:
            return
        self._render_single_run(run_dir)
        metrics = self._load_metric_map(run_dir)
        equity_rows = self._load_rows(run_dir, "equity")
        drawdown_rows = self._load_rows(run_dir, "drawdown")
        regime_rows = self._load_rows(run_dir, "regime_pnl_attribution")
        turnover_rows = self._load_rows(run_dir, "turnover_by_symbol")
        risk_rows = self._load_rows(run_dir, "risk_diagnostics")

        equity_vals = [self._safe_float(r.get("equity")) for r in equity_rows]
        dd_vals = [self._safe_float(r.get("drawdown")) for r in drawdown_rows]
        turnover_vals = [self._safe_float(r.get("turnover")) for r in turnover_rows]
        cost_vals = [self._safe_float(r.get("cost_total")) for r in risk_rows]
        self._draw_line_canvas(self.results_equity_canvas, [v for v in equity_vals if v is not None], color="#1f77b4")
        if any(v is not None for v in dd_vals):
            self._draw_line_canvas(self.results_regime_canvas, [v for v in dd_vals if v is not None], color="#d62728")
        else:
            regime_scores = [self._safe_float(r.get("pnl_total")) for r in regime_rows]
            self._draw_line_canvas(self.results_regime_canvas, [v for v in regime_scores if v is not None], color="#9467bd")
        timeline = [v for v in turnover_vals if v is not None] + [v for v in cost_vals if v is not None]
        self._draw_line_canvas(self.results_turnover_canvas, timeline if len(timeline) > 1 else [0.0, 0.0], color="#ff7f0e")

        alpha_rows = [
            {
                "metric": "alpha_total",
                "value": float(metrics.get("alpha_total", metrics.get("excess_return", 0.0))),
            },
            {
                "metric": "information_ratio",
                "value": float(metrics.get("information_ratio", metrics.get("ir", 0.0))),
            },
            {
                "metric": "benchmark_sharpe",
                "value": float(metrics.get("benchmark_sharpe", 0.0)),
            },
        ]
        self._set_tree_data(self.alpha_ir_tree, alpha_rows)

        trades = self._load_rows(run_dir, "trade_log") or self._load_rows(run_dir, "trades")
        explain_rows: list[dict[str, object]] = []
        for row in trades[:400]:
            explain_rows.append({
                "trade_id": row.get("trade_id", row.get("id", "")),
                "timestamp": row.get("timestamp", row.get("entry_time", "")),
                "symbol": row.get("symbol", ""),
                "side": row.get("side", ""),
                "pnl": row.get("pnl", row.get("net_pnl", "")),
                "regime": row.get("regime", row.get("regime_label", "")),
                "market_state": row.get("market_state", row.get("state", "")),
                "stress_scenario": row.get("stress_scenario", row.get("scenario", "")),
            })
        self._trade_drilldown_rows = explain_rows
        self._set_tree_data(self.trade_explain_tree, explain_rows)

        cost_breakdown_rows = self._load_rows(run_dir, "cost_breakdown")
        if not cost_breakdown_rows:
            cost_breakdown_rows = [{
                "trade_id": row.get("trade_id", row.get("id", "")),
                "slippage": row.get("slippage", row.get("cost_slippage", 0.0)),
                "fees": row.get("fees", row.get("cost_fees", 0.0)),
                "borrow": row.get("borrow", row.get("cost_borrow", 0.0)),
                "total_cost": row.get("total_cost", row.get("cost_total", 0.0)),
            } for row in trades[:400]]
        self._cost_drilldown_rows = cost_breakdown_rows
        self._set_tree_data(self.cost_breakdown_tree, cost_breakdown_rows)

        failure_rows = self._load_rows(run_dir, "failure_windows")
        if not failure_rows:
            failure_rows = [
                {
                    "window": row.get("window", idx),
                    "timestamp": row.get("timestamp", row.get("entry_time", "")),
                    "failure_reason": row.get("failure_reason", row.get("reject_reason", "")),
                    "drawdown": row.get("drawdown", ""),
                    "regime": row.get("regime", ""),
                }
                for idx, row in enumerate((risk_rows or drawdown_rows)[:300])
            ]
        self._failure_drilldown_rows = failure_rows
        self._set_tree_data(self.failure_window_tree, failure_rows)

        agg = aggregate_regime_market_stress(explain_rows, pnl_field="pnl", cost_field="pnl")
        self._last_results_cards_payload = {
            "run": run_dir.name,
            "alpha_ir": alpha_rows,
            "aggregates": agg,
            "cost_summary": {
                "cost_slippage": float(metrics.get("cost_slippage", 0.0)),
                "cost_fees": float(metrics.get("cost_fees", 0.0)),
                "cost_borrow": float(metrics.get("cost_borrow", 0.0)),
                "cost_total": float(metrics.get("cost_total", 0.0)),
            },
            "diagnostic_counts": {
                "trades": len(explain_rows),
                "cost_rows": len(cost_breakdown_rows),
                "failure_windows": len(failure_rows),
            },
        }
        self.results_summary_var.set(
            f"Loaded {run_dir.name}: equity={len(equity_rows)} rows, regimes={len(regime_rows)} rows, trades={len(explain_rows)} rows."
        )

    def _on_trade_drilldown_selected(self, _event: tk.Event) -> None:
        selected = self.trade_explain_tree.selection()
        if not selected:
            self._set_tree_data(self.cost_breakdown_tree, self._cost_drilldown_rows)
            self._set_tree_data(self.failure_window_tree, self._failure_drilldown_rows)
            return
        idx = self.trade_explain_tree.index(selected[0])
        row = self._trade_drilldown_rows[idx] if idx < len(self._trade_drilldown_rows) else {}
        trade_id = str(row.get("trade_id", "")).strip()
        ts = str(row.get("timestamp", "")).strip()
        filtered_cost = [r for r in self._cost_drilldown_rows if str(r.get("trade_id", "")).strip() == trade_id] if trade_id else self._cost_drilldown_rows
        filtered_failure = [r for r in self._failure_drilldown_rows if ts and ts in str(r.get("timestamp", ""))]
        self._set_tree_data(self.cost_breakdown_tree, filtered_cost or self._cost_drilldown_rows)
        self._set_tree_data(self.failure_window_tree, filtered_failure or self._failure_drilldown_rows)

    def _load_run_parameters(self, run_dir: Path) -> dict[str, object]:
        manifest = self._read_json(run_dir / "manifest.json")
        if isinstance(manifest, dict):
            params = manifest.get("parameters")
            if isinstance(params, dict):
                return dict(params)
        return {}

    def _render_selected_run_comparison(self) -> None:
        base_run = self._resolve_run_by_name(self.compare_base_var.get())
        other_run = self._resolve_run_by_name(self.compare_other_var.get())
        if base_run is None or other_run is None or base_run == other_run:
            return
        base_metrics = self._load_metric_map(base_run)
        other_metrics = self._load_metric_map(other_run)
        delta_rows = metric_deltas(base_metrics, other_metrics)
        self._set_tree_data(self.metric_delta_tree, delta_rows[:80])

        base_params = self._load_run_parameters(base_run)
        other_params = self._load_run_parameters(other_run)
        param_rows = parameter_diffs(base_params, other_params)
        self._set_tree_data(self.param_diff_tree, param_rows[:120])

        base_scenarios = read_stress_scenarios(base_run)
        other_scenarios = read_stress_scenarios(other_run)
        scenario_rows = build_scenario_comparison(base_scenarios, other_scenarios)
        self._set_tree_data(self.scenario_compare_tree, scenario_rows[:120])

        base_folds = self._read_json(base_run / "fold_summary.json")
        other_folds = self._read_json(other_run / "fold_summary.json")
        fold_rows = fold_variance_rows(
            [row for row in base_folds if isinstance(row, dict)] if isinstance(base_folds, list) else [],
            [row for row in other_folds if isinstance(row, dict)] if isinstance(other_folds, list) else [],
        )
        self._set_tree_data(self.fold_variance_tree, fold_rows[:120])

        base_frontier = self._read_json(base_run / "robustness_frontier.json")
        other_frontier = self._read_json(other_run / "robustness_frontier.json")
        frontier_rows = compare_robustness_frontiers(
            base_frontier if isinstance(base_frontier, dict) else {},
            other_frontier if isinstance(other_frontier, dict) else {},
        )
        self._set_tree_data(self.frontier_compare_tree, frontier_rows[:120])

        base_manifest = self._read_json(base_run / "manifest.json")
        other_manifest = self._read_json(other_run / "manifest.json")
        cmp = compare_manifests(base_manifest if isinstance(base_manifest, dict) else {}, other_manifest if isinstance(other_manifest, dict) else {})
        manifest_rows: list[dict[str, object]] = []
        for row in cmp.get("parameter_diffs", []):
            if isinstance(row, dict):
                manifest_rows.append({"section": "parameters", **row})
        for row in cmp.get("dependency_diffs", []):
            if isinstance(row, dict):
                manifest_rows.append({"section": "dependency_versions", **row})
        for row in cmp.get("metric_table_diffs", []):
            if isinstance(row, dict):
                manifest_rows.append({"section": "metric_tables", **row})
        manifest_rows.append({
            "section": "provenance",
            "parameter": "config_hash_changed",
            "base": False,
            "compare": bool(cmp.get("config_hash_changed", False)),
        })
        manifest_rows.append({
            "section": "provenance",
            "parameter": "reproducibility_fingerprint_changed",
            "base": False,
            "compare": bool(cmp.get("reproducibility_fingerprint_changed", False)),
        })
        self._set_tree_data(self.manifest_diff_tree, manifest_rows[:200])
        trained_base = str((base_manifest or {}).get("run_type", "")) == "trained_regime"
        trained_other = str((other_manifest or {}).get("run_type", "")) == "trained_regime"
        focus = "trained-regime A/B" if trained_base and trained_other else "generic A/B"
        self.compare_manifest_summary_var.set(
            f"{focus}: params={len(cmp.get('parameter_diffs', []))}, dependency_diffs={len(cmp.get('dependency_diffs', []))}, metric_table_diffs={len(cmp.get('metric_table_diffs', []))}."
        )

    def _render_guardrails(self, run_dir: Path) -> None:
        for child in self.guardrail_frame.winfo_children():
            child.destroy()
        metrics = self._load_metric_map(run_dir)
        fold_rows = self._read_json(run_dir / "fold_summary.json")
        rows = fold_rows if isinstance(fold_rows, list) else None
        trade_count = int(metrics.get("trade_count", 0.0)) if "trade_count" in metrics else None
        robustness = self._read_json(run_dir / "robustness_report.json")
        robustness_payload = robustness if isinstance(robustness, dict) else None
        evidence_links = {
            "overfit_risk": str(run_dir / "robustness_report.json"),
            "low_sample": str(run_dir / "trades.csv"),
            "high_turnover": str(run_dir / "turnover_by_symbol.csv"),
            "unstable_params": str(run_dir / "fold_summary.json"),
            "weak_rc": str(run_dir / "robustness_report.json"),
            "weak_spa": str(run_dir / "robustness_report.json"),
            "alpha_not_robust": str(run_dir / "robustness_report.json"),
            "default": str(run_dir / "manifest.json"),
            "robustness_frontier": str(run_dir / "robustness_frontier.json"),
        }
        badges = build_guardrails(metrics, fold_rows=rows, trade_count=trade_count, robustness=robustness_payload, evidence_links=evidence_links)
        scenario_payload = read_stress_scenarios(run_dir)
        scenario_checks = scenario_payload.get("scenario_guardrails", []) if isinstance(scenario_payload, dict) else []
        if isinstance(scenario_checks, list):
            failed = [row for row in scenario_checks if isinstance(row, dict) and not bool(row.get("passed", False))]
            if failed:
                badges.append({"label": "Stress Failures", "severity": "high", "reason": f"{len(failed)} scenario guardrail checks failed."})
        manifest = self._read_json(run_dir / "manifest.json")
        if isinstance(manifest, dict):
            governance = manifest.get("governance", {}) if isinstance(manifest.get("governance"), dict) else {}
            drift_monitoring = governance.get("drift_monitoring", {}) if isinstance(governance.get("drift_monitoring"), dict) else {}
            if drift_monitoring:
                if bool(drift_monitoring.get("within_tolerance", False)):
                    badges.append({"label": "Drift Monitor OK", "severity": "low", "reason": "Signal/fill/PnL drift are within governance tolerances."})
                for alert in drift_monitoring.get("alert_summaries", []):
                    if isinstance(alert, dict):
                        badges.append({"label": "Drift Alert", "severity": str(alert.get("severity", "high")), "reason": str(alert.get("summary", "Drift tolerance breached."))})
            gate_checks = governance.get("gate_checks", {}) if isinstance(governance.get("gate_checks"), dict) else {}
            missing_checks = governance.get("missing_required_checks", []) if isinstance(governance.get("missing_required_checks"), list) else []
            governance_diagnostics = governance.get("governance_diagnostics", {}) if isinstance(governance.get("governance_diagnostics"), dict) else {}
            dsr_diag = governance_diagnostics.get("deflated_sharpe_reality_check", {}) if isinstance(governance_diagnostics.get("deflated_sharpe_reality_check"), dict) else {}
            stability_diag = governance_diagnostics.get("parameter_stability", {}) if isinstance(governance_diagnostics.get("parameter_stability"), dict) else {}
            split_drift_diag = governance_diagnostics.get("train_validation_test_drift", {}) if isinstance(governance_diagnostics.get("train_validation_test_drift"), dict) else {}
            if gate_checks:
                if gate_checks.get("deflated_sharpe_reality_check", False):
                    badges.append({"label": "DSR/RC OK", "severity": "low", "reason": "Deflated Sharpe and reality-check diagnostics passed."})
                else:
                    badges.append({"label": "DSR/RC Breach", "severity": "high", "reason": f"Deflated Sharpe/reality-check failed (DSR={float(dsr_diag.get('deflated_sharpe_ratio', 0.0) or 0.0):.2f}, p={float(dsr_diag.get('combined_reality_check_pvalue', 1.0) or 1.0):.2f})."})
                if gate_checks.get("parameter_stability_penalty", False):
                    badges.append({"label": "Stability Penalty OK", "severity": "low", "reason": f"Parameter stability penalty within tolerance ({float(stability_diag.get('parameter_stability_penalty', 0.0) or 0.0):.2f})."})
                else:
                    badges.append({"label": "Stability Penalty", "severity": "medium", "reason": f"Parameter stability penalty exceeded threshold ({float(stability_diag.get('parameter_stability_penalty', 1.0) or 1.0):.2f})."})
                if gate_checks.get("train_validation_test_drift", False):
                    badges.append({"label": "Split Drift OK", "severity": "low", "reason": "Train/validation/test performance drift is within tolerance."})
                else:
                    badges.append({"label": "Split Drift", "severity": "medium", "reason": f"Split drift breached (tv={float(split_drift_diag.get('train_validation_abs_drift', 0.0) or 0.0):.2f}, vt={float(split_drift_diag.get('validation_test_abs_drift', 0.0) or 0.0):.2f})."})
            if missing_checks:
                badges.append({"label": "Missing Required Checks", "severity": "high", "reason": ", ".join(str(item) for item in missing_checks)})
        palette = {"high": "#d9534f", "medium": "#f0ad4e", "low": "#5cb85c"}
        for badge in badges:
            artifact = str(badge.get("artifact", ""))
            suffix = f" [evidence: {Path(artifact).name}]" if artifact else ""
            label = ttk.Label(
                self.guardrail_frame,
                text=f"{badge['label']}: {badge['reason']}{suffix}",
                background=palette.get(str(badge.get("severity", "low")), "#5cb85c"),
                foreground="#ffffff",
                padding=(6, 2),
                cursor="hand2" if artifact else "",
            )
            if artifact:
                label.bind("<Button-1>", lambda _e, p=artifact: webbrowser.open(f"file://{Path(p).resolve()}"))
            label.pack(side="right", padx=3)

    def _draw_line_canvas(self, canvas: tk.Canvas, values: list[float], color: str = "#1f77b4") -> None:
        canvas.delete("all")
        w = max(20, int(canvas.winfo_width()))
        h = max(20, int(canvas.winfo_height()))
        if len(values) < 2:
            canvas.create_text(w // 2, h // 2, text="No data", fill="#777")
            return
        min_v = min(values)
        max_v = max(values)
        if max_v <= min_v:
            max_v = min_v + 1.0
        points: list[float] = []
        for idx, value in enumerate(values):
            x = 10 + (w - 20) * (idx / max(1, len(values) - 1))
            y = 8 + (h - 16) * (1.0 - (value - min_v) / (max_v - min_v))
            points.extend([x, y])
        canvas.create_line(*points, fill=color, width=2.0)

    def _draw_bar_canvas(self, canvas: tk.Canvas, rows: list[tuple[str, float]], color: str = "#2ca02c") -> None:
        canvas.delete("all")
        w = max(20, int(canvas.winfo_width()))
        h = max(20, int(canvas.winfo_height()))
        if not rows:
            canvas.create_text(w // 2, h // 2, text="No data", fill="#777")
            return
        values = [abs(v) for _n, v in rows]
        max_v = max(values) if values else 1.0
        if max_v <= 0:
            max_v = 1.0
        bar_h = max(12, (h - 16) // max(1, len(rows)))
        y = 8
        for name, value in rows[:10]:
            width = int((w - 120) * (abs(value) / max_v))
            canvas.create_rectangle(110, y, 110 + width, y + bar_h - 2, fill=color, outline=color)
            canvas.create_text(6, y + bar_h / 2, anchor="w", text=name[:14], fill="#333")
            canvas.create_text(114 + width, y + bar_h / 2, anchor="w", text=f"{value:.4f}", fill="#333")
            y += bar_h

    def _render_parameter_stability(self, run_dir: Path) -> None:
        leaderboard_rows = self._load_rows(run_dir, "leaderboard")
        points: list[tuple[float, float]] = []
        for row in leaderboard_rows:
            x = self._safe_float(row.get("lookback_days"))
            y = self._safe_float(row.get("sharpe"))
            if x is not None and y is not None:
                points.append((x, y))
        self.heatmap_canvas.delete("all")
        w = max(20, int(self.heatmap_canvas.winfo_width()))
        h = max(20, int(self.heatmap_canvas.winfo_height()))
        if not points:
            self.heatmap_canvas.create_text(w // 2, h // 2, text="No parameter sweep data for heatmap")
        else:
            xs = [p[0] for p in points]
            ys = [p[1] for p in points]
            xmin, xmax = min(xs), max(xs)
            ymin, ymax = min(ys), max(ys)
            if xmax <= xmin:
                xmax = xmin + 1
            if ymax <= ymin:
                ymax = ymin + 1
            for x, y in points[:300]:
                px = 14 + (w - 28) * ((x - xmin) / (xmax - xmin))
                py = 10 + (h - 20) * (1.0 - (y - ymin) / (ymax - ymin))
                self.heatmap_canvas.create_oval(px - 3, py - 3, px + 3, py + 3, fill="#1f77b4", outline="")
        fold_rows = self._read_json(run_dir / "fold_summary.json")
        seq: list[float] = []
        if isinstance(fold_rows, list):
            for row in fold_rows:
                if isinstance(row, dict):
                    params = row.get("selected_params")
                    if isinstance(params, dict):
                        seq.append(float(len(params)))
        self._draw_line_canvas(self.stability_canvas, seq, color="#9467bd")

    def _on_fold_selected(self) -> None:
        selection = self.wf_tree.selection()
        if not selection:
            return
        idx = self.wf_tree.index(selection[0])
        if idx >= len(self._last_fold_rows):
            return
        row = self._last_fold_rows[idx]
        params = row.get("selected_params") if isinstance(row, dict) else None
        diag = row.get("diagnostics") if isinstance(row, dict) else None
        equity = row.get("oos_equity") if isinstance(row, dict) else None

        param_rows: list[dict[str, object]] = []
        if isinstance(params, dict):
            param_rows = [{"parameter": k, "value": v} for k, v in sorted(params.items())]
        self._set_tree_data(self.wf_params_tree, param_rows)

        diag_rows: list[dict[str, object]] = []
        if isinstance(diag, dict):
            diag_rows = [{"diagnostic": k, "value": v} for k, v in sorted(diag.items())]
        self._set_tree_data(self.wf_diag_tree, diag_rows)

        series: list[float] = []
        if isinstance(equity, list):
            for item in equity:
                if isinstance(item, dict):
                    value = self._safe_float(item.get("equity"))
                    if value is not None:
                        series.append(value)
        self._draw_line_canvas(self.wf_equity_canvas, series, color="#d62728")

    def _render_cost_attribution(self, run_dir: Path) -> None:
        metrics = self._load_metric_map(run_dir)
        cost_rows = [
            ("slippage", float(metrics.get("cost_slippage", 0.0))),
            ("fees", float(metrics.get("cost_fees", 0.0))),
            ("carry", float(metrics.get("cost_borrow", 0.0))),
            ("total", float(metrics.get("cost_total", 0.0))),
        ]
        self._draw_bar_canvas(self.cost_canvas, cost_rows, color="#ff7f0e")
        slippage_payload = self._read_json(run_dir / "slippage_decomposition.json")
        if isinstance(slippage_payload, dict):
            drift = float(slippage_payload.get("expected_vs_observed_fill_slippage_drift_bps", 0.0) or 0.0)
            regime_rows = slippage_payload.get("by_regime", []) if isinstance(slippage_payload.get("by_regime"), list) else []
            liquidity_rows = slippage_payload.get("by_liquidity_bucket", []) if isinstance(slippage_payload.get("by_liquidity_bucket"), list) else []
            top_regime = regime_rows[0] if regime_rows else {}
            top_bucket = liquidity_rows[0] if liquidity_rows else {}
            self.slippage_summary_var.set(
                f"Slippage drift card → total drift: {drift:.2f} bps | regimes: {len(regime_rows)} | liquidity buckets: {len(liquidity_rows)} | "
                f"sample regime={top_regime.get('regime', 'n/a')} drift={float(top_regime.get('drift_bps', 0.0) or 0.0):.2f} bps | "
                f"sample bucket={top_bucket.get('liquidity_bucket', 'n/a')} drift={float(top_bucket.get('drift_bps', 0.0) or 0.0):.2f} bps"
            )
        else:
            self.slippage_summary_var.set("Slippage decomposition card unavailable for this run.")

        turnover_rows = self._load_rows(run_dir, "turnover_by_symbol")
        aggregate: dict[str, float] = {}
        for row in turnover_rows:
            sym = str(row.get("symbol", ""))
            val = self._safe_float(row.get("turnover"))
            if not sym or val is None:
                continue
            aggregate[sym] = aggregate.get(sym, 0.0) + val
        ordered = sorted(aggregate.items(), key=lambda kv: kv[1], reverse=True)
        self._draw_bar_canvas(self.turnover_canvas, ordered[:10], color="#2ca02c")

        regime_rows = self._load_rows(run_dir, "regime_pnl_attribution")
        self._set_tree_data(self.regime_tree, regime_rows[:120])

        drawdown_rows = self._load_rows(run_dir, "drawdown")
        decomposition: list[dict[str, object]] = []
        for idx, row in enumerate(drawdown_rows[:50]):
            dd = self._safe_float(row.get("drawdown"))
            if dd is None:
                continue
            decomposition.append({
                "rank": idx + 1,
                "timestamp": row.get("timestamp", ""),
                "drawdown": dd,
                "depth_pct": abs(dd) * 100.0,
            })
        self._set_tree_data(self.drawdown_decomp_tree, decomposition)

        frontier_rows = self._read_json(run_dir / "capacity_frontier.json")
        if isinstance(frontier_rows, list):
            points = []
            for row in frontier_rows:
                if not isinstance(row, dict):
                    continue
                x = self._safe_float(row.get("aum_scale"))
                y = self._safe_float(row.get("expected_alpha_net_cost_bps"))
                if x is not None and y is not None:
                    points.append((x, y))
            points.sort(key=lambda item: item[0])
            self._draw_line_canvas(self.capacity_canvas, [val for _, val in points], color="#1f77b4")
        else:
            self._draw_line_canvas(self.capacity_canvas, [], color="#1f77b4")

    def _refresh_wf_fraction_labels(self) -> None:
        self.wf_train_fraction_label.config(text=f"{float(self.wf_train_fraction_var.get()):.2f}")
        self.wf_validation_fraction_label.config(text=f"{float(self.wf_validation_fraction_var.get()):.2f}")
        self.wf_test_fraction_label.config(text=f"{float(self.wf_test_fraction_var.get()):.2f}")
        self.wf_step_fraction_label.config(text=f"{float(self.wf_step_fraction_var.get()):.2f}")

    def _on_wf_fraction_changed(self, changed: str) -> None:
        if self._updating_wf_fractions:
            return
        values = {
            "train": max(0.01, float(self.wf_train_fraction_var.get())),
            "validation": max(0.01, float(self.wf_validation_fraction_var.get())),
            "test": max(0.01, float(self.wf_test_fraction_var.get())),
        }
        total = values["train"] + values["validation"] + values["test"]
        if total <= 0:
            return
        if abs(total - 1.0) < 1e-6:
            self._refresh_wf_fraction_labels()
            return

        others = [key for key in values if key != changed]
        other_total = values[others[0]] + values[others[1]]
        target_other_total = max(0.02, 1.0 - values[changed])

        if other_total <= 0:
            values[others[0]] = target_other_total / 2.0
            values[others[1]] = target_other_total / 2.0
        else:
            scale = target_other_total / other_total
            values[others[0]] *= scale
            values[others[1]] *= scale

        # clamp and renormalize
        for key in values:
            values[key] = min(0.98, max(0.01, values[key]))
        renorm = values["train"] + values["validation"] + values["test"]
        values["train"] /= renorm
        values["validation"] /= renorm
        values["test"] /= renorm

        self._updating_wf_fractions = True
        try:
            self.wf_train_fraction_var.set(values["train"])
            self.wf_validation_fraction_var.set(values["validation"])
            self.wf_test_fraction_var.set(values["test"])
        finally:
            self._updating_wf_fractions = False
        self._refresh_wf_fraction_labels()

    def _build_risk_limit_slider_row(self, parent: ttk.Frame, row: int, label: str, key: str) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=8, pady=4)
        minimum, maximum, resolution = self._risk_limit_ranges[key]
        slider = tk.Scale(
            parent,
            from_=minimum,
            to=maximum,
            resolution=resolution,
            orient="horizontal",
            showvalue=False,
            variable=self._risk_limit_slider_vars[key],
            command=lambda value, risk_key=key: self._on_risk_limit_slider_changed(risk_key, value),
        )
        slider.grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        entry = ttk.Entry(parent, textvariable=self._risk_limit_vars[key], width=14)
        entry.grid(row=row, column=2, sticky="w", padx=8, pady=4)
        self._risk_limit_vars[key].trace_add("write", lambda *_args, risk_key=key: self._on_risk_limit_entry_changed(risk_key))

    def _format_risk_limit_value(self, key: str, value: float) -> str:
        if key == "max_participation_rate":
            return f"{value:.2f}"
        if key == "portfolio_max_net_gamma":
            return f"{value:.3f}".rstrip("0").rstrip(".")
        return f"{value:.1f}".rstrip("0").rstrip(".")

    def _on_risk_limit_slider_changed(self, key: str, raw_value: str) -> None:
        if getattr(self, "_updating_risk_limit_controls", False):
            return
        self._updating_risk_limit_controls = True
        try:
            self._risk_limit_vars[key].set(self._format_risk_limit_value(key, float(raw_value)))
        finally:
            self._updating_risk_limit_controls = False
        self._regime_backtest_options: list[RegimeBacktestOption] = []
        self._regime_option_lookup: dict[str, RegimeBacktestOption] = {}
        self._active_regime_contract: RegimeBacktestContract | None = None
        self._regime_loading_defaults = False
        self._regime_locked_fields = True
        self._regime_loaded_values: dict[str, str] = {}

    def _on_risk_limit_entry_changed(self, key: str) -> None:
        if getattr(self, "_updating_risk_limit_controls", False):
            return
        text = self._risk_limit_vars[key].get().strip()
        if not text:
            return
        parsed = parse_float(text)
        if parsed is None:
            return
        minimum, maximum, _resolution = self._risk_limit_ranges[key]
        clamped = max(minimum, min(maximum, float(parsed)))
        self._updating_risk_limit_controls = True
        try:
            self._risk_limit_slider_vars[key].set(clamped)
        finally:
            self._updating_risk_limit_controls = False
        self._regime_backtest_options: list[RegimeBacktestOption] = []
        self._regime_option_lookup: dict[str, RegimeBacktestOption] = {}
        self._active_regime_contract: RegimeBacktestContract | None = None
        self._regime_loading_defaults = False
        self._regime_locked_fields = True
        self._regime_loaded_values: dict[str, str] = {}

    def _apply_options_risk_limit_preset(self, preset_name: str) -> None:
        presets = {
            "conservative": {
                "portfolio_max_net_gamma": 0.50,
                "portfolio_max_abs_vega_bucket": 2_500.0,
                "portfolio_max_abs_delta_per_underlying": 600.0,
                "max_participation_rate": 0.10,
            },
            "balanced": {
                "portfolio_max_net_gamma": 1.25,
                "portfolio_max_abs_vega_bucket": 7_500.0,
                "portfolio_max_abs_delta_per_underlying": 1_800.0,
                "max_participation_rate": 0.25,
            },
            "aggressive": {
                "portfolio_max_net_gamma": 2.50,
                "portfolio_max_abs_vega_bucket": 15_000.0,
                "portfolio_max_abs_delta_per_underlying": 3_500.0,
                "max_participation_rate": 0.45,
            },
        }
        preset = presets.get(preset_name)
        if not preset:
            return
        self._updating_risk_limit_controls = True
        try:
            for key, value in preset.items():
                self._risk_limit_slider_vars[key].set(float(value))
                self._risk_limit_vars[key].set(self._format_risk_limit_value(key, float(value)))
        finally:
            self._updating_risk_limit_controls = False
        self._regime_backtest_options: list[RegimeBacktestOption] = []
        self._regime_option_lookup: dict[str, RegimeBacktestOption] = {}
        self._active_regime_contract: RegimeBacktestContract | None = None
        self._regime_loading_defaults = False
        self._regime_locked_fields = True
        self._regime_loaded_values: dict[str, str] = {}

    def _build_momentum_options(self, parent: ttk.Frame) -> None:
        self.momentum_options_frame = ttk.LabelFrame(parent, text="Momentum Options")
        self.momentum_options_frame.grid(row=0, column=0, sticky="ew", padx=8, pady=6)

        ttk.Label(self.momentum_options_frame, text="Entry Signals").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        entry_row = ttk.Frame(self.momentum_options_frame)
        entry_row.grid(row=0, column=1, sticky="ew", padx=8, pady=6)
        self.entry_signal_vars: dict[str, tk.BooleanVar] = {}
        for idx, name in enumerate(ENTRY_SIGNALS):
            var = tk.BooleanVar(value=False)
            self.entry_signal_vars[name] = var
            ttk.Checkbutton(entry_row, text=name, variable=var).grid(row=0, column=idx, sticky="w", padx=(0, 10))

        ttk.Label(self.momentum_options_frame, text="Exit Signals").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        exit_row = ttk.Frame(self.momentum_options_frame)
        exit_row.grid(row=1, column=1, sticky="ew", padx=8, pady=6)
        self.exit_signal_vars: dict[str, tk.BooleanVar] = {}
        for idx, name in enumerate(EXIT_SIGNALS):
            var = tk.BooleanVar(value=False)
            self.exit_signal_vars[name] = var
            ttk.Checkbutton(exit_row, text=name, variable=var).grid(row=0, column=idx, sticky="w", padx=(0, 10))

    def _build_xsmom_options(self, parent: ttk.Frame) -> None:
        self.xsmom_options_frame = ttk.LabelFrame(parent, text="Cross-Sectional Momentum Options")
        self.xsmom_options_frame.grid(row=0, column=0, sticky="ew", padx=8, pady=6)
        self.xsmom_options_frame.columnconfigure(1, weight=1)

        ttk.Label(self.xsmom_options_frame, text="Top Quantile (0-1)").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        self.xsmom_top_quantile_var = tk.StringVar()
        ttk.Entry(self.xsmom_options_frame, textvariable=self.xsmom_top_quantile_var).grid(row=0, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(self.xsmom_options_frame, text="Bottom Quantile (0-1)").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        self.xsmom_bottom_quantile_var = tk.StringVar()
        ttk.Entry(self.xsmom_options_frame, textvariable=self.xsmom_bottom_quantile_var).grid(row=1, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(self.xsmom_options_frame, text="Volatility Lookback (bars)").grid(row=2, column=0, sticky="w", padx=8, pady=6)
        self.xsmom_vol_lookback_days_var = tk.StringVar()
        ttk.Entry(self.xsmom_options_frame, textvariable=self.xsmom_vol_lookback_days_var).grid(row=2, column=1, sticky="ew", padx=8, pady=6)

        self.xsmom_long_only_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            self.xsmom_options_frame,
            text="Long Only",
            variable=self.xsmom_long_only_var,
        ).grid(row=3, column=0, columnspan=2, sticky="w", padx=8, pady=6)

    def _on_strategy_changed(self, _event: object | None = None) -> None:
        strategy = self.strategy_var.get().strip() or "momentum"
        if strategy == "xsmom":
            self.momentum_options_frame.grid_remove()
            self.xsmom_options_frame.grid()
            self.walk_forward_frame.grid_remove()
        else:
            self.xsmom_options_frame.grid_remove()
            self.momentum_options_frame.grid()
            self.walk_forward_frame.grid()
        self._update_validation_hint()

    def _register_advanced_widgets(self) -> None:
        self._advanced_widgets = {
            "custom_bet_pct": self.custom_bet_pct_entry,
            "portfolio_method": self.portfolio_method_combo,
            "portfolio_vol_lookback": self.portfolio_vol_lookback_entry,
            "portfolio_target_vol": self.portfolio_target_vol_entry,
            "portfolio_max_symbol": self.portfolio_max_symbol_entry,
            "portfolio_max_sector": self.portfolio_max_sector_entry,
            "portfolio_rebalance_frequency": self.portfolio_rebalance_frequency_entry,
            "portfolio_clustering_linkage": self.portfolio_clustering_linkage_combo,
            "portfolio_covariance_shrinkage": self.portfolio_covariance_shrinkage_entry,
            "portfolio_max_gross": self.portfolio_max_gross_entry,
            "portfolio_min_net": self.portfolio_min_net_entry,
            "portfolio_max_net": self.portfolio_max_net_entry,
            "use_walk_forward": self.use_walk_forward_check,
            "use_optimizer": self.use_optimizer_check,
            "walk_forward_frame": self.walk_forward_frame,
        }
        self._advanced_tooltips = {
            "custom_bet_pct": "Portfolio domain: custom bet size percentage used when Bet Sizing Mode is custom_pct.",
            "portfolio_method": "Portfolio domain: weighting/optimization method (equal weight, vol targeting, risk parity variants, etc.).",
            "portfolio_vol_lookback": "Portfolio domain: lookback window for realized volatility and covariance estimation.",
            "portfolio_target_vol": "Portfolio domain: annualized volatility target used by vol-targeting methods.",
            "portfolio_max_symbol": "Portfolio domain: per-symbol max weight cap.",
            "portfolio_max_sector": "Portfolio domain: per-sector max weight cap.",
            "portfolio_rebalance_frequency": "Portfolio domain: rebalance cadence measured in bars.",
            "portfolio_clustering_linkage": "Portfolio domain: linkage method used for clustered allocators (HRP/HERC).",
            "portfolio_covariance_shrinkage": "Portfolio domain: covariance shrinkage strength (0=no shrinkage, 1=full shrinkage).",
            "portfolio_max_gross": "Portfolio domain: gross exposure ceiling.",
            "portfolio_min_net": "Portfolio domain: minimum net exposure bound.",
            "portfolio_max_net": "Portfolio domain: maximum net exposure bound.",
            "use_walk_forward": "CV domain: enables walk-forward train/validation/test sequencing.",
            "use_optimizer": "Optimization domain: enables multi-objective search over signal/parameter combinations.",
            "walk_forward_frame": "CV domain: full cross-validation, CPCV, nested optimization, and lineage controls.",
        }
        for key, widget in self._advanced_widgets.items():
            help_text = self._advanced_tooltips.get(key)
            if help_text:
                self._attach_tooltip(widget, help_text)

    def _bind_validation_watchers(self) -> None:
        for var in (
            self.lookback_days_var,
            self.skip_days_var,
            self.timeframe_var,
            self.start_date_var,
            self.end_date_var,
            self.strategy_var,
            self.use_walk_forward_var,
            self.use_optimizer_var,
            self.wf_train_bars_var,
            self.wf_validation_bars_var,
            self.wf_test_bars_var,
            self.wf_step_bars_var,
        ):
            var.trace_add("write", lambda *_args: self._update_validation_hint())

    def _on_mode_changed(self) -> None:
        is_advanced = self.ui_mode_var.get() == "advanced"
        self.show_advanced_controls_var.set(is_advanced)
        state = "normal" if is_advanced else "disabled"
        for key, widget in self._advanced_widgets.items():
            if key == "walk_forward_frame":
                if is_advanced and self.strategy_var.get().strip() == "momentum":
                    widget.grid()
                else:
                    widget.grid_remove()
                continue
            try:
                widget.config(state=state)
            except tk.TclError:
                pass
        if not is_advanced:
            self.use_walk_forward_var.set(False)
        self._update_validation_hint()

    def _on_show_advanced_controls_toggled(self) -> None:
        self.ui_mode_var.set("advanced" if self.show_advanced_controls_var.get() else "basic")
        self._on_mode_changed()

    def _update_validation_hint(self) -> bool:
        messages: list[str] = []
        disable_run = False
        self._update_ticker_universe_source()
        timeframe = self.timeframe_var.get().strip() or "1m"
        start_date = parse_date(self.start_date_var.get())
        end_date = parse_date(self.end_date_var.get())

        if timeframe in TIMEFRAME_HISTORY_DAYS and start_date is not None and end_date is not None and start_date < end_date:
            requested_days = (end_date - start_date).days
            needed_days = TIMEFRAME_HISTORY_DAYS[timeframe]
            if requested_days < needed_days:
                messages.append(
                    f"Selected window has {requested_days} days; {timeframe} commonly needs {needed_days}+ days of history."
                )

        if bool(self.use_walk_forward_var.get()) and self.ui_mode_var.get() != "advanced":
            messages.append("Walk-forward requires Advanced mode.")
            disable_run = True

        has_bar_values = any(
            (
                self.wf_train_bars_var.get().strip(),
                self.wf_validation_bars_var.get().strip(),
                self.wf_test_bars_var.get().strip(),
                self.wf_step_bars_var.get().strip(),
            )
        )
        self.wf_mode_hint_var.set(
            "Bars mode active: provide all train/validation/test/step bars and fractions will be ignored."
            if has_bar_values
            else "Fractions mode active: leave all bars fields blank."
        )

        if bool(self.use_optimizer_var.get()) and self.ui_mode_var.get() != "advanced":
            messages.append("Optimizer requires Advanced mode.")
            disable_run = True

        messages.extend(self._stale_preset_messages)
        self._validation_messages = messages
        self.validation_hint_var.set("\n".join(messages))
        self._set_run_controls_state("disabled" if disable_run else "normal")
        return disable_run

    def _toggle_regime_advanced_overrides(self) -> None:
        if self.regime_overrides_expanded_var.get():
            self.regime_advanced_overrides_frame.grid(row=7, column=0, columnspan=2, sticky="ew", padx=6, pady=(0, 4))
        else:
            self.regime_advanced_overrides_frame.grid_remove()

    def _update_ticker_universe_source(self) -> None:
        tickers = [str(item).strip().upper() for item in self.controller.state.tickers if str(item).strip()]
        if not tickers:
            self.regime_ticker_source_var.set("Ticker Entry page (0 symbols loaded)")
            return
        preview = ", ".join(tickers[:5])
        suffix = "" if len(tickers) <= 5 else ", …"
        self.regime_ticker_source_var.set(f"Ticker Entry page ({len(tickers)} symbols): {preview}{suffix}")

    def _run_exact_training_window_replay(self) -> None:
        if self._active_regime_contract is None:
            messagebox.showinfo("Missing trained regime", "Select a trained regime artifact before using quick actions.")
            return
        self.backtest_type_var.set("Trained Regime")
        start_date = str(self._active_regime_contract.defaults.get("start_date", "")).strip()
        end_date = str(self._active_regime_contract.defaults.get("end_date", "")).strip()
        if start_date and end_date:
            self.start_date_var.set(start_date)
            self.end_date_var.set(end_date)
        self.run_full_chain()

    def _run_full_history_replay(self) -> None:
        if self._active_regime_contract is None:
            messagebox.showinfo("Missing trained regime", "Select a trained regime artifact before using quick actions.")
            return
        self.backtest_type_var.set("Trained Regime")
        self.start_date_var.set("")
        self.end_date_var.set("")
        self.run_full_chain()

    def _show_regime_compatibility_details(self) -> None:
        contract = self._active_regime_contract
        if contract is None:
            messagebox.showinfo(
                "Compatibility troubleshooting",
                "Load a trained regime artifact to inspect replay compatibility diagnostics.",
            )
            return
        details = dict(contract.compatibility_metadata or {})
        details_text = json.dumps(details, indent=2, sort_keys=True) if details else "No compatibility metadata was attached to this artifact."
        messagebox.showinfo(
            "Compatibility troubleshooting",
            f"Reproducibility: {contract.reproducibility_status}\n"
            f"Manifest: {contract.manifest_path}\n\n"
            f"Diagnostics:\n{details_text}",
        )

    def _refresh_template_choices(self) -> None:
        names = sorted(self.controller.state.backtest_templates.keys())
        self.template_combo.configure(values=names)
        selected = self.template_var.get().strip()
        if selected and selected not in names:
            self.template_var.set("")

    def _refresh_trained_regime_choices(self) -> None:
        self._regime_backtest_options = discover_regime_backtest_options(self.controller.state.regime_training_runs)
        labels = ["(none)"] + [item.label for item in self._regime_backtest_options]
        self._regime_option_lookup = {item.label: item for item in self._regime_backtest_options}
        self.trained_regime_combo.configure(values=labels)
        selected = self.trained_regime_var.get().strip()
        if selected and selected not in labels:
            self.trained_regime_var.set("(none)")

    def _normalized_backtest_type(self) -> str:
        selected = self.backtest_type_var.get().strip()
        if selected in BACKTEST_WORKFLOW_TYPE_MAP:
            return BACKTEST_WORKFLOW_TYPE_MAP[selected]
        normalized = selected.lower().replace(" ", "_")
        if normalized in set(BACKTEST_WORKFLOW_TYPES):
            return normalized
        return "classic_strategy"

    def _selected_regime_option(self) -> RegimeBacktestOption | None:
        label = self.trained_regime_var.get().strip()
        option = self._regime_option_lookup.get(label)
        if option is not None:
            return option
        selected_manifest = str(self.controller.state.backtest_settings.get("selected_trained_regime_manifest_path", "")).strip()
        selected_source = str(self.controller.state.backtest_settings.get("selected_trained_regime_source", "")).strip()
        selected_option_id = str(self.controller.state.backtest_settings.get("selected_trained_regime_option_id", "")).strip()
        for candidate in self._regime_backtest_options:
            if selected_option_id and candidate.option_id == selected_option_id:
                return candidate
            if selected_manifest and candidate.manifest_path == selected_manifest and (not selected_source or candidate.source == selected_source):
                return candidate
        return None

    def _on_trained_regime_selected(self, _event: object | None = None) -> None:
        selected = self.trained_regime_var.get().strip()
        if not selected or selected == "(none)":
            self._active_regime_contract = None
            self._regime_loaded_values = {}
            self.regime_provenance_var.set("No trained regime loaded.")
            self.regime_diff_var.set("")
            self._regime_immutable_values = {}
            if hasattr(self, "regime_immutable_reason_var"):
                self.regime_immutable_reason_var.set("Immutable fields: select a trained regime to inspect replay constraints.")
            self._apply_regime_lock_state()
            return
        option = self._selected_regime_option()
        if option is None:
            return
        try:
            contract = load_regime_backtest_contract(option)
        except RegimeBundleCompatibilityError as exc:
            messagebox.showerror(
                "Regime bundle compatibility error",
                f"Cannot load selected trained regime.\n\n{exc}",
            )
            self.trained_regime_var.set("(none)")
            self._active_regime_contract = None
            self._regime_loaded_values = {}
            self.regime_provenance_var.set("No trained regime loaded.")
            self.regime_diff_var.set("")
            self._regime_immutable_values = {}
            if hasattr(self, "regime_immutable_reason_var"):
                self.regime_immutable_reason_var.set("Immutable fields: select a trained regime to inspect replay constraints.")
            self._apply_regime_lock_state()
            return
        self._active_regime_contract = contract
        self._apply_regime_contract_defaults(contract)

    def _apply_regime_contract_defaults(self, contract: RegimeBacktestContract) -> None:
        var_map = self._regime_override_field_var_map()
        self._regime_loading_defaults = True
        try:
            for key, value in contract.defaults.items():
                var = var_map.get(key)
                if var is None:
                    continue
                var.set(str(value))
            selected_packs = [item.strip() for item in str(contract.defaults.get("selected_scenario_packs", "")).split(",") if item.strip()]
            self._set_listbox_selection(self.scenario_pack_listbox, selected_packs, valid_options=self._scenario_pack_options)
        finally:
            self._regime_loading_defaults = False

        if str(contract.defaults.get("start_date", "")).strip():
            self.start_date_var.set(str(contract.defaults.get("start_date", "")).strip())
        if str(contract.defaults.get("end_date", "")).strip():
            self.end_date_var.set(str(contract.defaults.get("end_date", "")).strip())

        self._regime_loaded_values = {
            key: str(var.get())
            for key, var in var_map.items()
            if key in contract.defaults
        }
        self._regime_loaded_values["selected_scenario_packs"] = ",".join(selected_packs)
        self._regime_immutable_values = {key: str(var.get()) for key, var in self._regime_immutable_field_var_map().items()}

        status_map = {
            "exact_replay_compatible": "Exact Replay Compatible",
            "compatible_with_migration": "Compatible with Migration",
            "incompatible": "Incompatible",
        }
        repro_status = status_map.get(contract.reproducibility_status, "Compatible with Migration")
        self.regime_provenance_var.set(
            f"Reproducibility: {repro_status}\n"
            f"Regime: {contract.regime_name} ({contract.source})\n"
            f"Manifest: {contract.manifest_path}"
        )
        self._apply_regime_lock_state()
        self._refresh_regime_diff_indicator()

    def _regime_override_field_var_map(self) -> dict[str, tk.Variable]:
        return {
            "strategy": self.strategy_var,
            "lookback_days": self.lookback_days_var,
            "skip_days": self.skip_days_var,
            "portfolio_max_gross_exposure": self.portfolio_max_gross_var,
            "portfolio_min_net_exposure": self.portfolio_min_net_var,
            "portfolio_max_net_exposure": self.portfolio_max_net_var,
            "portfolio_max_symbol_weight": self.portfolio_max_symbol_var,
            "portfolio_max_sector_weight": self.portfolio_max_sector_var,
            "governance_min_stability_score": self.gov_min_stability_var,
            "governance_expected_signal_agreement": self.gov_expected_signal_agreement_var,
            "governance_max_signal_agreement_drift": self.gov_max_signal_agreement_drift_var,
            "stress_enable_historical_replay_regimes": self.stress_enable_historical_replay_var,
        }

    def _regime_immutable_field_var_map(self) -> dict[str, tk.Variable]:
        mapping: dict[str, tk.Variable] = {}
        for key, attr in (
            ("start_date", "start_date_var"),
            ("end_date", "end_date_var"),
            ("starting_capital", "starting_capital_var"),
            ("selected_trained_regime", "trained_regime_var"),
        ):
            var = getattr(self, attr, None)
            if var is not None:
                mapping[key] = var
        return mapping

    def _bind_regime_override_watchers(self) -> None:
        for var in self._regime_override_field_var_map().values():
            var.trace_add("write", lambda *_args: self._refresh_regime_diff_indicator())

    def _current_regime_field_snapshot(self) -> dict[str, str]:
        snapshot = {key: str(var.get()) for key, var in self._regime_override_field_var_map().items()}
        snapshot["selected_scenario_packs"] = ",".join(self._selected_listbox_values(self.scenario_pack_listbox))
        return snapshot

    def _refresh_regime_diff_indicator(self) -> None:
        if self._regime_loading_defaults:
            return
        if self._regime_locked_fields and self._active_regime_contract is not None:
            self._apply_regime_contract_lock_enforcement()
            return
        if not self._regime_loaded_values:
            self.regime_diff_var.set("")
            return
        current = self._current_regime_field_snapshot()
        changed = [key for key, expected in self._regime_loaded_values.items() if current.get(key, "") != expected]
        if not changed:
            self.regime_diff_var.set("No overrides yet.")
            return
        self.regime_diff_var.set("Overrides: " + ", ".join(changed))

    def _apply_regime_contract_lock_enforcement(self) -> None:
        if not self._regime_loaded_values:
            return
        self._regime_loading_defaults = True
        try:
            for key, expected in self._regime_loaded_values.items():
                if key == "selected_scenario_packs":
                    packs = [item.strip() for item in expected.split(",") if item.strip()]
                    self._set_listbox_selection(self.scenario_pack_listbox, packs, valid_options=self._scenario_pack_options)
                    continue
                var = self._regime_override_field_var_map().get(key)
                if var is not None and str(var.get()) != expected:
                    var.set(expected)
        finally:
            self._regime_loading_defaults = False
        self.regime_diff_var.set("Overrides disabled while lock is enabled.")

    def _on_regime_lock_toggled(self) -> None:
        self._regime_locked_fields = bool(self.regime_lock_var.get())
        self._apply_regime_lock_state()
        self._refresh_regime_diff_indicator()

    def _apply_regime_lock_state(self) -> None:
        enabled = bool(self._active_regime_contract)
        state = "disabled" if enabled and self._regime_locked_fields else "normal"
        for widget in (
            self.strategy_combo,
            self.lookback_days_entry,
            self.skip_days_entry,
            self.portfolio_max_gross_entry,
            self.portfolio_min_net_entry,
            self.portfolio_max_net_entry,
            self.portfolio_max_symbol_entry,
            self.portfolio_max_sector_entry,
            self.gov_min_stability_entry,
            self.gov_expected_signal_agreement_entry,
            self.gov_max_signal_agreement_drift_entry,
            self.scenario_pack_listbox,
        ):
            try:
                widget.configure(state=state)
            except tk.TclError:
                pass
        for widget in (
            self.trained_regime_combo,
        ):
            try:
                widget.configure(state="readonly")
            except tk.TclError:
                pass
        if not enabled:
            self.regime_lock_var.set(False)
            self._regime_locked_fields = False
            return
        if self.regime_lock_var.get() != self._regime_locked_fields:
            self.regime_lock_var.set(self._regime_locked_fields)

    def _migrate_backtest_settings(self, settings: dict[str, object]) -> tuple[dict[str, object], list[str], bool]:
        migrated = dict(settings)
        warnings: list[str] = []
        changed = False

        raw_schema = migrated.get("schema_version", 1)
        try:
            schema_version = int(raw_schema)
        except (TypeError, ValueError):
            warnings.append(
                f"Invalid key at $.backtest_settings.schema_version: expected integer, got {raw_schema!r}; treating as legacy schema 1."
            )
            schema_version = 1
            changed = True

        if schema_version < BACKTEST_SETTINGS_SCHEMA_VERSION:
            defaults_to_backfill = (
                "selected_preset",
                "selected_template",
                "selected_backtest_type",
                "selected_trained_regime",
                "selected_trained_regime_option_id",
                "selected_trained_regime_manifest_path",
                "selected_trained_regime_source",
                "ui_mode",
                "show_advanced_controls",
                "selected_stress_profile",
            )
            for key in defaults_to_backfill:
                if key not in migrated:
                    migrated[key] = DEFAULT_BACKTEST_SETTINGS[key]
                    warnings.append(
                        f"Missing key at $.backtest_settings.{key}; backfilled with default value {DEFAULT_BACKTEST_SETTINGS[key]!r}."
                    )
                    changed = True

            selected_preset = str(migrated.get("selected_preset", "custom"))
            if selected_preset not in {"custom", *BACKTEST_STRATEGY_PRESETS.keys()}:
                warnings.append(
                    f"Invalid key at $.backtest_settings.selected_preset: unsupported value {selected_preset!r}; using 'custom'."
                )
                migrated["selected_preset"] = "custom"
                changed = True
            selected_backtest_type = str(migrated.get("selected_backtest_type", "classic_strategy")).strip().lower().replace(" ", "_")
            if selected_backtest_type not in set(BACKTEST_WORKFLOW_TYPES):
                warnings.append(
                    f"Invalid key at $.backtest_settings.selected_backtest_type: unsupported value {selected_backtest_type!r}; using 'classic_strategy'."
                )
                migrated["selected_backtest_type"] = "classic_strategy"
                changed = True

        if migrated.get("schema_version") != BACKTEST_SETTINGS_SCHEMA_VERSION:
            migrated["schema_version"] = BACKTEST_SETTINGS_SCHEMA_VERSION
            changed = True

        return migrated, warnings, changed

    def _apply_settings(self, settings: dict[str, object]) -> None:
        merged = dict(DEFAULT_BACKTEST_SETTINGS)
        merged.update(settings)
        self.controller.state.backtest_settings = merged
        self.refresh()

    def _on_preset_selected(self, _event: object | None = None) -> None:
        preset_display = self.preset_var.get().strip()
        preset_key = self._preset_display_to_key.get(preset_display, "custom")
        if preset_key == "custom":
            return
        preset = BACKTEST_STRATEGY_PRESETS.get(preset_key)
        if not preset:
            return
        preset_settings = preset.get("settings", {})
        if isinstance(preset_settings, dict):
            self._apply_settings(preset_settings)
            self.preset_var.set(self._preset_key_to_display.get(preset_key, "Custom"))
            self.controller.state.backtest_settings["selected_preset"] = preset_key
            self.controller.persist_state()

    def _on_test_suite_selected(self, _event: object | None = None) -> None:
        suite_display = self.test_suite_var.get().strip()
        suite_key = self._suite_display_to_key.get(suite_display, "custom")
        if suite_key == "custom":
            return
        suite = BACKTEST_TEST_SUITE_PRESETS.get(suite_key)
        if not suite:
            return
        suite_settings = suite.get("settings", {})
        if isinstance(suite_settings, dict):
            self._apply_settings(suite_settings)
            self.test_suite_var.set(self._suite_key_to_display.get(suite_key, "Custom"))
            self.controller.state.backtest_settings["selected_test_suite"] = suite_key
            self.controller.persist_state()

    def save_template(self) -> None:
        template_name = self.template_var.get().strip()
        if not template_name:
            messagebox.showinfo("Template name required", "Choose a template name in the Experiment Template field.")
            return
        if not self.save_settings(show_confirmation=False):
            return
        snapshot = dict(self.controller.state.backtest_settings)
        self.controller.state.backtest_templates[template_name] = snapshot
        self.controller.state.backtest_settings["selected_template"] = template_name
        self.controller.persist_state()
        self._refresh_template_choices()
        self._refresh_trained_regime_choices()
        messagebox.showinfo("Saved", f"Template '{template_name}' saved.")

    def load_template(self) -> None:
        template_name = self.template_var.get().strip()
        template = self.controller.state.backtest_templates.get(template_name)
        if template is None:
            messagebox.showinfo("Template missing", "Select a saved template first.")
            return
        self._apply_settings(template)
        self.controller.state.backtest_settings["selected_template"] = template_name
        self.controller.persist_state()
        messagebox.showinfo("Loaded", f"Template '{template_name}' loaded.")

    def refresh(self) -> None:
        self._refresh_trained_regime_choices()
        if not self.trained_regime_var.get().strip():
            self.trained_regime_var.set("(none)")
        settings = dict(DEFAULT_BACKTEST_SETTINGS)
        settings.update(self.controller.state.backtest_settings)
        settings, migration_messages, migration_changed = self._migrate_backtest_settings(settings)
        if migration_changed:
            self.controller.state.backtest_settings = dict(settings)
            self.controller.persist_state()

        strategy = str(settings.get("strategy", "momentum"))
        if strategy not in STRATEGIES:
            strategy = "momentum"
        self.strategy_var.set(strategy)
        self.ui_mode_var.set(str(settings.get("ui_mode", "basic")))
        self.show_advanced_controls_var.set(bool(settings.get("show_advanced_controls", self.ui_mode_var.get() == "advanced")))
        self.ui_mode_var.set("advanced" if self.show_advanced_controls_var.get() else "basic")
        selected_preset = str(settings.get("selected_preset", "custom"))
        if selected_preset not in {"custom", *BACKTEST_STRATEGY_PRESETS.keys()}:
            selected_preset = "custom"
        self.preset_var.set(self._preset_key_to_display.get(selected_preset, "Custom"))
        selected_test_suite = str(settings.get("selected_test_suite", "custom"))
        if selected_test_suite not in {"custom", *BACKTEST_TEST_SUITE_PRESETS.keys()}:
            selected_test_suite = "custom"
        self.test_suite_var.set(self._suite_key_to_display.get(selected_test_suite, "Custom"))

        selected_backtest_type = str(settings.get("selected_backtest_type", "classic_strategy")).strip().lower().replace(" ", "_")
        if selected_backtest_type not in set(BACKTEST_WORKFLOW_TYPES):
            selected_backtest_type = "classic_strategy"
        self.backtest_type_var.set(BACKTEST_WORKFLOW_TYPE_LABELS.get(selected_backtest_type, "Classic Strategy"))

        selected_trained_regime = str(settings.get("selected_trained_regime", "(none)"))
        regime_labels = set(self.trained_regime_combo.cget("values"))
        selected_option_id = str(settings.get("selected_trained_regime_option_id", "")).strip()
        selected_manifest = str(settings.get("selected_trained_regime_manifest_path", "")).strip()
        selected_source = str(settings.get("selected_trained_regime_source", "")).strip()
        if selected_trained_regime in regime_labels:
            self.trained_regime_var.set(selected_trained_regime)
        else:
            recovered = "(none)"
            for option in self._regime_backtest_options:
                if selected_option_id and option.option_id == selected_option_id:
                    recovered = option.label
                    break
                if selected_manifest and option.manifest_path == selected_manifest and (not selected_source or option.source == selected_source):
                    recovered = option.label
                    break
            self.trained_regime_var.set(recovered)

        self.lookback_days_var.set(str(settings.get("lookback_days", "90")))
        self.skip_days_var.set(str(settings.get("skip_days", "5")))
        self.costs_bps_var.set(str(settings.get("costs_bps", "5")))
        execution_model_value = normalize_supported_option(str(settings.get("execution_model", "bps")), EXECUTION_MODELS) or "bps"
        self.execution_model_var.set(execution_model_value)
        self.execution_spread_bps_var.set(str(settings.get("execution_spread_bps", "2")))
        self.execution_max_participation_var.set(str(settings.get("execution_max_participation", "1.0")))
        self.execution_impact_bps_var.set(str(settings.get("execution_impact_bps", "5")))
        self.execution_latency_bars_var.set(str(settings.get("execution_latency_bars", "0")))
        self.execution_latency_ms_var.set(str(settings.get("execution_latency_ms", "0")))
        self.stress_enable_historical_replay_var.set(bool(settings.get("stress_enable_historical_replay_regimes", True)))
        self.stress_historical_window_fraction_var.set(str(settings.get("stress_historical_window_fraction", "0.20")))
        self.stress_historical_replay_window_bars_var.set(str(settings.get("stress_historical_replay_window_bars", "20")))
        self.stress_synthetic_jump_magnitude_var.set(str(settings.get("stress_synthetic_jump_magnitude", "0.02")))
        self.stress_synthetic_jump_interval_var.set(str(settings.get("stress_synthetic_jump_interval", "7")))
        self.stress_synthetic_vol_cluster_multiplier_var.set(str(settings.get("stress_synthetic_vol_cluster_multiplier", "1.6")))
        self.stress_overlay_spread_multiplier_var.set(str(settings.get("stress_overlay_spread_multiplier", "2.5")))
        self.stress_overlay_liquidity_multiplier_var.set(str(settings.get("stress_overlay_liquidity_multiplier", "0.4")))
        selected_packs = [item.strip() for item in str(settings.get("selected_scenario_packs", "")).split(",") if item.strip()]
        self._set_listbox_selection(self.scenario_pack_listbox, selected_packs, valid_options=self._scenario_pack_options)
        saved_profile = str(settings.get("selected_stress_profile", "Base"))
        self.selected_stress_profile_var.set(saved_profile if saved_profile in STRESS_PROFILES else "Base")
        self.starting_capital_var.set(str(settings.get("starting_capital", "100000")))
        self.bet_sizing_mode_var.set(str(settings.get("bet_sizing_mode", "half_kelly")))
        self.custom_bet_pct_var.set(str(settings.get("custom_bet_pct", "10")))
        timeframe = str(settings.get("timeframe", "1m"))
        self.timeframe_var.set(timeframe if timeframe in TIMEFRAMES else "1m")
        self.use_walk_forward_var.set(bool(settings.get("use_walk_forward", False)))
        self.use_optimizer_var.set(bool(settings.get("use_optimizer", False)))
        optimizer_sampler = normalize_supported_option(
            str(settings.get("optimizer_sampler", "tpe")),
            OPTIMIZER_SAMPLERS,
            field_name="optimizer sampler",
        ) or "tpe"
        self.optimizer_sampler_var.set(optimizer_sampler)
        self.optimizer_trials_var.set(str(settings.get("optimizer_n_trials", "20")))
        self.optimizer_enable_pruning_var.set(bool(settings.get("optimizer_enable_pruning", True)))
        self.optimizer_prune_constraint_var.set(bool(settings.get("optimizer_prune_constraint", True)))
        self.optimizer_prune_lcb_var.set(bool(settings.get("optimizer_prune_lcb", True)))
        self.optimizer_min_completed_var.set(str(settings.get("optimizer_min_completed_for_pruning", "5")))
        self.optimizer_staged_budgets_var.set(str(settings.get("optimizer_staged_budgets", self.optimizer_staged_budgets_var.get())))
        self.optimizer_search_space_var.set(str(settings.get("optimizer_search_space", DEFAULT_OPTIMIZER_SEARCH_SPACE)))
        self.optimizer_objectives_var.set(str(settings.get("optimizer_objectives", DEFAULT_OPTIMIZER_OBJECTIVES)))
        self.optimizer_max_turnover_var.set(str(settings.get("optimizer_max_turnover", "")))
        self.optimizer_max_drawdown_floor_var.set(str(settings.get("optimizer_max_drawdown_floor", "")))
        self.optimizer_min_trades_var.set(str(settings.get("optimizer_min_trades", "")))
        self.optimizer_objective_weights_var.set(str(settings.get("optimizer_objective_weights", DEFAULT_OPTIMIZER_OBJECTIVE_WEIGHTS)))
        self.optimizer_overfitting_penalty_var.set(str(settings.get("optimizer_overfitting_penalty", DEFAULT_OPTIMIZER_OVERFITTING_PENALTY)))
        self.portfolio_method_var.set(str(settings.get("portfolio_method", "equal_weight")))
        self.portfolio_vol_lookback_var.set(str(settings.get("portfolio_vol_lookback_bars", "20")))
        self.portfolio_target_vol_var.set(str(settings.get("portfolio_target_volatility", "0.10")))
        self.portfolio_max_symbol_var.set(str(settings.get("portfolio_max_symbol_weight", "0.25")))
        self.portfolio_max_sector_var.set(str(settings.get("portfolio_max_sector_weight", "0.60")))
        self.portfolio_rebalance_frequency_var.set(str(settings.get("portfolio_rebalance_frequency_bars", "1")))
        self.portfolio_clustering_linkage_var.set(str(settings.get("portfolio_clustering_linkage", "single")))
        self.portfolio_covariance_shrinkage_var.set(str(settings.get("portfolio_covariance_shrinkage", "0.15")))
        self.portfolio_max_gross_var.set(str(settings.get("portfolio_max_gross_exposure", "1.0")))
        self.portfolio_min_net_var.set(str(settings.get("portfolio_min_net_exposure", "-1.0")))
        self.portfolio_max_net_var.set(str(settings.get("portfolio_max_net_exposure", "1.0")))
        self.portfolio_max_net_gamma_var.set(str(settings.get("portfolio_max_net_gamma", "")))
        self.portfolio_max_abs_vega_bucket_var.set(str(settings.get("portfolio_max_abs_vega_bucket", "")))
        self.portfolio_max_abs_delta_underlying_var.set(str(settings.get("portfolio_max_abs_delta_per_underlying", "")))
        self.portfolio_max_participation_rate_var.set(str(settings.get("max_participation_rate", "")))
        for risk_key, variable in self._risk_limit_vars.items():
            parsed_value = parse_float(variable.get())
            if parsed_value is None:
                continue
            low, high, _ = self._risk_limit_ranges[risk_key]
            self._risk_limit_slider_vars[risk_key].set(max(low, min(high, float(parsed_value))))
        self.wf_train_fraction_var.set(float(settings.get("wf_train_fraction", "0.70")))
        self.wf_validation_fraction_var.set(float(settings.get("wf_validation_fraction", "0.15")))
        self.wf_test_fraction_var.set(float(settings.get("wf_test_fraction", "0.15")))
        self.wf_step_fraction_var.set(float(settings.get("wf_step_fraction", "0.15")))
        self.wf_train_bars_var.set(str(settings.get("wf_train_bars", "")))
        self.wf_validation_bars_var.set(str(settings.get("wf_validation_bars", "")))
        self.wf_test_bars_var.set(str(settings.get("wf_test_bars", "")))
        self.wf_step_bars_var.set(str(settings.get("wf_step_bars", "")))
        self.wf_cv_scheme_var.set(str(settings.get("wf_cv_scheme", "walk_forward")))
        self.wf_purge_bars_var.set(str(settings.get("wf_purge_window_bars", "0")))
        self.wf_embargo_bars_var.set(str(settings.get("wf_embargo_window_bars", "0")))
        self.wf_cpcv_groups_var.set(str(settings.get("wf_cpcv_n_groups", "6")))
        self.wf_cpcv_test_groups_var.set(str(settings.get("wf_cpcv_n_test_groups", "2")))
        self.wf_cv_seed_var.set(str(settings.get("wf_cv_seed", "42")))
        self.wf_label_horizon_bars_var.set(str(settings.get("wf_label_horizon_bars", "1")))
        self.wf_nested_optimization_var.set(bool(settings.get("wf_nested_optimization", False)))
        self.wf_inner_train_fraction_var.set(str(settings.get("wf_inner_train_fraction", "0.70")))
        self.wf_objective_weights_var.set(str(settings.get("wf_objective_weights", "")))
        self.wf_overfitting_penalty_var.set(str(settings.get("wf_overfitting_penalty", "")))
        self.wf_strategy_key_var.set(str(settings.get("wf_strategy_key", "")))
        self.wf_prior_strategy_keys_var.set(str(settings.get("wf_prior_strategy_keys", "")))
        self._refresh_wf_fraction_labels()

        selected_entries, stale_entries, migrated_entries = self._resolve_supported_csv_setting(
            settings.get("selected_entry_signals", "ts_momentum"),
            supported=ENTRY_SIGNALS,
            fallback=("ts_momentum",),
            field_name="entry signals",
        )
        selected_exits, stale_exits, migrated_exits = self._resolve_supported_csv_setting(
            settings.get("selected_exit_signals", "none"),
            supported=EXIT_SIGNALS,
            fallback=("none",),
            field_name="exit signals",
        )
        for name, var in self.entry_signal_vars.items():
            var.set(name in selected_entries)
        for name, var in self.exit_signal_vars.items():
            var.set(name in selected_exits)

        stale_messages = [
            migration_hint_text(
                stale=stale_entries,
                migrations=migrated_entries,
                supported=ENTRY_SIGNALS,
                field_name="entry signal preset values",
            ),
            migration_hint_text(
                stale=stale_exits,
                migrations=migrated_exits,
                supported=EXIT_SIGNALS,
                field_name="exit signal preset values",
            ),
        ]
        if normalize_supported_option(str(settings.get("execution_model", "bps")), EXECUTION_MODELS) is None:
            stale_messages.append(
                f"Unsupported execution model preset value: {settings.get('execution_model')} | Supported values: {', '.join(EXECUTION_MODELS)}"
            )
        if normalize_supported_option(
            str(settings.get("optimizer_sampler", "tpe")),
            OPTIMIZER_SAMPLERS,
            field_name="optimizer sampler",
        ) is None:
            stale_messages.append(
                f"Unsupported optimizer sampler preset value: {settings.get('optimizer_sampler')} | Supported values: {', '.join(OPTIMIZER_SAMPLERS)}"
            )
        stale_messages.extend(migration_messages)
        self._stale_preset_messages = [item for item in stale_messages if item]

        self.xsmom_top_quantile_var.set(str(settings.get("xsmom_top_quantile", "0.2")))
        self.xsmom_bottom_quantile_var.set(str(settings.get("xsmom_bottom_quantile", "0.2")))
        self.xsmom_vol_lookback_days_var.set(str(settings.get("xsmom_vol_lookback_days", "20")))
        self.xsmom_long_only_var.set(bool(settings.get("xsmom_long_only", False)))

        self.start_date_var.set(str(settings.get("start_date", "")))
        self.end_date_var.set(str(settings.get("end_date", "")))
        self.backtest_root_var.set(str(settings.get("backtest_data_root", str(BACKTEST_CACHE_DIR))))
        self.template_var.set(str(settings.get("selected_template", "")))

        self.notes_text.delete("1.0", tk.END)
        self.notes_text.insert("1.0", str(settings.get("notes", "")))
        self.gov_hypothesis_id_var.set(str(settings.get("governance_hypothesis_id", "")))
        self.gov_experiment_id_var.set(str(settings.get("governance_experiment_id", "")))
        self.gov_owner_var.set(str(settings.get("governance_owner", "")))
        self.gov_dataset_lock_var.set(str(settings.get("governance_dataset_snapshot_lock", "")))
        self.gov_acceptance_text.delete("1.0", tk.END)
        self.gov_acceptance_text.insert("1.0", str(settings.get("governance_acceptance_criteria", "")))
        self.gov_promotion_state_var.set(str(settings.get("governance_promotion_state", "research")))
        self.gov_approval_status_var.set(str(settings.get("governance_approval_status", "pending")))
        self.gov_min_oos_periods_var.set(str(settings.get("governance_min_oos_periods", "3")))
        self.gov_min_stability_var.set(str(settings.get("governance_min_stability_score", "0.55")))
        self.gov_max_turnover_var.set(str(settings.get("governance_max_turnover_total", "4.0")))
        self.gov_min_capacity_var.set(str(settings.get("governance_min_capacity_score", "0.5")))
        self.gov_max_signal_agreement_drift_var.set(str(settings.get("governance_max_signal_agreement_drift", "0.10")))
        self.gov_max_fill_slippage_drift_bps_var.set(str(settings.get("governance_max_fill_slippage_drift_bps", "5.0")))
        self.gov_max_pnl_attribution_divergence_var.set(str(settings.get("governance_max_pnl_attribution_divergence", "0.15")))
        self.gov_expected_signal_agreement_var.set(str(settings.get("governance_expected_signal_agreement", "1.0")))
        self.gov_expected_fill_slippage_bps_var.set(str(settings.get("governance_expected_fill_slippage_bps", "0.0")))
        self.gov_expected_pnl_attribution_var.set(str(settings.get("governance_expected_pnl_attribution", "1.0")))
        self.gov_observed_signal_agreement_var.set(str(settings.get("governance_observed_signal_agreement", "1.0")))
        self.gov_observed_fill_slippage_bps_var.set(str(settings.get("governance_observed_fill_slippage_bps", "0.0")))
        self.gov_observed_pnl_attribution_var.set(str(settings.get("governance_observed_pnl_attribution", "1.0")))
        self._governance_comments = [
            dict(item) for item in settings.get("governance_comments", []) if isinstance(item, dict)
        ]
        self._governance_review_actions = [
            dict(item) for item in settings.get("governance_review_actions", []) if isinstance(item, dict)
        ]
        self._governance_decision_log = [
            dict(item) for item in settings.get("governance_decision_log", []) if isinstance(item, dict)
        ]
        self._render_governance_log()
        self.logs_text.delete("1.0", tk.END)
        self.logs_text.insert("1.0", str(settings.get("notes", "")))
        self._refresh_template_choices()
        self._refresh_trained_regime_choices()
        self._on_strategy_changed()
        self._on_mode_changed()
        self._update_validation_hint()

    def save_settings(self, show_confirmation: bool = True) -> bool:
        validated = self._validate_common_inputs()
        if validated is None:
            return False

        strategy, lookback, skip, costs_bps, starting_capital, custom_bet_pct = validated

        selected_entries: list[str] = []
        selected_exits: list[str] = []
        xsmom_params = {
            "xsmom_top_quantile": self.xsmom_top_quantile_var.get().strip() or "0.2",
            "xsmom_bottom_quantile": self.xsmom_bottom_quantile_var.get().strip() or "0.2",
            "xsmom_vol_lookback_days": self.xsmom_vol_lookback_days_var.get().strip() or "20",
            "xsmom_long_only": bool(self.xsmom_long_only_var.get()),
        }

        if strategy == "momentum":
            selected_entries = self._selected_signal_names(self.entry_signal_vars)
            selected_exits = self._selected_signal_names(self.exit_signal_vars)
            if not selected_entries:
                messagebox.showinfo("Invalid input", "Select at least one entry signal.")
                return False
            if not selected_exits:
                messagebox.showinfo("Invalid input", "Select at least one exit signal.")
                return False
        else:
            xsmom_valid = self._validate_xsmom_inputs()
            if xsmom_valid is None:
                return False
            xsmom_top_quantile, xsmom_bottom_quantile, xsmom_vol_lookback_days = xsmom_valid
            xsmom_params = {
                "xsmom_top_quantile": str(xsmom_top_quantile),
                "xsmom_bottom_quantile": str(xsmom_bottom_quantile),
                "xsmom_vol_lookback_days": str(xsmom_vol_lookback_days),
                "xsmom_long_only": bool(self.xsmom_long_only_var.get()),
            }

        selected_regime_option = self._selected_regime_option()
        self.controller.state.backtest_settings = {
            "strategy": strategy,
            "strategy_name": "Cross-Sectional Momentum" if strategy == "xsmom" else "Time-Series Momentum",
            "lookback_days": str(int(lookback)),
            "skip_days": str(int(skip)),
            "costs_bps": str(costs_bps),
            "execution_model": self.execution_model_var.get().strip() or "bps",
            "execution_spread_bps": self.execution_spread_bps_var.get().strip() or "2",
            "execution_max_participation": self.execution_max_participation_var.get().strip() or "1.0",
            "execution_impact_bps": self.execution_impact_bps_var.get().strip() or "5",
            "execution_latency_bars": self.execution_latency_bars_var.get().strip() or "0",
            "execution_latency_ms": self.execution_latency_ms_var.get().strip() or "0",
            "stress_enable_historical_replay_regimes": bool(self.stress_enable_historical_replay_var.get()),
            "stress_historical_window_fraction": self.stress_historical_window_fraction_var.get().strip() or "0.20",
            "stress_historical_replay_window_bars": self.stress_historical_replay_window_bars_var.get().strip() or "20",
            "stress_synthetic_jump_magnitude": self.stress_synthetic_jump_magnitude_var.get().strip() or "0.02",
            "stress_synthetic_jump_interval": self.stress_synthetic_jump_interval_var.get().strip() or "7",
            "stress_synthetic_vol_cluster_multiplier": self.stress_synthetic_vol_cluster_multiplier_var.get().strip() or "1.6",
            "stress_overlay_spread_multiplier": self.stress_overlay_spread_multiplier_var.get().strip() or "2.5",
            "stress_overlay_liquidity_multiplier": self.stress_overlay_liquidity_multiplier_var.get().strip() or "0.4",
            "selected_scenario_packs": ",".join(self._selected_listbox_values(self.scenario_pack_listbox)),
            "selected_stress_profile": self.selected_stress_profile_var.get().strip() or "Base",
            "starting_capital": str(starting_capital),
            "bet_sizing_mode": self.bet_sizing_mode_var.get().strip() or "half_kelly",
            "custom_bet_pct": str(custom_bet_pct),
            "timeframe": self.timeframe_var.get().strip() or "1m",
            "use_walk_forward": bool(self.use_walk_forward_var.get()),
            "use_optimizer": bool(self.use_optimizer_var.get()),
            "optimizer_sampler": self.optimizer_sampler_var.get().strip() or "tpe",
            "optimizer_n_trials": self.optimizer_trials_var.get().strip() or "20",
            "optimizer_enable_pruning": bool(self.optimizer_enable_pruning_var.get()),
            "optimizer_prune_constraint": bool(self.optimizer_prune_constraint_var.get()),
            "optimizer_prune_lcb": bool(self.optimizer_prune_lcb_var.get()),
            "optimizer_min_completed_for_pruning": self.optimizer_min_completed_var.get().strip() or "5",
            "optimizer_staged_budgets": self.optimizer_staged_budgets_var.get().strip(),
            "optimizer_search_space": self.optimizer_search_space_var.get().strip(),
            "optimizer_objectives": self.optimizer_objectives_var.get().strip(),
            "optimizer_max_turnover": self.optimizer_max_turnover_var.get().strip(),
            "optimizer_max_drawdown_floor": self.optimizer_max_drawdown_floor_var.get().strip(),
            "optimizer_min_trades": self.optimizer_min_trades_var.get().strip(),
            "optimizer_objective_weights": self.optimizer_objective_weights_var.get().strip(),
            "optimizer_overfitting_penalty": self.optimizer_overfitting_penalty_var.get().strip(),
            "wf_train_fraction": f"{float(self.wf_train_fraction_var.get()):.2f}",
            "wf_validation_fraction": f"{float(self.wf_validation_fraction_var.get()):.2f}",
            "wf_test_fraction": f"{float(self.wf_test_fraction_var.get()):.2f}",
            "wf_step_fraction": f"{float(self.wf_step_fraction_var.get()):.2f}",
            "wf_train_bars": self.wf_train_bars_var.get().strip(),
            "wf_validation_bars": self.wf_validation_bars_var.get().strip(),
            "wf_test_bars": self.wf_test_bars_var.get().strip(),
            "wf_step_bars": self.wf_step_bars_var.get().strip(),
            "wf_cv_scheme": self.wf_cv_scheme_var.get().strip() or "walk_forward",
            "wf_label_horizon_bars": self.wf_label_horizon_bars_var.get().strip() or "1",
            "wf_nested_optimization": bool(self.wf_nested_optimization_var.get()),
            "wf_inner_train_fraction": self.wf_inner_train_fraction_var.get().strip() or "0.70",
            "wf_objective_weights": self.wf_objective_weights_var.get().strip(),
            "wf_overfitting_penalty": self.wf_overfitting_penalty_var.get().strip(),
            "wf_strategy_key": self.wf_strategy_key_var.get().strip(),
            "wf_prior_strategy_keys": self.wf_prior_strategy_keys_var.get().strip(),
            "wf_purge_window_bars": self.wf_purge_bars_var.get().strip() or "0",
            "wf_embargo_window_bars": self.wf_embargo_bars_var.get().strip() or "0",
            "wf_cpcv_n_groups": self.wf_cpcv_groups_var.get().strip() or "6",
            "wf_cpcv_n_test_groups": self.wf_cpcv_test_groups_var.get().strip() or "2",
            "wf_cv_seed": self.wf_cv_seed_var.get().strip() or "42",
            "portfolio_method": self.portfolio_method_var.get().strip() or "equal_weight",
            "portfolio_vol_lookback_bars": self.portfolio_vol_lookback_var.get().strip() or "20",
            "portfolio_target_volatility": self.portfolio_target_vol_var.get().strip() or "0.10",
            "portfolio_max_symbol_weight": self.portfolio_max_symbol_var.get().strip() or "0.25",
            "portfolio_max_sector_weight": self.portfolio_max_sector_var.get().strip() or "0.60",
            "portfolio_rebalance_frequency_bars": self.portfolio_rebalance_frequency_var.get().strip() or "1",
            "portfolio_clustering_linkage": self.portfolio_clustering_linkage_var.get().strip() or "single",
            "portfolio_covariance_shrinkage": self.portfolio_covariance_shrinkage_var.get().strip() or "0.15",
            "portfolio_max_gross_exposure": self.portfolio_max_gross_var.get().strip() or "1.0",
            "portfolio_min_net_exposure": self.portfolio_min_net_var.get().strip() or "-1.0",
            "portfolio_max_net_exposure": self.portfolio_max_net_var.get().strip() or "1.0",
            "portfolio_max_net_gamma": self.portfolio_max_net_gamma_var.get().strip(),
            "portfolio_max_abs_vega_bucket": self.portfolio_max_abs_vega_bucket_var.get().strip(),
            "portfolio_max_abs_delta_per_underlying": self.portfolio_max_abs_delta_underlying_var.get().strip(),
            "max_participation_rate": self.portfolio_max_participation_rate_var.get().strip(),
            "selected_entry_signals": ",".join(selected_entries),
            "selected_exit_signals": ",".join(selected_exits),
            "start_date": self.start_date_var.get().strip(),
            "end_date": self.end_date_var.get().strip(),
            "backtest_data_root": self.backtest_root_var.get().strip(),
            "notes": self.notes_text.get("1.0", tk.END).strip(),
            "governance_hypothesis_id": self.gov_hypothesis_id_var.get().strip(),
            "governance_experiment_id": self.gov_experiment_id_var.get().strip(),
            "governance_owner": self.gov_owner_var.get().strip(),
            "governance_dataset_snapshot_lock": self.gov_dataset_lock_var.get().strip(),
            "governance_acceptance_criteria": self.gov_acceptance_text.get("1.0", tk.END).strip(),
            "governance_promotion_state": self.gov_promotion_state_var.get().strip() or "research",
            "governance_approval_status": self.gov_approval_status_var.get().strip() or "pending",
            "governance_min_oos_periods": self.gov_min_oos_periods_var.get().strip() or "3",
            "governance_min_stability_score": self.gov_min_stability_var.get().strip() or "0.55",
            "governance_max_turnover_total": self.gov_max_turnover_var.get().strip() or "4.0",
            "governance_min_capacity_score": self.gov_min_capacity_var.get().strip() or "0.5",
            "governance_max_signal_agreement_drift": self.gov_max_signal_agreement_drift_var.get().strip() or "0.10",
            "governance_max_fill_slippage_drift_bps": self.gov_max_fill_slippage_drift_bps_var.get().strip() or "5.0",
            "governance_max_pnl_attribution_divergence": self.gov_max_pnl_attribution_divergence_var.get().strip() or "0.15",
            "governance_expected_signal_agreement": self.gov_expected_signal_agreement_var.get().strip() or "1.0",
            "governance_expected_fill_slippage_bps": self.gov_expected_fill_slippage_bps_var.get().strip() or "0.0",
            "governance_expected_pnl_attribution": self.gov_expected_pnl_attribution_var.get().strip() or "1.0",
            "governance_observed_signal_agreement": self.gov_observed_signal_agreement_var.get().strip() or "1.0",
            "governance_observed_fill_slippage_bps": self.gov_observed_fill_slippage_bps_var.get().strip() or "0.0",
            "governance_observed_pnl_attribution": self.gov_observed_pnl_attribution_var.get().strip() or "1.0",
            "governance_comments": list(self._governance_comments),
            "governance_review_actions": list(self._governance_review_actions),
            "governance_decision_log": list(self._governance_decision_log),
            "ui_mode": self.ui_mode_var.get().strip() or "basic",
            "show_advanced_controls": bool(self.show_advanced_controls_var.get()),
            "selected_preset": self._preset_display_to_key.get(self.preset_var.get().strip(), "custom"),
            "selected_test_suite": self._suite_display_to_key.get(self.test_suite_var.get().strip(), "custom"),
            "selected_template": self.template_var.get().strip(),
            "selected_backtest_type": self._normalized_backtest_type(),
            "selected_trained_regime": self.trained_regime_var.get().strip(),
            "selected_trained_regime_option_id": (selected_regime_option.option_id if selected_regime_option else ""),
            "selected_trained_regime_manifest_path": (selected_regime_option.manifest_path if selected_regime_option else ""),
            "selected_trained_regime_source": (selected_regime_option.source if selected_regime_option else ""),
            "schema_version": BACKTEST_SETTINGS_SCHEMA_VERSION,
            **xsmom_params,
        }
        self.controller.persist_state()
        self._refresh_template_choices()
        self._refresh_trained_regime_choices()
        if show_confirmation:
            messagebox.showinfo("Saved", "Backtesting parameters saved.")
        return True

    def run_stress_only(self) -> None:
        self.run_backtest(run_mode="stress_only")

    def run_full_chain(self) -> None:
        self.run_backtest(run_mode="full_chain")

    def run_backtest(self, run_mode: str = "full_chain") -> None:
        if self._update_validation_hint():
            messagebox.showinfo("Validation warning", "Resolve validation hints before running the backtest.")
            return

        tickers = list(self.controller.state.tickers)
        if not tickers:
            messagebox.showinfo("No tickers", "Add tickers before running a backtest.")
            return

        validated = self._validate_common_inputs()
        if validated is None:
            return
        strategy, lookback, skip, costs_bps, starting_capital, custom_bet_pct = validated

        start_date = parse_date(self.start_date_var.get())
        end_date = parse_date(self.end_date_var.get())
        if start_date is None or end_date is None:
            messagebox.showinfo("Invalid dates", "Both start and end dates must be valid YYYY-MM-DD values.")
            return
        if start_date >= end_date:
            messagebox.showinfo("Invalid dates", "Start date must be before end date.")
            return

        cache_root = normalize_cache_root(self.backtest_root_var.get())
        bet_sizing_mode = self.bet_sizing_mode_var.get().strip() or "half_kelly"
        timeframe = self.timeframe_var.get().strip() or "1m"
        if timeframe not in TIMEFRAMES:
            messagebox.showinfo("Invalid input", "Please select a valid resolution.")
            return

        selected_backtest_type = self._normalized_backtest_type()
        if selected_backtest_type not in set(BACKTEST_WORKFLOW_TYPES):
            messagebox.showinfo("Invalid input", "Please select a valid backtest type.")
            return

        parsed_vol_lookback = parse_float(self.portfolio_vol_lookback_var.get())
        parsed_target_vol = parse_float(self.portfolio_target_vol_var.get())
        parsed_max_symbol = parse_float(self.portfolio_max_symbol_var.get())
        parsed_max_sector = parse_float(self.portfolio_max_sector_var.get())
        parsed_rebalance_frequency = parse_float(self.portfolio_rebalance_frequency_var.get())
        parsed_cov_shrinkage = parse_float(self.portfolio_covariance_shrinkage_var.get())
        parsed_max_gross = parse_float(self.portfolio_max_gross_var.get())
        parsed_min_net = parse_float(self.portfolio_min_net_var.get())
        parsed_max_net = parse_float(self.portfolio_max_net_var.get())
        parsed_max_gamma = parse_float(self.portfolio_max_net_gamma_var.get())
        parsed_max_vega_bucket = parse_float(self.portfolio_max_abs_vega_bucket_var.get())
        parsed_max_delta_underlying = parse_float(self.portfolio_max_abs_delta_underlying_var.get())
        parsed_max_participation_rate = parse_float(self.portfolio_max_participation_rate_var.get())

        portfolio_cfg = {
            "portfolio_method": self.portfolio_method_var.get().strip() or "equal_weight",
            "portfolio_vol_lookback_bars": int(parsed_vol_lookback) if parsed_vol_lookback is not None else 20,
            "portfolio_target_volatility": float(parsed_target_vol) if parsed_target_vol is not None else 0.10,
            "portfolio_max_symbol_weight": float(parsed_max_symbol) if parsed_max_symbol is not None else 0.25,
            "portfolio_max_sector_weight": float(parsed_max_sector) if parsed_max_sector is not None else 0.60,
            "portfolio_rebalance_frequency_bars": int(parsed_rebalance_frequency) if parsed_rebalance_frequency is not None else 1,
            "portfolio_clustering_linkage": self.portfolio_clustering_linkage_var.get().strip() or "single",
            "portfolio_covariance_shrinkage": float(parsed_cov_shrinkage) if parsed_cov_shrinkage is not None else 0.15,
            "portfolio_max_gross_exposure": float(parsed_max_gross) if parsed_max_gross is not None else 1.0,
            "portfolio_min_net_exposure": float(parsed_min_net) if parsed_min_net is not None else -1.0,
            "portfolio_max_net_exposure": float(parsed_max_net) if parsed_max_net is not None else 1.0,
            "portfolio_max_net_gamma": float(parsed_max_gamma) if parsed_max_gamma is not None else None,
            "portfolio_max_abs_vega_bucket": float(parsed_max_vega_bucket) if parsed_max_vega_bucket is not None else None,
            "portfolio_max_abs_delta_per_underlying": float(parsed_max_delta_underlying) if parsed_max_delta_underlying is not None else None,
            "max_participation_rate": float(parsed_max_participation_rate) if parsed_max_participation_rate is not None else None,
        }
        if portfolio_cfg["portfolio_method"] not in PORTFOLIO_METHODS:
            messagebox.showinfo("Invalid input", "Please select a valid portfolio method.")
            return

        execution_model = self.execution_model_var.get().strip() or "bps"
        if execution_model not in EXECUTION_MODELS:
            messagebox.showinfo("Invalid input", f"Execution model must be one of: {', '.join(EXECUTION_MODELS)}.")
            return
        execution_model_params = {
            "spread_bps": float(parse_float(self.execution_spread_bps_var.get()) or 2.0),
            "max_participation": float(parse_float(self.execution_max_participation_var.get()) or 1.0),
            "impact_bps": float(parse_float(self.execution_impact_bps_var.get()) or costs_bps),
            "latency_bars": int(parse_float(self.execution_latency_bars_var.get()) or 0),
            "latency_ms": int(parse_float(self.execution_latency_ms_var.get()) or 0),
            "drift_bps_per_bar": float(parse_float(self.execution_impact_bps_var.get()) or 1.0),
        }
        selected_suite_key = self._suite_display_to_key.get(self.test_suite_var.get().strip(), "custom")
        suite_config = BACKTEST_TEST_SUITE_PRESETS.get(selected_suite_key, {}) if selected_suite_key != "custom" else {}
        suite_composition = dict(suite_config.get("composition", {})) if isinstance(suite_config, dict) else {}
        selected_scenario_packs = self._selected_listbox_values(self.scenario_pack_listbox)
        if not selected_scenario_packs and isinstance(suite_composition.get("scenario_packs"), list):
            selected_scenario_packs = [str(item) for item in suite_composition.get("scenario_packs", [])]
        selected_profile = self.selected_stress_profile_var.get().strip() or "Base"
        stress_controls = {
            "enable_historical_replay_regimes": bool(self.stress_enable_historical_replay_var.get()),
            "historical_window_fraction": float(parse_float(self.stress_historical_window_fraction_var.get()) or 0.20),
            "historical_replay_window_bars": int(parse_float(self.stress_historical_replay_window_bars_var.get()) or 20),
            "synthetic_jump_magnitude": float(parse_float(self.stress_synthetic_jump_magnitude_var.get()) or 0.02),
            "synthetic_jump_interval": int(parse_float(self.stress_synthetic_jump_interval_var.get()) or 7),
            "synthetic_vol_cluster_multiplier": float(parse_float(self.stress_synthetic_vol_cluster_multiplier_var.get()) or 1.6),
            "overlay_spread_multiplier": float(parse_float(self.stress_overlay_spread_multiplier_var.get()) or 2.5),
            "overlay_liquidity_multiplier": float(parse_float(self.stress_overlay_liquidity_multiplier_var.get()) or 0.4),
            "selected_profile": selected_profile,
            "selected_test_suite": selected_suite_key,
            "suite_composition": suite_composition,
            "run_mode": run_mode,
        }
        self.run_selection_summary_var.set(
            "Run config → mode: "
            + run_mode
            + ", stress profile: "
            + selected_profile
            + ", suite: "
            + selected_suite_key
            + ", scenario packs: "
            + (", ".join(selected_scenario_packs) if selected_scenario_packs else "none")
        )

        governance_payload = {
            "hypothesis_id": self.gov_hypothesis_id_var.get().strip(),
            "experiment_id": self.gov_experiment_id_var.get().strip(),
            "owner": self.gov_owner_var.get().strip(),
            "dataset_snapshot_lock": self.gov_dataset_lock_var.get().strip(),
            "acceptance_criteria": self.gov_acceptance_text.get("1.0", tk.END).strip(),
            "approval_status": self.gov_approval_status_var.get().strip() or "pending",
            "promotion_state": self.gov_promotion_state_var.get().strip() or "research",
            "min_oos_periods": int(parse_float(self.gov_min_oos_periods_var.get()) or 3),
            "min_stability_score": float(parse_float(self.gov_min_stability_var.get()) or 0.55),
            "max_turnover_total": float(parse_float(self.gov_max_turnover_var.get()) or 4.0),
            "min_capacity_score": float(parse_float(self.gov_min_capacity_var.get()) or 0.5),
            "max_signal_agreement_drift": float(parse_float(self.gov_max_signal_agreement_drift_var.get()) or 0.10),
            "max_fill_slippage_drift_bps": float(parse_float(self.gov_max_fill_slippage_drift_bps_var.get()) or 5.0),
            "max_pnl_attribution_divergence": float(parse_float(self.gov_max_pnl_attribution_divergence_var.get()) or 0.15),
            "expected_outcomes": {
                "signal_agreement": float(parse_float(self.gov_expected_signal_agreement_var.get()) or 1.0),
                "fill_slippage_bps": float(parse_float(self.gov_expected_fill_slippage_bps_var.get()) or 0.0),
                "pnl_attribution": float(parse_float(self.gov_expected_pnl_attribution_var.get()) or 1.0),
            },
            "observed_outcomes": {
                "signal_agreement": float(parse_float(self.gov_observed_signal_agreement_var.get()) or 1.0),
                "fill_slippage_bps": float(parse_float(self.gov_observed_fill_slippage_bps_var.get()) or 0.0),
                "pnl_attribution": float(parse_float(self.gov_observed_pnl_attribution_var.get()) or 1.0),
            },
            "comments": list(self._governance_comments),
            "review_actions": list(self._governance_review_actions),
            "decision_log": list(self._governance_decision_log),
            "selected_test_suite": selected_suite_key,
            "suite_composition": suite_composition,
        }

        decision_owner = self.gov_owner_var.get().strip() or "research_lab_ui"
        decision_timestamp = datetime.now().isoformat(timespec="seconds")
        self._governance_decision_log.append(
            {
                "owner": decision_owner,
                "decision": self.gov_approval_status_var.get().strip() or "pending",
                "reason": self.gov_acceptance_text.get("1.0", tk.END).strip(),
                "promotion_state": self.gov_promotion_state_var.get().strip() or "research",
                "timestamp": decision_timestamp,
            }
        )
        self._governance_review_actions.append(
            {
                "owner": decision_owner,
                "action": "run_requested",
                "status": self.gov_approval_status_var.get().strip() or "pending",
                "timestamp": decision_timestamp,
            }
        )

        worker_args: tuple[object, ...]
        status_line: str

        if selected_backtest_type == "trained_regime":
            regime_option = self._selected_regime_option()
            try:
                trained_regime_request = self._application_service.build_trained_regime_replay_request(
                    tickers=tickers,
                    start_date=start_date,
                    end_date=end_date,
                    run_mode=run_mode,
                    selected_backtest_type=selected_backtest_type,
                    timeframe=timeframe,
                    cache_root=cache_root,
                    governance_payload=governance_payload,
                    stress_controls=stress_controls,
                    selected_scenario_packs=selected_scenario_packs,
                    selected_suite_key=selected_suite_key,
                    suite_composition=suite_composition,
                    regime_option=regime_option,
                )
            except BacktestRequestValidationError as exc:
                messagebox.showinfo("Invalid input", str(exc))
                return
            except RegimeBundleCompatibilityError as exc:
                messagebox.showerror("Regime bundle compatibility error", f"Cannot run selected trained regime.\n\n{exc}")
                return
            worker_target = self._run_trained_regime_worker
            worker_args = (trained_regime_request,)
            status_line = f"Running trained regime policy '{trained_regime_request.regime_contract.regime_name}' ({run_mode})...\n"
        elif strategy == "momentum":
            selected_entries = self._selected_signal_names(self.entry_signal_vars)
            selected_exits = self._selected_signal_names(self.exit_signal_vars)
            if not selected_entries:
                messagebox.showinfo("Invalid input", "Select at least one entry signal.")
                return
            if not selected_exits:
                messagebox.showinfo("Invalid input", "Select at least one exit signal.")
                return

            optimizer_n_trials = int(parse_float(self.optimizer_trials_var.get()) or 20)
            if optimizer_n_trials <= 0:
                messagebox.showinfo("Invalid input", "Optimizer trials must be greater than zero.")
                return False
            optimizer_sampler = (self.optimizer_sampler_var.get().strip().lower() or "tpe")
            if optimizer_sampler not in set(OPTIMIZER_SAMPLERS):
                messagebox.showinfo("Invalid input", f"Optimizer sampler must be one of: {', '.join(OPTIMIZER_SAMPLERS)}.")
                return False
            optimizer_min_completed = int(parse_float(self.optimizer_min_completed_var.get()) or 5)
            if optimizer_min_completed < 1:
                messagebox.showinfo("Invalid input", "Min completed for pruning must be >= 1.")
                return False
            staged_budgets_raw = self.optimizer_staged_budgets_var.get().strip()
            staged_budgets: list[dict[str, object]] | None = None
            if staged_budgets_raw:
                try:
                    parsed_stage = json.loads(staged_budgets_raw)
                except json.JSONDecodeError:
                    messagebox.showinfo("Invalid input", "Staged budgets must be valid JSON list.")
                    return False
                if not isinstance(parsed_stage, list):
                    messagebox.showinfo("Invalid input", "Staged budgets must be a JSON list.")
                    return False
                staged_budgets = [dict(item) for item in parsed_stage if isinstance(item, dict)]

            if bool(self.use_optimizer_var.get()):
                search_space: dict[str, object] | None = None
                objectives: list[dict[str, str]] | None = None
                objective_weights: dict[str, float] | None = None
                overfitting_penalty: dict[str, float] | None = None
                max_turnover = parse_float(self.optimizer_max_turnover_var.get())
                max_drawdown_floor = parse_float(self.optimizer_max_drawdown_floor_var.get())
                min_trades = parse_float(self.optimizer_min_trades_var.get())

                search_space_raw = self.optimizer_search_space_var.get().strip()
                if search_space_raw:
                    try:
                        search_space_payload = json.loads(search_space_raw)
                    except json.JSONDecodeError:
                        messagebox.showinfo("Invalid input", "Search space must be valid JSON object.")
                        return False
                    if not isinstance(search_space_payload, dict):
                        messagebox.showinfo("Invalid input", "Search space must be a JSON object.")
                        return False
                    search_space = {str(k): v for k, v in search_space_payload.items()}

                objectives_raw = self.optimizer_objectives_var.get().strip()
                if objectives_raw:
                    try:
                        objectives_payload = json.loads(objectives_raw)
                    except json.JSONDecodeError:
                        messagebox.showinfo("Invalid input", "Objectives must be valid JSON list.")
                        return False
                    if not isinstance(objectives_payload, list):
                        messagebox.showinfo("Invalid input", "Objectives must be a JSON list.")
                        return False
                    objectives = [dict(item) for item in objectives_payload if isinstance(item, dict)]

                objective_weights_raw = self.optimizer_objective_weights_var.get().strip()
                if objective_weights_raw:
                    try:
                        objective_weights_payload = json.loads(objective_weights_raw)
                    except json.JSONDecodeError:
                        messagebox.showinfo("Invalid input", "Objective weights must be valid JSON object.")
                        return False
                    if not isinstance(objective_weights_payload, dict):
                        messagebox.showinfo("Invalid input", "Objective weights must be a JSON object.")
                        return False
                    objective_weights = {str(k): float(v) for k, v in objective_weights_payload.items()}

                overfitting_penalty_raw = self.optimizer_overfitting_penalty_var.get().strip()
                if overfitting_penalty_raw:
                    try:
                        overfitting_penalty_payload = json.loads(overfitting_penalty_raw)
                    except json.JSONDecodeError:
                        messagebox.showinfo("Invalid input", "Overfitting penalty must be valid JSON object.")
                        return False
                    if not isinstance(overfitting_penalty_payload, dict):
                        messagebox.showinfo("Invalid input", "Overfitting penalty must be a JSON object.")
                        return False
                    overfitting_penalty = {str(k): float(v) for k, v in overfitting_penalty_payload.items()}

                worker_target = self._run_optimizer_worker
                worker_args = (
                    tickers,
                    start_date,
                    end_date,
                    cache_root,
                    lookback,
                    skip,
                    costs_bps,
                    selected_entries,
                    selected_exits,
                    execution_model,
                    execution_model_params,
                    governance_payload,
                    stress_controls,
                    optimizer_n_trials,
                    optimizer_sampler,
                    bool(self.optimizer_enable_pruning_var.get()),
                    bool(self.optimizer_prune_constraint_var.get()),
                    bool(self.optimizer_prune_lcb_var.get()),
                    optimizer_min_completed,
                    staged_budgets,
                    search_space,
                    objectives,
                    max_turnover,
                    max_drawdown_floor,
                    min_trades,
                    objective_weights,
                    overfitting_penalty,
                )
                status_line = f"Running optimizer across {len(selected_entries) * len(selected_exits)} candidates...\n"
            elif bool(self.use_walk_forward_var.get()):
                walk_forward_windows = self._validate_walk_forward_inputs()
                if walk_forward_windows is None:
                    return
                wf_payload, purge_bars, embargo_bars, cpcv_groups, cpcv_test_groups, cv_seed = walk_forward_windows
                worker_target = self._run_walk_forward_worker
                worker_args = (
                    tickers,
                    start_date,
                    end_date,
                    cache_root,
                    lookback,
                    skip,
                    costs_bps,
                    selected_entries,
                    selected_exits,
                    execution_model,
                    execution_model_params,
                    wf_payload,
                    purge_bars,
                    embargo_bars,
                    cpcv_groups,
                    cpcv_test_groups,
                    cv_seed,
                    governance_payload,
                    stress_controls,
                    selected_scenario_packs,
                    run_mode,
                )
                status_line = f"Running walk-forward with {len(selected_entries) * len(selected_exits)} candidates...\n"
            else:
                try:
                    classic_request = self._application_service.build_classic_strategy_request(
                        tickers=tickers,
                        start_date=start_date,
                        end_date=end_date,
                        run_mode=run_mode,
                        selected_backtest_type=selected_backtest_type,
                        strategy=strategy,
                        lookback=lookback,
                        skip=skip,
                        costs_bps=costs_bps,
                        starting_capital=starting_capital,
                        custom_bet_pct=custom_bet_pct,
                        cache_root=cache_root,
                        bet_sizing_mode=bet_sizing_mode,
                        timeframe=timeframe,
                        execution_model=execution_model,
                        execution_model_params=execution_model_params,
                        portfolio_cfg=portfolio_cfg,
                        governance_payload=governance_payload,
                        stress_controls=stress_controls,
                        selected_scenario_packs=selected_scenario_packs,
                        selected_suite_key=selected_suite_key,
                        suite_composition=suite_composition,
                    )
                except BacktestRequestValidationError as exc:
                    messagebox.showinfo("Invalid input", str(exc))
                    return
                worker_target = self._run_momentum_worker
                worker_args = (classic_request, selected_entries, selected_exits)
                status_line = f"Running {len(selected_entries) * len(selected_exits)} momentum entry/exit combinations ({run_mode})...\n"
        else:
            xsmom_valid = self._validate_xsmom_inputs()
            if xsmom_valid is None:
                return
            xsmom_top_quantile, xsmom_bottom_quantile, xsmom_vol_lookback_days = xsmom_valid
            xsmom_long_only = bool(self.xsmom_long_only_var.get())
            worker_target = self._run_xsmom_worker
            worker_args = (
                tickers,
                start_date,
                end_date,
                cache_root,
                lookback,
                skip,
                costs_bps,
                starting_capital,
                bet_sizing_mode,
                custom_bet_pct,
                timeframe,
                xsmom_top_quantile,
                xsmom_bottom_quantile,
                xsmom_long_only,
                xsmom_vol_lookback_days,
                execution_model,
                execution_model_params,
                portfolio_cfg,
                governance_payload,
                stress_controls,
            )
            status_line = "Running cross-sectional momentum backtest...\n"

        regime_option = self._selected_regime_option()
        self.controller.state.backtest_settings.update(
            {
                "selected_backtest_type": selected_backtest_type,
                "selected_trained_regime": self.trained_regime_var.get().strip(),
                "selected_trained_regime_option_id": regime_option.option_id if regime_option else "",
                "selected_trained_regime_manifest_path": regime_option.manifest_path if regime_option else "",
                "selected_trained_regime_source": regime_option.source if regime_option else "",
            }
        )
        self.controller.persist_state()

        self._set_run_controls_state("disabled")
        self.logs_text.delete("1.0", tk.END)
        self.logs_text.insert("1.0", status_line)
        self.section_notebook.select(5)

        thread = threading.Thread(target=worker_target, args=worker_args, daemon=True)
        thread.start()

    def _validate_common_inputs(self) -> tuple[str, int, int, float, float, float] | None:
        strategy = self.strategy_var.get().strip() or "momentum"
        if strategy not in STRATEGIES:
            messagebox.showinfo("Invalid input", "Please select a valid strategy.")
            return None

        lookback = parse_float(self.lookback_days_var.get())
        skip = parse_float(self.skip_days_var.get())
        costs_bps = parse_float(self.costs_bps_var.get())
        starting_capital = parse_float(self.starting_capital_var.get())
        custom_bet_pct = parse_float(self.custom_bet_pct_var.get())

        if lookback is None or lookback < 1 or int(lookback) != lookback:
            messagebox.showinfo("Invalid input", "Lookback must be a positive integer.")
            return None
        if skip is None or skip < 0 or int(skip) != skip:
            messagebox.showinfo("Invalid input", "Skip must be a non-negative integer.")
            return None
        if costs_bps is None or costs_bps < 0:
            messagebox.showinfo("Invalid input", "Costs must be zero or positive.")
            return None
        if starting_capital is None or starting_capital <= 0:
            messagebox.showinfo("Invalid input", "Starting capital must be > 0.")
            return None
        if custom_bet_pct is None or custom_bet_pct <= 0:
            messagebox.showinfo("Invalid input", "Custom bet % must be > 0.")
            return None

        timeframe = self.timeframe_var.get().strip() or "1m"
        if timeframe not in TIMEFRAMES:
            messagebox.showinfo("Invalid input", "Please select a valid resolution.")
            return None
        return (
            strategy,
            int(lookback),
            int(skip),
            float(costs_bps),
            float(starting_capital),
            float(custom_bet_pct),
        )

    def _validate_xsmom_inputs(self) -> tuple[float, float, int] | None:
        top_quantile = parse_float(self.xsmom_top_quantile_var.get())
        bottom_quantile = parse_float(self.xsmom_bottom_quantile_var.get())
        vol_lookback = parse_float(self.xsmom_vol_lookback_days_var.get())

        if top_quantile is None or top_quantile <= 0 or top_quantile > 1:
            messagebox.showinfo("Invalid input", "Top quantile must be in (0, 1].")
            return None
        if bottom_quantile is None or bottom_quantile < 0 or bottom_quantile > 1:
            messagebox.showinfo("Invalid input", "Bottom quantile must be in [0, 1].")
            return None
        if vol_lookback is None or vol_lookback < 2 or int(vol_lookback) != vol_lookback:
            messagebox.showinfo("Invalid input", "Volatility lookback must be an integer >= 2.")
            return None

        return float(top_quantile), float(bottom_quantile), int(vol_lookback)

    def _validate_walk_forward_inputs(self) -> tuple[dict[str, float | int | str | bool | dict[str, float] | list[str] | None], int, int, int, int, int] | None:
        train_bars = parse_float(self.wf_train_bars_var.get())
        validation_bars = parse_float(self.wf_validation_bars_var.get())
        test_bars = parse_float(self.wf_test_bars_var.get())
        step_bars = parse_float(self.wf_step_bars_var.get())
        bars_values = (train_bars, validation_bars, test_bars, step_bars)
        has_any_bars = any(value is not None for value in bars_values)
        has_all_bars = all(value is not None for value in bars_values)

        fractions = {
            "train_fraction": float(self.wf_train_fraction_var.get()),
            "validation_fraction": float(self.wf_validation_fraction_var.get()),
            "test_fraction": float(self.wf_test_fraction_var.get()),
            "step_fraction": float(self.wf_step_fraction_var.get()),
        }

        if has_any_bars and not has_all_bars:
            messagebox.showinfo("Invalid input", "Bars mode requires train/validation/test/step bars together.")
            return None

        if has_all_bars:
            if any(value is None or value <= 0 or int(value) != value for value in bars_values):
                messagebox.showinfo("Invalid input", "Walk-forward bars must be positive integers.")
                return None
            wf_payload: dict[str, float | int | str | bool | dict[str, float] | list[str] | None] = {
                "train_bars": int(train_bars),
                "validation_bars": int(validation_bars),
                "test_bars": int(test_bars),
                "step_bars": int(step_bars),
                "train_fraction": None,
                "validation_fraction": None,
                "test_fraction": None,
                "step_fraction": None,
            }
        else:
            values = [fractions["train_fraction"], fractions["validation_fraction"], fractions["test_fraction"]]
            if any(value <= 0.0 or value >= 1.0 for value in values):
                messagebox.showinfo("Invalid input", "Train/validation/test fractions must be in (0, 1).")
                return None
            if abs(sum(values) - 1.0) > 1e-6:
                messagebox.showinfo("Invalid input", "Train, validation, and test fractions must sum to 1.0.")
                return None
            if fractions["step_fraction"] <= 0.0 or fractions["step_fraction"] > 1.0:
                messagebox.showinfo("Invalid input", "Step fraction must be in (0, 1].")
                return None
            wf_payload = {
                "train_bars": None,
                "validation_bars": None,
                "test_bars": None,
                "step_bars": None,
                "train_fraction": fractions["train_fraction"],
                "validation_fraction": fractions["validation_fraction"],
                "test_fraction": fractions["test_fraction"],
                "step_fraction": fractions["step_fraction"],
            }

        cv_scheme = self.wf_cv_scheme_var.get().strip() or "walk_forward"
        purge = parse_float(self.wf_purge_bars_var.get())
        embargo = parse_float(self.wf_embargo_bars_var.get())
        n_groups = parse_float(self.wf_cpcv_groups_var.get())
        n_test_groups = parse_float(self.wf_cpcv_test_groups_var.get())
        cv_seed = parse_float(self.wf_cv_seed_var.get())
        label_horizon_bars = parse_float(self.wf_label_horizon_bars_var.get())
        inner_train_fraction = parse_float(self.wf_inner_train_fraction_var.get())
        if purge is None or purge < 0 or int(purge) != purge:
            messagebox.showinfo("Invalid input", "Purge bars must be a non-negative integer.")
            return None
        if embargo is None or embargo < 0 or int(embargo) != embargo:
            messagebox.showinfo("Invalid input", "Embargo bars must be a non-negative integer.")
            return None
        if n_groups is None or n_groups < 3 or int(n_groups) != n_groups:
            messagebox.showinfo("Invalid input", "CPCV groups must be an integer >= 3.")
            return None
        if n_test_groups is None or n_test_groups < 1 or int(n_test_groups) != n_test_groups:
            messagebox.showinfo("Invalid input", "CPCV test groups must be an integer >= 1.")
            return None
        if cv_seed is None or int(cv_seed) != cv_seed:
            messagebox.showinfo("Invalid input", "CV seed must be an integer.")
            return None
        if label_horizon_bars is None or label_horizon_bars < 1 or int(label_horizon_bars) != label_horizon_bars:
            messagebox.showinfo("Invalid input", "Label horizon bars must be an integer >= 1.")
            return None
        if inner_train_fraction is None or inner_train_fraction <= 0.0 or inner_train_fraction >= 1.0:
            messagebox.showinfo("Invalid input", "Inner train fraction must be in (0, 1).")
            return None

        objective_weights: dict[str, float] | None = None
        overfitting_penalty: dict[str, float] | None = None
        if self.wf_objective_weights_var.get().strip():
            try:
                payload = json.loads(self.wf_objective_weights_var.get().strip())
            except json.JSONDecodeError:
                messagebox.showinfo("Invalid input", "WF objective weights must be valid JSON object.")
                return None
            if not isinstance(payload, dict):
                messagebox.showinfo("Invalid input", "WF objective weights must be a JSON object.")
                return None
            objective_weights = {str(k): float(v) for k, v in payload.items()}
        if self.wf_overfitting_penalty_var.get().strip():
            try:
                payload = json.loads(self.wf_overfitting_penalty_var.get().strip())
            except json.JSONDecodeError:
                messagebox.showinfo("Invalid input", "WF overfitting penalty must be valid JSON object.")
                return None
            if not isinstance(payload, dict):
                messagebox.showinfo("Invalid input", "WF overfitting penalty must be a JSON object.")
                return None
            overfitting_penalty = {str(k): float(v) for k, v in payload.items()}

        prior_strategy_keys = [item.strip() for item in self.wf_prior_strategy_keys_var.get().split(",") if item.strip()]
        wf_payload.update({
            "cv_scheme": cv_scheme,
            "label_horizon_bars": int(label_horizon_bars),
            "nested_optimization": bool(self.wf_nested_optimization_var.get()),
            "inner_train_fraction": float(inner_train_fraction),
            "objective_weights": objective_weights,
            "overfitting_penalty": overfitting_penalty,
            "strategy_key": self.wf_strategy_key_var.get().strip() or None,
            "prior_strategy_keys": prior_strategy_keys or None,
        })

        return wf_payload, int(purge), int(embargo), int(n_groups), int(n_test_groups), int(cv_seed)

    def _build_optimizer_json_input(
        self,
        parent: ttk.Frame,
        row: int,
        label_text: str,
        variable: tk.StringVar,
        default_value: str,
        field_key: str,
    ) -> None:
        ttk.Label(parent, text=label_text).grid(row=row, column=0, sticky="w", padx=8, pady=6)
        field_row = ttk.Frame(parent)
        field_row.grid(row=row, column=1, sticky="ew", padx=8, pady=6)
        field_row.columnconfigure(0, weight=1)
        ttk.Entry(field_row, textvariable=variable).grid(row=0, column=0, sticky="ew")
        ttk.Button(field_row, text="Reset to defaults", command=lambda: variable.set(default_value)).grid(row=0, column=1, sticky="e", padx=(8, 0))
        lint_var = tk.StringVar(value="JSON looks valid.")
        self._optimizer_json_lint_vars[field_key] = lint_var
        ttk.Label(parent, textvariable=lint_var, foreground="#2f6f44", justify="left").grid(row=row + 1, column=1, sticky="w", padx=8, pady=(0, 4))
        variable.trace_add("write", lambda *_: self._update_optimizer_json_lint(field_key, variable))
        self._update_optimizer_json_lint(field_key, variable)

    def _update_optimizer_json_lint(self, field_key: str, variable: tk.StringVar) -> None:
        lint_var = self._optimizer_json_lint_vars.get(field_key)
        if lint_var is None:
            return
        raw_value = variable.get().strip()
        if not raw_value:
            lint_var.set("Using optimizer defaults (empty field).")
            return
        try:
            json.loads(raw_value)
        except json.JSONDecodeError as exc:
            lint_var.set(f"Invalid JSON: {exc.msg} (line {exc.lineno}, col {exc.colno})")
            return
        lint_var.set("JSON looks valid.")

    def _toggle_advanced_optimization(self) -> None:
        if self.show_advanced_optimization_var.get():
            self.advanced_optimization_frame.grid()
        else:
            self.advanced_optimization_frame.grid_remove()

    def _restore_advanced_defaults(self) -> None:
        defaults = {
            "custom_bet_pct": "10",
            "portfolio_method": "equal_weight",
            "portfolio_vol_lookback": "20",
            "portfolio_target_vol": "0.10",
            "portfolio_max_symbol": "0.25",
            "portfolio_max_sector": "0.60",
            "portfolio_rebalance_frequency": "1",
            "portfolio_clustering_linkage": "single",
            "portfolio_covariance_shrinkage": "0.0",
            "portfolio_max_gross": "1.0",
            "portfolio_min_net": "-1.0",
            "portfolio_max_net": "1.0",
            "use_walk_forward": False,
            "use_optimizer": False,
        }
        self.custom_bet_pct_var.set(str(defaults["custom_bet_pct"]))
        self.portfolio_method_var.set(str(defaults["portfolio_method"]))
        self.portfolio_vol_lookback_var.set(str(defaults["portfolio_vol_lookback"]))
        self.portfolio_target_vol_var.set(str(defaults["portfolio_target_vol"]))
        self.portfolio_max_symbol_var.set(str(defaults["portfolio_max_symbol"]))
        self.portfolio_max_sector_var.set(str(defaults["portfolio_max_sector"]))
        self.portfolio_rebalance_frequency_var.set(str(defaults["portfolio_rebalance_frequency"]))
        self.portfolio_clustering_linkage_var.set(str(defaults["portfolio_clustering_linkage"]))
        self.portfolio_covariance_shrinkage_var.set(str(defaults["portfolio_covariance_shrinkage"]))
        self.portfolio_max_gross_var.set(str(defaults["portfolio_max_gross"]))
        self.portfolio_min_net_var.set(str(defaults["portfolio_min_net"]))
        self.portfolio_max_net_var.set(str(defaults["portfolio_max_net"]))
        self.use_walk_forward_var.set(bool(defaults["use_walk_forward"]))
        self.use_optimizer_var.set(bool(defaults["use_optimizer"]))
        self._toggle_advanced_optimization()
        self._update_validation_hint()

    def _attach_tooltip(self, widget: tk.Widget, text: str) -> None:
        tooltip_window: tk.Toplevel | None = None

        def show_tooltip(_event: object) -> None:
            nonlocal tooltip_window
            if tooltip_window is not None:
                return
            tooltip_window = tk.Toplevel(self)
            tooltip_window.wm_overrideredirect(True)
            x = widget.winfo_rootx() + 16
            y = widget.winfo_rooty() + 16
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

    def _run_optimizer_worker(
        self,
        tickers: list[str],
        start_date: date,
        end_date: date,
        cache_root: Path,
        lookback: int,
        skip: int,
        costs_bps: float,
        entry_signals: list[str],
        exit_signals: list[str],
        execution_model: str,
        execution_model_params: dict[str, object],
        governance_payload: dict[str, object],
        stress_controls: dict[str, object],
        optimizer_n_trials: int,
        optimizer_sampler: str,
        enable_pruning: bool,
        prune_on_constraint_violation: bool,
        prune_on_lcb: bool,
        min_completed_for_pruning: int,
        staged_budgets: list[dict[str, object]] | None,
        search_space: dict[str, object] | None,
        objectives: list[dict[str, str]] | None,
        max_turnover: float | None,
        max_drawdown_floor: float | None,
        min_trades: float | None,
        objective_weights: dict[str, float] | None,
        overfitting_penalty: dict[str, float] | None,
    ) -> None:
        try:
            entry_grid = {signal: [{}] for signal in entry_signals}
            exit_grid = {signal: [{}] for signal in exit_signals}
            core_grid = {
                "lookback_days": [int(lookback)],
                "skip_days": [int(skip)],
                "costs_bps": [float(costs_bps)],
            }
            output_text = self._run_backend_job(JOB_BACKTEST_OPTIMIZATION, dict(
                tickers=tickers,
                start_date=start_date,
                end_date=end_date,
                cache_root=cache_root,
                entry_grid=entry_grid,
                exit_grid=exit_grid,
                core_grid=core_grid,
                seed=42,
                n_trials=max(1, int(optimizer_n_trials)),
                sampler_name=str(optimizer_sampler),
                search_space=None if search_space is None else dict(search_space),
                objectives=None if objectives is None else [dict(item) for item in objectives],
                max_turnover=None if max_turnover is None else float(max_turnover),
                max_drawdown_floor=None if max_drawdown_floor is None else float(max_drawdown_floor),
                min_trades=None if min_trades is None else float(min_trades),
                partial_period_fractions=[0.33, 0.66, 1.0],
                enable_pruning=bool(enable_pruning),
                prune_on_constraint_violation=bool(prune_on_constraint_violation),
                prune_on_lcb=bool(prune_on_lcb),
                min_completed_for_pruning=int(min_completed_for_pruning),
                staged_budgets=None if staged_budgets is None else [dict(stage) for stage in staged_budgets],
                objective_weights=None if objective_weights is None else dict(objective_weights),
                overfitting_penalty=None if overfitting_penalty is None else dict(overfitting_penalty),
                governance_metadata=dict(governance_payload),
                stress_controls=dict(stress_controls),
            ))
        except Exception as exc:
            output_text = f"Backtest failed: {exc}"
        self.after(0, lambda: self._finish_backtest_run(output_text))

    def _run_walk_forward_worker(
        self,
        tickers: list[str],
        start_date: date,
        end_date: date,
        cache_root: Path,
        lookback: int,
        skip: int,
        costs_bps: float,
        entry_signals: list[str],
        exit_signals: list[str],
        execution_model: str,
        execution_model_params: dict[str, object],
        wf_payload: dict[str, float | int | str | bool | dict[str, float] | list[str] | None],
        purge_window_bars: int,
        embargo_window_bars: int,
        cpcv_n_groups: int,
        cpcv_n_test_groups: int,
        cv_seed: int,
        governance_payload: dict[str, object],
        stress_controls: dict[str, object],
        scenario_packs: list[str],
        run_mode: str,
    ) -> None:
        try:
            entry_grid = {signal: [{}] for signal in entry_signals}
            exit_grid = {signal: [{}] for signal in exit_signals}
            core_grid = {
                "lookback_days": [int(lookback)],
                "skip_days": [int(skip)],
                "costs_bps": [float(costs_bps)],
            }
            output_text = self._run_backend_job(JOB_BACKTEST_WALK_FORWARD, dict(
                tickers=tickers,
                start_date=start_date,
                end_date=end_date,
                cache_root=cache_root,
                entry_grid=entry_grid,
                exit_grid=exit_grid,
                core_grid=core_grid,
                train_bars=None if wf_payload.get("train_bars") is None else int(wf_payload["train_bars"]),
                validation_bars=None if wf_payload.get("validation_bars") is None else int(wf_payload["validation_bars"]),
                test_bars=None if wf_payload.get("test_bars") is None else int(wf_payload["test_bars"]),
                step_bars=None if wf_payload.get("step_bars") is None else int(wf_payload["step_bars"]),
                train_fraction=None if wf_payload.get("train_fraction") is None else float(wf_payload["train_fraction"]),
                validation_fraction=None if wf_payload.get("validation_fraction") is None else float(wf_payload["validation_fraction"]),
                test_fraction=None if wf_payload.get("test_fraction") is None else float(wf_payload["test_fraction"]),
                step_fraction=None if wf_payload.get("step_fraction") is None else float(wf_payload["step_fraction"]),
                cv_scheme=str(wf_payload.get("cv_scheme", "walk_forward")),
                purge_window_bars=int(purge_window_bars),
                embargo_window_bars=int(embargo_window_bars),
                cpcv_n_groups=int(cpcv_n_groups),
                cpcv_n_test_groups=int(cpcv_n_test_groups),
                cv_seed=int(cv_seed),
                label_horizon_bars=int(wf_payload.get("label_horizon_bars", 1)),
                nested_optimization=bool(wf_payload.get("nested_optimization", False)),
                inner_train_fraction=float(wf_payload.get("inner_train_fraction", 0.7)),
                objective_weights=None if wf_payload.get("objective_weights") is None else dict(wf_payload["objective_weights"]),
                overfitting_penalty=None if wf_payload.get("overfitting_penalty") is None else dict(wf_payload["overfitting_penalty"]),
                strategy_key=None if wf_payload.get("strategy_key") is None else str(wf_payload["strategy_key"]),
                prior_strategy_keys=None if wf_payload.get("prior_strategy_keys") is None else list(wf_payload["prior_strategy_keys"]),
                governance_metadata=dict(governance_payload),
                stress_controls=dict(stress_controls),
            ))
        except Exception as exc:
            output_text = f"Backtest failed: {exc}"
        self.after(0, lambda: self._finish_backtest_run(output_text))

    def _run_momentum_worker(
        self,
        request: ClassicStrategyRunRequest,
        entry_signals: list[str],
        exit_signals: list[str],
    ) -> None:
        try:
            controls = dict(request.stress_controls)
            controls["run_mode"] = request.run_mode
            if request.run_mode == "stress_only":
                controls["stress_only"] = True
            output_text = self._run_backend_job(JOB_BACKTEST_MULTI_SIGNAL, dict(
                tickers=request.tickers,
                start_date=request.start_date,
                end_date=request.end_date,
                cache_root=request.artifact_routing.cache_root,
                lookback_days=request.lookback,
                skip_days=request.skip,
                costs_bps=request.costs_bps,
                execution_model=request.execution_model,
                execution_model_params=request.execution_model_params,
                starting_capital=request.starting_capital,
                bet_sizing_mode=request.bet_sizing_mode,
                custom_bet_pct=request.custom_bet_pct,
                timeframe=request.timeframe,
                entry_signals=entry_signals,
                exit_signals=exit_signals,
                **request.portfolio_cfg,
                governance_metadata=dict(request.governance_payload),
                stress_controls=controls,
                scenario_packs=list(request.selected_scenario_packs),
            ))
            output_text += "\nApplied stress profile: " + str(request.stress_controls.get("selected_profile", "Base")) + " | Scenario packs: " + (", ".join(request.selected_scenario_packs) if request.selected_scenario_packs else "none") + " | Run mode: " + request.run_mode
        except Exception as exc:
            output_text = f"Backtest failed: {exc}"
        self.after(0, lambda: self._finish_backtest_run(output_text))

    def _run_trained_regime_worker(
        self,
        request: TrainedRegimeReplayRunRequest,
    ) -> None:
        try:
            output_text = self._run_backend_job(JOB_BACKTEST_TRAINED_REGIME, dict(
                tickers=request.tickers,
                start_date=request.start_date,
                end_date=request.end_date,
                cache_root=request.artifact_routing.cache_root,
                timeframe=request.timeframe,
                regime_contract=request.regime_contract,
                governance_metadata=dict(request.governance_payload),
                stress_controls=dict(request.stress_controls),
                scenario_packs=list(request.selected_scenario_packs),
                selected_test_suite=str(request.selected_suite_key),
                suite_composition=dict(request.suite_composition),
            ))
        except Exception as exc:
            output_text = f"Backtest failed: {exc}"
        self.after(0, lambda: self._finish_backtest_run(output_text))

    def _run_xsmom_worker(
        self,
        tickers: list[str],
        start_date: date,
        end_date: date,
        cache_root: Path,
        lookback: int,
        skip: int,
        costs_bps: float,
        starting_capital: float,
        bet_sizing_mode: str,
        custom_bet_pct: float,
        timeframe: str,
        top_quantile: float,
        bottom_quantile: float,
        long_only: bool,
        vol_lookback_days: int,
        execution_model: str,
        execution_model_params: dict[str, object],
        portfolio_cfg: dict[str, object],
        governance_payload: dict[str, object],
        stress_controls: dict[str, object],
        optimizer_n_trials: int,
        optimizer_sampler: str,
        enable_pruning: bool,
        prune_on_constraint_violation: bool,
        prune_on_lcb: bool,
        min_completed_for_pruning: int,
        staged_budgets: list[dict[str, object]] | None,
        search_space: dict[str, object] | None,
        objectives: list[dict[str, str]] | None,
        max_turnover: float | None,
        max_drawdown_floor: float | None,
        min_trades: float | None,
        objective_weights: dict[str, float] | None,
        overfitting_penalty: dict[str, float] | None,
    ) -> None:
        try:
            output_text = self._run_backend_job(JOB_BACKTEST_TIME_SERIES, dict(
                tickers=tickers,
                start_date=start_date,
                end_date=end_date,
                cache_root=cache_root,
                lookback_days=lookback,
                skip_days=skip,
                costs_bps=costs_bps,
                execution_model=execution_model,
                execution_model_params=execution_model_params,
                starting_capital=starting_capital,
                bet_sizing_mode=bet_sizing_mode,
                custom_bet_pct=custom_bet_pct,
                strategy="xsmom",
                xsmom_top_quantile=top_quantile,
                xsmom_bottom_quantile=bottom_quantile,
                xsmom_long_only=long_only,
                xsmom_vol_lookback_days=vol_lookback_days,
                timeframe=timeframe,
                **portfolio_cfg,
                governance_metadata=dict(governance_payload),
                stress_controls=dict(stress_controls),
            ))
        except Exception as exc:
            output_text = f"Backtest failed: {exc}"
        self.after(0, lambda: self._finish_backtest_run(output_text))


    def _run_backend_job(self, job_type: str, payload: dict[str, object]) -> object:
        backend = self.controller.execution_backend
        job_id = backend.submit_job(job_type, payload)
        while True:
            status = backend.get_status(job_id)
            if status in {"succeeded", "failed"}:
                break
            threading.Event().wait(0.1)
        if status == "failed":
            logs = backend.stream_logs(job_id)
            raise RuntimeError(logs[-1] if logs else f"{job_type} failed")
        return backend.get_result(job_id) if hasattr(backend, "get_result") else ""

    def _set_run_controls_state(self, state: str) -> None:
        if hasattr(self, "run_stress_only_button"):
            self.run_stress_only_button.config(state=state)
        if hasattr(self, "run_full_chain_button"):
            self.run_full_chain_button.config(state=state)
        elif hasattr(self, "run_button"):
            self.run_button.config(state=state)

    def _selected_listbox_values(self, listbox: tk.Listbox) -> list[str]:
        return [str(listbox.get(index)) for index in listbox.curselection()]

    def _set_listbox_selection(self, listbox: tk.Listbox, values: list[str], *, valid_options: tuple[str, ...]) -> None:
        filtered = [item for item in values if item in valid_options]
        options = [str(listbox.get(index)) for index in range(listbox.size())]
        listbox.selection_clear(0, "end")
        for item in filtered:
            if item in options:
                listbox.selection_set(options.index(item))

    def _apply_stress_profile(self, profile_name: str) -> None:
        profile = STRESS_PROFILES.get(profile_name)
        if not profile:
            return
        self.selected_stress_profile_var.set(profile_name)
        self.stress_historical_window_fraction_var.set(f"{float(profile['historical_window_fraction']):.2f}")
        self.stress_historical_replay_window_bars_var.set(str(int(profile['historical_replay_window_bars'])))
        self.stress_synthetic_jump_magnitude_var.set(f"{float(profile['synthetic_jump_magnitude']):.3f}")
        self.stress_synthetic_jump_interval_var.set(str(int(profile['synthetic_jump_interval'])))
        self.stress_synthetic_vol_cluster_multiplier_var.set(f"{float(profile['synthetic_vol_cluster_multiplier']):.2f}")
        self.stress_overlay_spread_multiplier_var.set(f"{float(profile['overlay_spread_multiplier']):.2f}")
        self.stress_overlay_liquidity_multiplier_var.set(f"{float(profile['overlay_liquidity_multiplier']):.2f}")

    def _finish_backtest_run(self, output_text: str) -> None:
        self._consume_run_outputs(output_text)
        self._set_run_controls_state("normal")

    def _resolve_supported_csv_setting(
        self,
        raw: object,
        *,
        supported: tuple[str, ...],
        fallback: tuple[str, ...],
        field_name: str,
    ) -> tuple[set[str], list[str], dict[str, str]]:
        values = [part.strip() for part in str(raw).split(",") if part.strip()]
        valid, stale, migrations = validate_option_values(values, supported=supported, field_name=field_name)
        if not valid:
            valid = list(fallback)
        return set(valid), stale, migrations

    def _split_csv_setting(self, raw: object) -> set[str]:
        return {part.strip() for part in str(raw).split(",") if part.strip()}

    def _selected_signal_names(self, signals: dict[str, tk.BooleanVar]) -> list[str]:
        return [name for name, var in signals.items() if bool(var.get())]

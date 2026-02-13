from __future__ import annotations

import threading
import csv
import json
import webbrowser
from datetime import date
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

from backtesting.cache_runner import (
    run_multi_signal_backtest,
    run_strategy_optimization,
    run_time_series_momentum_backtest,
    run_walk_forward_backtest,
)
from config import BACKTEST_CACHE_DIR, BACKTEST_OUTPUT_DIR, BACKTEST_STRATEGY_PRESETS, DEFAULT_BACKTEST_SETTINGS
from ui.backtesting_insights import (
    build_guardrails,
    build_scenario_comparison,
    fold_variance_rows,
    metric_deltas,
    parameter_diffs,
    parse_tags,
    read_experiment_index,
    read_stress_scenarios,
)
from utils.parsing import normalize_cache_root, parse_date, parse_float

ENTRY_SIGNALS = ["ts_momentum", "ma_trend", "breakout"]
EXIT_SIGNALS = ["none", "momentum_flip", "trailing_stop", "max_hold"]
STRATEGIES = ["momentum", "xsmom"]
TIMEFRAMES = ["1m", "5m", "15m", "30m", "1h", "1d"]
PORTFOLIO_METHODS = ["equal_weight", "vol_target", "inverse_vol", "capped_optimization", "hrp", "herc"]
EXECUTION_MODELS = ["bps", "spread", "participation", "square_root", "latency_drift", "modular", "volatility_scaled"]

TIMEFRAME_HISTORY_DAYS = {"1m": 14, "5m": 30, "15m": 60, "30m": 120, "1h": 365, "1d": 3650}
GOVERNANCE_PROMOTION_STATES = ["research", "paper", "shadow", "production"]
GOVERNANCE_APPROVAL_STATES = ["pending", "in_review", "approved", "rejected", "waived"]

class BacktestingPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller
        self._updating_wf_fractions = False
        self._advanced_widgets: dict[str, tk.Widget] = {}
        self._validation_messages: list[str] = []

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

        workflow_frame = ttk.LabelFrame(run_setup_tab, text="Workflow Mode & Presets")
        workflow_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(2, 8))
        workflow_frame.columnconfigure(1, weight=1)
        ttk.Label(workflow_frame, text="Mode").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        mode_row = ttk.Frame(workflow_frame)
        mode_row.grid(row=0, column=1, sticky="w", padx=8, pady=6)
        ttk.Radiobutton(mode_row, text="Basic", value="basic", variable=self.ui_mode_var, command=self._on_mode_changed).pack(side="left")
        ttk.Radiobutton(mode_row, text="Advanced", value="advanced", variable=self.ui_mode_var, command=self._on_mode_changed).pack(side="left", padx=(8, 0))

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

        strategy_frame = ttk.LabelFrame(run_setup_tab, text="Strategy")
        strategy_frame.grid(row=3, column=0, sticky="nsew", padx=10, pady=10)
        strategy_frame.columnconfigure(1, weight=1)

        row = 0
        ttk.Label(strategy_frame, text="Strategy").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.strategy_var = tk.StringVar(value="momentum")
        strategy_combo = ttk.Combobox(
            strategy_frame,
            textvariable=self.strategy_var,
            state="readonly",
            values=STRATEGIES,
        )
        strategy_combo.grid(row=row, column=1, sticky="ew", padx=8, pady=6)
        strategy_combo.bind("<<ComboboxSelected>>", self._on_strategy_changed)

        row += 1
        ttk.Label(strategy_frame, text="Lookback (bars)").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.lookback_days_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.lookback_days_var).grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="Skip (bars)").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.skip_days_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.skip_days_var).grid(row=row, column=1, sticky="ew", padx=8, pady=6)

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
        ttk.Label(strategy_frame, text="Starting Capital").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.starting_capital_var = tk.StringVar()
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
        ttk.Label(strategy_frame, text="Max Net Gamma").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.portfolio_max_net_gamma_var = tk.StringVar(value="")
        ttk.Entry(strategy_frame, textvariable=self.portfolio_max_net_gamma_var).grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="Max Abs Vega Bucket").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.portfolio_max_abs_vega_bucket_var = tk.StringVar(value="")
        ttk.Entry(strategy_frame, textvariable=self.portfolio_max_abs_vega_bucket_var).grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="Max Abs Delta / Underlying").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.portfolio_max_abs_delta_underlying_var = tk.StringVar(value="")
        ttk.Entry(strategy_frame, textvariable=self.portfolio_max_abs_delta_underlying_var).grid(row=row, column=1, sticky="ew", padx=8, pady=6)

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
            text="Walk-forward tunes on train+validation, then evaluates only on out-of-sample test folds. Optimizer runs constrained multi-objective search with early stopping.",
            wraplength=520,
            justify="left",
        ).grid(row=row, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 6))

        row += 1
        self.walk_forward_frame = ttk.LabelFrame(strategy_frame, text="Walk-Forward Windows (fractions of data)")
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

        ttk.Label(self.walk_forward_frame, text="CV Scheme").grid(row=4, column=0, sticky="w", padx=8, pady=6)
        self.wf_cv_scheme_var = tk.StringVar(value="walk_forward")
        self.wf_cv_scheme_combo = ttk.Combobox(self.walk_forward_frame, textvariable=self.wf_cv_scheme_var, state="readonly", values=["walk_forward", "cpcv"])
        self.wf_cv_scheme_combo.grid(row=4, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(self.walk_forward_frame, text="Purge/Embargo Bars").grid(row=5, column=0, sticky="w", padx=8, pady=6)
        pe_row = ttk.Frame(self.walk_forward_frame)
        pe_row.grid(row=5, column=1, sticky="ew", padx=8, pady=6)
        self.wf_purge_bars_var = tk.StringVar(value="0")
        self.wf_embargo_bars_var = tk.StringVar(value="0")
        ttk.Entry(pe_row, textvariable=self.wf_purge_bars_var, width=8).pack(side="left")
        ttk.Label(pe_row, text=" / ").pack(side="left")
        ttk.Entry(pe_row, textvariable=self.wf_embargo_bars_var, width=8).pack(side="left")

        ttk.Label(self.walk_forward_frame, text="CPCV Groups/Test Groups").grid(row=6, column=0, sticky="w", padx=8, pady=6)
        cpcv_row = ttk.Frame(self.walk_forward_frame)
        cpcv_row.grid(row=6, column=1, sticky="ew", padx=8, pady=6)
        self.wf_cpcv_groups_var = tk.StringVar(value="6")
        self.wf_cpcv_test_groups_var = tk.StringVar(value="2")
        ttk.Entry(cpcv_row, textvariable=self.wf_cpcv_groups_var, width=8).pack(side="left")
        ttk.Label(cpcv_row, text=" / ").pack(side="left")
        ttk.Entry(cpcv_row, textvariable=self.wf_cpcv_test_groups_var, width=8).pack(side="left")

        ttk.Label(self.walk_forward_frame, text="CV Seed").grid(row=7, column=0, sticky="w", padx=8, pady=6)
        self.wf_cv_seed_var = tk.StringVar(value="42")
        ttk.Entry(self.walk_forward_frame, textvariable=self.wf_cv_seed_var).grid(row=7, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        self.strategy_specific_container = ttk.Frame(strategy_frame)
        self.strategy_specific_container.grid(row=row, column=0, columnspan=2, sticky="ew", padx=0, pady=0)
        self.strategy_specific_container.columnconfigure(0, weight=1)

        self._build_momentum_options(self.strategy_specific_container)
        self._build_xsmom_options(self.strategy_specific_container)

        row += 1
        ttk.Label(strategy_frame, text="Start Date (YYYY-MM-DD)").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.start_date_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.start_date_var).grid(row=row, column=1, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(strategy_frame, text="End Date (YYYY-MM-DD)").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.end_date_var = tk.StringVar()
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
        gate_row = ttk.Frame(governance_frame)
        gate_row.grid(row=g_row, column=0, columnspan=2, sticky="ew", padx=8, pady=4)
        for idx, (label, default) in enumerate([
            ("Min OOS Periods", "3"),
            ("Min Stability", "0.55"),
            ("Max Turnover", "4.0"),
            ("Min Capacity", "0.5"),
        ]):
            ttk.Label(gate_row, text=label).grid(row=0, column=idx * 2, sticky="w", padx=(0, 4))
            var = tk.StringVar(value=default)
            entry = ttk.Entry(gate_row, textvariable=var, width=8)
            entry.grid(row=0, column=idx * 2 + 1, sticky="w", padx=(0, 10))
            if label == "Min OOS Periods":
                self.gov_min_oos_periods_var = var
            elif label == "Min Stability":
                self.gov_min_stability_var = var
            elif label == "Max Turnover":
                self.gov_max_turnover_var = var
            else:
                self.gov_min_capacity_var = var

        notes_frame = ttk.LabelFrame(run_setup_tab, text="Run Notes")
        notes_frame.grid(row=5, column=0, sticky="nsew", padx=10, pady=10)
        notes_frame.columnconfigure(0, weight=1)
        self.notes_text = tk.Text(notes_frame, height=14)
        self.notes_text.grid(row=0, column=0, sticky="nsew", padx=8, pady=6)

        button_row = ttk.Frame(run_setup_tab)
        button_row.grid(row=6, column=0, sticky="ew", padx=10, pady=(4, 10))
        button_row.columnconfigure(0, weight=1)
        button_row.columnconfigure(1, weight=1)
        button_row.columnconfigure(2, weight=1)

        ttk.Button(button_row, text="Save Parameters", command=self.save_settings).grid(row=0, column=0, padx=10, pady=4)
        self.run_button = ttk.Button(button_row, text="Run Backtest", command=self.run_backtest)
        self.run_button.grid(row=0, column=1, padx=10, pady=4)
        ttk.Button(
            button_row,
            text="Back to Main Menu",
            command=lambda: controller.show_frame("MainMenu"),
        ).grid(row=0, column=2, padx=10, pady=4)

        self._register_advanced_widgets()
        self._bind_validation_watchers()
        self._build_results_tabs()
        self._on_strategy_changed()
        self._on_mode_changed()
        self._refresh_template_choices()
        self._update_validation_hint()

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

        browser_tab = ttk.Frame(self.section_notebook)
        browser_tab.columnconfigure(0, weight=1)
        browser_tab.rowconfigure(1, weight=1)
        self.section_notebook.add(browser_tab, text="Experiment Browser")
        ttk.Button(browser_tab, text="Refresh Browser", command=self._refresh_experiment_browser).grid(row=0, column=0, sticky="w", padx=10, pady=(8, 4))
        ttk.Button(browser_tab, text="Export Notebook Bundle", command=self._export_selected_review_packet).grid(row=0, column=0, sticky="e", padx=10, pady=(8, 4))
        self.experiment_tree = ttk.Treeview(browser_tab, show="headings", height=10)
        self.experiment_tree.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 4))
        self.experiment_tree.bind("<<TreeviewSelect>>", lambda _e: self._on_experiment_tree_selected())
        self.experiment_detail_var = tk.StringVar(value="Select an experiment run to inspect tags, metrics, and reproducibility.")
        ttk.Label(browser_tab, textvariable=self.experiment_detail_var, justify="left").grid(row=2, column=0, sticky="ew", padx=10, pady=(2, 8))

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
        cmp_pane.add(metrics_frame, weight=1)
        cmp_pane.add(params_frame, weight=1)
        cmp_pane.add(variance_frame, weight=1)
        cmp_pane.add(scenario_frame, weight=1)
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
            self._render_guardrails(selected[0])
            return
        if len(selected) < 2:
            self.delta_summary_var.set("Select at least two runs for delta metrics.")
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
        self._update_equity_overlap(selected)
        self._populate_run_compare_combos(selected)
        self._render_selected_run_comparison()
        self._render_guardrails(selected[0])

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
        self._render_guardrails(run_dir)
        self._populate_run_compare_combos()
        self._refresh_experiment_browser()

    def _load_historical_runs(self) -> None:
        run_dirs = self._scan_backtest_output_runs()
        if not run_dirs:
            return
        self._register_run_dirs(run_dirs)
        self._refresh_artifacts_view()
        self._render_single_run(self.current_run_dirs[0])

    def _scan_backtest_output_runs(self) -> list[Path]:
        if not BACKTEST_OUTPUT_DIR.exists():
            return []
        candidates: list[Path] = []
        for path in BACKTEST_OUTPUT_DIR.iterdir():
            if not path.is_dir():
                continue
            if (path / "leaderboard.json").exists() or (path / "metrics.json").exists() or (path / "aggregate_metrics.json").exists() or (path / "manifest.json").exists() or (path / "fold_summary.json").exists():
                candidates.append(path)
        candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        return candidates

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
        for row in reversed(rows[-500:]):
            run_dir = Path(str(row.get("run_dir", "")))
            manifest = self._read_json(run_dir / "manifest.json")
            metrics = self._load_metric_map(run_dir)
            tags = parse_tags(manifest if isinstance(manifest, dict) else None)
            rendered.append(
                {
                    "timestamp": row.get("timestamp", ""),
                    "run_type": row.get("run_type", ""),
                    "run": run_dir.name,
                    "tags": ", ".join(tags[:4]),
                    "best_sharpe": metrics.get("sharpe", row.get("primary_metric_value", "")),
                    "approval": str((row.get("governance") or {}).get("approval_status", "")),
                    "fingerprint": str(row.get("reproducibility_fingerprint", ""))[:12],
                }
            )
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
        msg = [
            f"Run: {run_dir.name}",
            f"Tags: {', '.join(tags) if tags else '-'}",
            f"Sharpe: {metrics.get('sharpe', 'n/a')}",
            f"CAGR: {metrics.get('cagr', 'n/a')}",
            f"Fingerprint: {fp[:24] if fp else 'n/a'}",
        ]
        self.experiment_detail_var.set("\n".join(msg))

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
        }
        badges = build_guardrails(metrics, fold_rows=rows, trade_count=trade_count, robustness=robustness_payload, evidence_links=evidence_links)
        scenario_payload = read_stress_scenarios(run_dir)
        scenario_checks = scenario_payload.get("scenario_guardrails", []) if isinstance(scenario_payload, dict) else []
        if isinstance(scenario_checks, list):
            failed = [row for row in scenario_checks if isinstance(row, dict) and not bool(row.get("passed", False))]
            if failed:
                badges.append({"label": "Stress Failures", "severity": "high", "reason": f"{len(failed)} scenario guardrail checks failed."})
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
        ):
            var.trace_add("write", lambda *_args: self._update_validation_hint())

    def _on_mode_changed(self) -> None:
        is_advanced = self.ui_mode_var.get() == "advanced"
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

    def _update_validation_hint(self) -> bool:
        messages: list[str] = []
        disable_run = False
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
        if bool(self.use_optimizer_var.get()) and self.ui_mode_var.get() != "advanced":
            messages.append("Optimizer requires Advanced mode.")
            disable_run = True

        self._validation_messages = messages
        self.validation_hint_var.set("\n".join(messages))
        self.run_button.config(state="disabled" if disable_run else "normal")
        return disable_run

    def _refresh_template_choices(self) -> None:
        names = sorted(self.controller.state.backtest_templates.keys())
        self.template_combo.configure(values=names)
        selected = self.template_var.get().strip()
        if selected and selected not in names:
            self.template_var.set("")

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
        settings = dict(DEFAULT_BACKTEST_SETTINGS)
        settings.update(self.controller.state.backtest_settings)

        strategy = str(settings.get("strategy", "momentum"))
        if strategy not in STRATEGIES:
            strategy = "momentum"
        self.strategy_var.set(strategy)
        self.ui_mode_var.set(str(settings.get("ui_mode", "basic")))
        selected_preset = str(settings.get("selected_preset", "custom"))
        if selected_preset not in {"custom", *BACKTEST_STRATEGY_PRESETS.keys()}:
            selected_preset = "custom"
        self.preset_var.set(self._preset_key_to_display.get(selected_preset, "Custom"))

        self.lookback_days_var.set(str(settings.get("lookback_days", "90")))
        self.skip_days_var.set(str(settings.get("skip_days", "5")))
        self.costs_bps_var.set(str(settings.get("costs_bps", "5")))
        self.execution_model_var.set(str(settings.get("execution_model", "bps")))
        self.execution_spread_bps_var.set(str(settings.get("execution_spread_bps", "2")))
        self.execution_max_participation_var.set(str(settings.get("execution_max_participation", "1.0")))
        self.execution_impact_bps_var.set(str(settings.get("execution_impact_bps", "5")))
        self.execution_latency_bars_var.set(str(settings.get("execution_latency_bars", "0")))
        self.execution_latency_ms_var.set(str(settings.get("execution_latency_ms", "0")))
        self.starting_capital_var.set(str(settings.get("starting_capital", "100000")))
        self.bet_sizing_mode_var.set(str(settings.get("bet_sizing_mode", "half_kelly")))
        self.custom_bet_pct_var.set(str(settings.get("custom_bet_pct", "10")))
        timeframe = str(settings.get("timeframe", "1m"))
        self.timeframe_var.set(timeframe if timeframe in TIMEFRAMES else "1m")
        self.use_walk_forward_var.set(bool(settings.get("use_walk_forward", False)))
        self.use_optimizer_var.set(bool(settings.get("use_optimizer", False)))
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
        self.wf_train_fraction_var.set(float(settings.get("wf_train_fraction", "0.70")))
        self.wf_validation_fraction_var.set(float(settings.get("wf_validation_fraction", "0.15")))
        self.wf_test_fraction_var.set(float(settings.get("wf_test_fraction", "0.15")))
        self.wf_step_fraction_var.set(float(settings.get("wf_step_fraction", "0.15")))
        self.wf_cv_scheme_var.set(str(settings.get("wf_cv_scheme", "walk_forward")))
        self.wf_purge_bars_var.set(str(settings.get("wf_purge_window_bars", "0")))
        self.wf_embargo_bars_var.set(str(settings.get("wf_embargo_window_bars", "0")))
        self.wf_cpcv_groups_var.set(str(settings.get("wf_cpcv_n_groups", "6")))
        self.wf_cpcv_test_groups_var.set(str(settings.get("wf_cpcv_n_test_groups", "2")))
        self.wf_cv_seed_var.set(str(settings.get("wf_cv_seed", "42")))
        self._refresh_wf_fraction_labels()

        selected_entries = self._split_csv_setting(settings.get("selected_entry_signals", "ts_momentum"))
        selected_exits = self._split_csv_setting(settings.get("selected_exit_signals", "none"))
        for name, var in self.entry_signal_vars.items():
            var.set(name in selected_entries)
        for name, var in self.exit_signal_vars.items():
            var.set(name in selected_exits)

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
        self.logs_text.delete("1.0", tk.END)
        self.logs_text.insert("1.0", str(settings.get("notes", "")))
        self._refresh_template_choices()
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
            "starting_capital": str(starting_capital),
            "bet_sizing_mode": self.bet_sizing_mode_var.get().strip() or "half_kelly",
            "custom_bet_pct": str(custom_bet_pct),
            "timeframe": self.timeframe_var.get().strip() or "1m",
            "use_walk_forward": bool(self.use_walk_forward_var.get()),
            "use_optimizer": bool(self.use_optimizer_var.get()),
            "wf_train_fraction": f"{float(self.wf_train_fraction_var.get()):.2f}",
            "wf_validation_fraction": f"{float(self.wf_validation_fraction_var.get()):.2f}",
            "wf_test_fraction": f"{float(self.wf_test_fraction_var.get()):.2f}",
            "wf_step_fraction": f"{float(self.wf_step_fraction_var.get()):.2f}",
            "wf_cv_scheme": self.wf_cv_scheme_var.get().strip() or "walk_forward",
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
            "selected_entry_signals": ",".join(selected_entries),
            "selected_exit_signals": ",".join(selected_exits),
            "start_date": self.start_date_var.get().strip(),
            "end_date": self.end_date_var.get().strip(),
            "backtest_data_root": self.backtest_root_var.get().strip(),
            "notes": self.notes_text.get("1.0", tk.END).strip(),
            "governance_hypothesis_id": self.gov_hypothesis_id_var.get().strip(),
            "governance_owner": self.gov_owner_var.get().strip(),
            "governance_dataset_snapshot_lock": self.gov_dataset_lock_var.get().strip(),
            "governance_acceptance_criteria": self.gov_acceptance_text.get("1.0", tk.END).strip(),
            "governance_promotion_state": self.gov_promotion_state_var.get().strip() or "research",
            "governance_approval_status": self.gov_approval_status_var.get().strip() or "pending",
            "governance_min_oos_periods": self.gov_min_oos_periods_var.get().strip() or "3",
            "governance_min_stability_score": self.gov_min_stability_var.get().strip() or "0.55",
            "governance_max_turnover_total": self.gov_max_turnover_var.get().strip() or "4.0",
            "governance_min_capacity_score": self.gov_min_capacity_var.get().strip() or "0.5",
            "ui_mode": self.ui_mode_var.get().strip() or "basic",
            "selected_preset": self._preset_display_to_key.get(self.preset_var.get().strip(), "custom"),
            "selected_template": self.template_var.get().strip(),
            **xsmom_params,
        }
        self.controller.persist_state()
        self._refresh_template_choices()
        if show_confirmation:
            messagebox.showinfo("Saved", "Backtesting parameters saved.")
        return True

    def run_backtest(self) -> None:
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
        }
        if portfolio_cfg["portfolio_method"] not in PORTFOLIO_METHODS:
            messagebox.showinfo("Invalid input", "Please select a valid portfolio method.")
            return

        execution_model = self.execution_model_var.get().strip() or "bps"
        if execution_model not in EXECUTION_MODELS:
            messagebox.showinfo("Invalid input", "Please select a valid execution model.")
            return
        execution_model_params = {
            "spread_bps": float(parse_float(self.execution_spread_bps_var.get()) or 2.0),
            "max_participation": float(parse_float(self.execution_max_participation_var.get()) or 1.0),
            "impact_bps": float(parse_float(self.execution_impact_bps_var.get()) or costs_bps),
            "latency_bars": int(parse_float(self.execution_latency_bars_var.get()) or 0),
            "latency_ms": int(parse_float(self.execution_latency_ms_var.get()) or 0),
            "drift_bps_per_bar": float(parse_float(self.execution_impact_bps_var.get()) or 1.0),
        }

        governance_payload = {
            "hypothesis_id": self.gov_hypothesis_id_var.get().strip(),
            "owner": self.gov_owner_var.get().strip(),
            "dataset_snapshot_lock": self.gov_dataset_lock_var.get().strip(),
            "acceptance_criteria": self.gov_acceptance_text.get("1.0", tk.END).strip(),
            "approval_status": self.gov_approval_status_var.get().strip() or "pending",
            "promotion_state": self.gov_promotion_state_var.get().strip() or "research",
            "min_oos_periods": int(parse_float(self.gov_min_oos_periods_var.get()) or 3),
            "min_stability_score": float(parse_float(self.gov_min_stability_var.get()) or 0.55),
            "max_turnover_total": float(parse_float(self.gov_max_turnover_var.get()) or 4.0),
            "min_capacity_score": float(parse_float(self.gov_min_capacity_var.get()) or 0.5),
        }

        worker_args: tuple[object, ...]
        status_line: str

        if strategy == "momentum":
            selected_entries = self._selected_signal_names(self.entry_signal_vars)
            selected_exits = self._selected_signal_names(self.exit_signal_vars)
            if not selected_entries:
                messagebox.showinfo("Invalid input", "Select at least one entry signal.")
                return
            if not selected_exits:
                messagebox.showinfo("Invalid input", "Select at least one exit signal.")
                return

            if bool(self.use_optimizer_var.get()):
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
                )
                status_line = f"Running optimizer across {len(selected_entries) * len(selected_exits)} candidates...\n"
            elif bool(self.use_walk_forward_var.get()):
                walk_forward_windows = self._validate_walk_forward_inputs()
                if walk_forward_windows is None:
                    return
                train_fraction, validation_fraction, test_fraction, step_fraction, cv_scheme, purge_bars, embargo_bars, cpcv_groups, cpcv_test_groups, cv_seed = walk_forward_windows
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
                    train_fraction,
                    validation_fraction,
                    test_fraction,
                    step_fraction,
                    cv_scheme,
                    purge_bars,
                    embargo_bars,
                    cpcv_groups,
                    cpcv_test_groups,
                    cv_seed,
                    governance_payload,
                )
                status_line = f"Running walk-forward with {len(selected_entries) * len(selected_exits)} candidates...\n"
            else:
                worker_target = self._run_momentum_worker
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
                    selected_entries,
                    selected_exits,
                    execution_model,
                    execution_model_params,
                    portfolio_cfg,
                    governance_payload,
                )
                status_line = f"Running {len(selected_entries) * len(selected_exits)} momentum entry/exit combinations...\n"
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
            )
            status_line = "Running cross-sectional momentum backtest...\n"

        self.run_button.config(state="disabled")
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

    def _validate_walk_forward_inputs(self) -> tuple[float, float, float, float, str, int, int, int, int, int] | None:
        train_fraction = float(self.wf_train_fraction_var.get())
        validation_fraction = float(self.wf_validation_fraction_var.get())
        test_fraction = float(self.wf_test_fraction_var.get())
        step_fraction = float(self.wf_step_fraction_var.get())

        values = [train_fraction, validation_fraction, test_fraction]
        if any(value <= 0.0 or value >= 1.0 for value in values):
            messagebox.showinfo("Invalid input", "Train/validation/test fractions must be in (0, 1).")
            return None
        if abs(sum(values) - 1.0) > 1e-6:
            messagebox.showinfo("Invalid input", "Train, validation, and test fractions must sum to 1.0.")
            return None
        if step_fraction <= 0.0 or step_fraction > 1.0:
            messagebox.showinfo("Invalid input", "Step fraction must be in (0, 1].")
            return None

        cv_scheme = self.wf_cv_scheme_var.get().strip() or "walk_forward"
        purge = parse_float(self.wf_purge_bars_var.get())
        embargo = parse_float(self.wf_embargo_bars_var.get())
        n_groups = parse_float(self.wf_cpcv_groups_var.get())
        n_test_groups = parse_float(self.wf_cpcv_test_groups_var.get())
        cv_seed = parse_float(self.wf_cv_seed_var.get())
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

        return train_fraction, validation_fraction, test_fraction, step_fraction, cv_scheme, int(purge), int(embargo), int(n_groups), int(n_test_groups), int(cv_seed)

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
    ) -> None:
        try:
            entry_grid = {signal: [{}] for signal in entry_signals}
            exit_grid = {signal: [{}] for signal in exit_signals}
            core_grid = {
                "lookback_days": [int(lookback)],
                "skip_days": [int(skip)],
                "costs_bps": [float(costs_bps)],
            }
            output_text = run_strategy_optimization(
                tickers=tickers,
                start_date=start_date,
                end_date=end_date,
                cache_root=cache_root,
                entry_grid=entry_grid,
                exit_grid=exit_grid,
                core_grid=core_grid,
                seed=42,
                n_trials=max(10, len(entry_signals) * len(exit_signals) * 4),
                sampler_name="tpe",
                partial_period_fractions=[0.33, 0.66, 1.0],
                governance_metadata=dict(governance_payload),
            )
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
        train_fraction: float,
        validation_fraction: float,
        test_fraction: float,
        step_fraction: float,
        cv_scheme: str,
        purge_window_bars: int,
        embargo_window_bars: int,
        cpcv_n_groups: int,
        cpcv_n_test_groups: int,
        cv_seed: int,
        governance_payload: dict[str, object],
    ) -> None:
        try:
            entry_grid = {signal: [{}] for signal in entry_signals}
            exit_grid = {signal: [{}] for signal in exit_signals}
            core_grid = {
                "lookback_days": [int(lookback)],
                "skip_days": [int(skip)],
                "costs_bps": [float(costs_bps)],
            }
            output_text = run_walk_forward_backtest(
                tickers=tickers,
                start_date=start_date,
                end_date=end_date,
                cache_root=cache_root,
                entry_grid=entry_grid,
                exit_grid=exit_grid,
                core_grid=core_grid,
                train_fraction=train_fraction,
                validation_fraction=validation_fraction,
                test_fraction=test_fraction,
                step_fraction=step_fraction,
                cv_scheme=cv_scheme,
                purge_window_bars=int(purge_window_bars),
                embargo_window_bars=int(embargo_window_bars),
                cpcv_n_groups=int(cpcv_n_groups),
                cpcv_n_test_groups=int(cpcv_n_test_groups),
                cv_seed=int(cv_seed),
                governance_metadata=dict(governance_payload),
            )
        except Exception as exc:
            output_text = f"Backtest failed: {exc}"
        self.after(0, lambda: self._finish_backtest_run(output_text))

    def _run_momentum_worker(
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
        entry_signals: list[str],
        exit_signals: list[str],
        execution_model: str,
        execution_model_params: dict[str, object],
        portfolio_cfg: dict[str, object],
        governance_payload: dict[str, object],
    ) -> None:
        try:
            output_text = run_multi_signal_backtest(
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
                timeframe=timeframe,
                entry_signals=entry_signals,
                exit_signals=exit_signals,
                **portfolio_cfg,
                governance_metadata=dict(governance_payload),
            )
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
    ) -> None:
        try:
            output_text = run_time_series_momentum_backtest(
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
            )
        except Exception as exc:
            output_text = f"Backtest failed: {exc}"
        self.after(0, lambda: self._finish_backtest_run(output_text))

    def _finish_backtest_run(self, output_text: str) -> None:
        self._consume_run_outputs(output_text)
        self.run_button.config(state="normal")

    def _split_csv_setting(self, raw: object) -> set[str]:
        return {part.strip() for part in str(raw).split(",") if part.strip()}

    def _selected_signal_names(self, signals: dict[str, tk.BooleanVar]) -> list[str]:
        return [name for name, var in signals.items() if bool(var.get())]

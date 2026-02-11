from __future__ import annotations

import threading
import csv
import json
from datetime import date
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

from backtesting.cache_runner import (
    run_multi_signal_backtest,
    run_time_series_momentum_backtest,
    run_walk_forward_backtest,
)
from config import BACKTEST_CACHE_DIR, BACKTEST_OUTPUT_DIR, BACKTEST_STRATEGY_PRESETS, DEFAULT_BACKTEST_SETTINGS
from utils.parsing import normalize_cache_root, parse_date, parse_float

ENTRY_SIGNALS = ["ts_momentum", "ma_trend", "breakout"]
EXIT_SIGNALS = ["none", "momentum_flip", "trailing_stop", "max_hold"]
STRATEGIES = ["momentum", "xsmom"]
TIMEFRAMES = ["1m", "5m", "15m", "30m", "1h", "1d"]
PORTFOLIO_METHODS = ["equal_weight", "vol_target", "inverse_vol", "capped_optimization"]

TIMEFRAME_MIN_LOOKBACK = {"1m": 120, "5m": 80, "15m": 60, "30m": 40, "1h": 30, "1d": 20}
TIMEFRAME_HISTORY_DAYS = {"1m": 14, "5m": 30, "15m": 60, "30m": 120, "1h": 365, "1d": 3650}

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
        run_setup_tab.rowconfigure(1, weight=1)
        self.section_notebook.add(run_setup_tab, text="Run Setup")

        strategy_frame = ttk.LabelFrame(run_setup_tab, text="Strategy")
        strategy_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        strategy_frame.columnconfigure(1, weight=1)

        row = 0
        ttk.Label(strategy_frame, text="Mode").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.ui_mode_var = tk.StringVar(value="basic")
        mode_row = ttk.Frame(strategy_frame)
        mode_row.grid(row=row, column=1, sticky="w", padx=8, pady=6)
        ttk.Radiobutton(mode_row, text="Basic", value="basic", variable=self.ui_mode_var, command=self._on_mode_changed).pack(side="left")
        ttk.Radiobutton(mode_row, text="Advanced", value="advanced", variable=self.ui_mode_var, command=self._on_mode_changed).pack(side="left", padx=(8, 0))

        row += 1
        ttk.Label(strategy_frame, text="Preset").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.preset_var = tk.StringVar(value="custom")
        preset_values = ["custom"] + list(BACKTEST_STRATEGY_PRESETS.keys())
        self.preset_combo = ttk.Combobox(
            strategy_frame,
            textvariable=self.preset_var,
            state="readonly",
            values=preset_values,
        )
        self.preset_combo.grid(row=row, column=1, sticky="ew", padx=8, pady=6)
        self.preset_combo.bind("<<ComboboxSelected>>", self._on_preset_selected)

        row += 1
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
        self.use_walk_forward_var = tk.BooleanVar(value=False)
        self.use_walk_forward_check = ttk.Checkbutton(
            strategy_frame,
            text="Use Walk-Forward (Momentum)",
            variable=self.use_walk_forward_var,
            command=self._update_validation_hint,
        )
        self.use_walk_forward_check.grid(row=row, column=0, columnspan=2, sticky="w", padx=8, pady=6)

        row += 1
        ttk.Label(
            strategy_frame,
            text="Walk-forward tunes on train+validation, then evaluates only on out-of-sample test folds.",
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

        notes_frame = ttk.LabelFrame(run_setup_tab, text="Run Notes")
        notes_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        notes_frame.columnconfigure(0, weight=1)
        self.notes_text = tk.Text(notes_frame, height=14)
        self.notes_text.grid(row=0, column=0, sticky="nsew", padx=8, pady=6)

        button_row = ttk.Frame(run_setup_tab)
        button_row.grid(row=2, column=0, sticky="ew", padx=10, pady=(4, 10))
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

        leaderboard_tab = ttk.Frame(self.section_notebook)
        leaderboard_tab.columnconfigure(0, weight=1)
        leaderboard_tab.rowconfigure(2, weight=1)
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
        trades_tab.rowconfigure(0, weight=1)
        self.section_notebook.add(trades_tab, text="Trades/Costs diagnostics")
        self.trades_tree = ttk.Treeview(trades_tab, show="headings", height=14)
        self.trades_tree.grid(row=0, column=0, sticky="nsew", padx=10, pady=8)

        wf_tab = ttk.Frame(self.section_notebook)
        wf_tab.columnconfigure(0, weight=1)
        wf_tab.rowconfigure(1, weight=1)
        self.section_notebook.add(wf_tab, text="Fold/WF diagnostics")
        self.wf_status_var = tk.StringVar(value="No walk-forward diagnostics loaded.")
        ttk.Label(wf_tab, textvariable=self.wf_status_var, justify="left").grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 4))
        self.wf_tree = ttk.Treeview(wf_tab, show="headings", height=14)
        self.wf_tree.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 8))

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
            self._set_tree_data(self.wf_tree, wf_rows)
            self.wf_status_var.set(f"Loaded {len(wf_rows)} walk-forward folds from {run_dir.name}.")
        else:
            self._set_tree_data(self.wf_tree, [])
            self.wf_status_var.set("No walk-forward diagnostics loaded.")

        self._update_equity_overlap([run_dir])

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
            "portfolio_max_gross": self.portfolio_max_gross_entry,
            "portfolio_min_net": self.portfolio_min_net_entry,
            "portfolio_max_net": self.portfolio_max_net_entry,
            "use_walk_forward": self.use_walk_forward_check,
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
        lookback = parse_float(self.lookback_days_var.get())
        start_date = parse_date(self.start_date_var.get())
        end_date = parse_date(self.end_date_var.get())

        if timeframe in TIMEFRAME_MIN_LOOKBACK and lookback is not None:
            min_lookback = TIMEFRAME_MIN_LOOKBACK[timeframe]
            if int(lookback) < min_lookback:
                messages.append(f"{timeframe} usually needs lookback >= {min_lookback} bars for stable signals.")
                disable_run = True

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
        preset_key = self.preset_var.get().strip()
        if not preset_key or preset_key == "custom":
            return
        preset = BACKTEST_STRATEGY_PRESETS.get(preset_key)
        if not preset:
            return
        preset_settings = preset.get("settings", {})
        if isinstance(preset_settings, dict):
            self._apply_settings(preset_settings)
            self.preset_var.set(preset_key)
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
        self.preset_var.set(selected_preset)

        self.lookback_days_var.set(str(settings.get("lookback_days", "90")))
        self.skip_days_var.set(str(settings.get("skip_days", "5")))
        self.costs_bps_var.set(str(settings.get("costs_bps", "5")))
        self.starting_capital_var.set(str(settings.get("starting_capital", "100000")))
        self.bet_sizing_mode_var.set(str(settings.get("bet_sizing_mode", "half_kelly")))
        self.custom_bet_pct_var.set(str(settings.get("custom_bet_pct", "10")))
        timeframe = str(settings.get("timeframe", "1m"))
        self.timeframe_var.set(timeframe if timeframe in TIMEFRAMES else "1m")
        self.use_walk_forward_var.set(bool(settings.get("use_walk_forward", False)))
        self.portfolio_method_var.set(str(settings.get("portfolio_method", "equal_weight")))
        self.portfolio_vol_lookback_var.set(str(settings.get("portfolio_vol_lookback_bars", "20")))
        self.portfolio_target_vol_var.set(str(settings.get("portfolio_target_volatility", "0.10")))
        self.portfolio_max_symbol_var.set(str(settings.get("portfolio_max_symbol_weight", "0.25")))
        self.portfolio_max_sector_var.set(str(settings.get("portfolio_max_sector_weight", "0.60")))
        self.portfolio_max_gross_var.set(str(settings.get("portfolio_max_gross_exposure", "1.0")))
        self.portfolio_min_net_var.set(str(settings.get("portfolio_min_net_exposure", "-1.0")))
        self.portfolio_max_net_var.set(str(settings.get("portfolio_max_net_exposure", "1.0")))
        self.wf_train_fraction_var.set(float(settings.get("wf_train_fraction", "0.70")))
        self.wf_validation_fraction_var.set(float(settings.get("wf_validation_fraction", "0.15")))
        self.wf_test_fraction_var.set(float(settings.get("wf_test_fraction", "0.15")))
        self.wf_step_fraction_var.set(float(settings.get("wf_step_fraction", "0.15")))
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
            "starting_capital": str(starting_capital),
            "bet_sizing_mode": self.bet_sizing_mode_var.get().strip() or "half_kelly",
            "custom_bet_pct": str(custom_bet_pct),
            "timeframe": self.timeframe_var.get().strip() or "1m",
            "use_walk_forward": bool(self.use_walk_forward_var.get()),
            "wf_train_fraction": f"{float(self.wf_train_fraction_var.get()):.2f}",
            "wf_validation_fraction": f"{float(self.wf_validation_fraction_var.get()):.2f}",
            "wf_test_fraction": f"{float(self.wf_test_fraction_var.get()):.2f}",
            "wf_step_fraction": f"{float(self.wf_step_fraction_var.get()):.2f}",
            "portfolio_method": self.portfolio_method_var.get().strip() or "equal_weight",
            "portfolio_vol_lookback_bars": self.portfolio_vol_lookback_var.get().strip() or "20",
            "portfolio_target_volatility": self.portfolio_target_vol_var.get().strip() or "0.10",
            "portfolio_max_symbol_weight": self.portfolio_max_symbol_var.get().strip() or "0.25",
            "portfolio_max_sector_weight": self.portfolio_max_sector_var.get().strip() or "0.60",
            "portfolio_max_gross_exposure": self.portfolio_max_gross_var.get().strip() or "1.0",
            "portfolio_min_net_exposure": self.portfolio_min_net_var.get().strip() or "-1.0",
            "portfolio_max_net_exposure": self.portfolio_max_net_var.get().strip() or "1.0",
            "selected_entry_signals": ",".join(selected_entries),
            "selected_exit_signals": ",".join(selected_exits),
            "start_date": self.start_date_var.get().strip(),
            "end_date": self.end_date_var.get().strip(),
            "backtest_data_root": self.backtest_root_var.get().strip(),
            "notes": self.notes_text.get("1.0", tk.END).strip(),
            "ui_mode": self.ui_mode_var.get().strip() or "basic",
            "selected_preset": self.preset_var.get().strip() or "custom",
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
        parsed_max_gross = parse_float(self.portfolio_max_gross_var.get())
        parsed_min_net = parse_float(self.portfolio_min_net_var.get())
        parsed_max_net = parse_float(self.portfolio_max_net_var.get())

        portfolio_cfg = {
            "portfolio_method": self.portfolio_method_var.get().strip() or "equal_weight",
            "portfolio_vol_lookback_bars": int(parsed_vol_lookback) if parsed_vol_lookback is not None else 20,
            "portfolio_target_volatility": float(parsed_target_vol) if parsed_target_vol is not None else 0.10,
            "portfolio_max_symbol_weight": float(parsed_max_symbol) if parsed_max_symbol is not None else 0.25,
            "portfolio_max_sector_weight": float(parsed_max_sector) if parsed_max_sector is not None else 0.60,
            "portfolio_max_gross_exposure": float(parsed_max_gross) if parsed_max_gross is not None else 1.0,
            "portfolio_min_net_exposure": float(parsed_min_net) if parsed_min_net is not None else -1.0,
            "portfolio_max_net_exposure": float(parsed_max_net) if parsed_max_net is not None else 1.0,
        }
        if portfolio_cfg["portfolio_method"] not in PORTFOLIO_METHODS:
            messagebox.showinfo("Invalid input", "Please select a valid portfolio method.")
            return

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

            if bool(self.use_walk_forward_var.get()):
                walk_forward_windows = self._validate_walk_forward_inputs()
                if walk_forward_windows is None:
                    return
                train_fraction, validation_fraction, test_fraction, step_fraction = walk_forward_windows
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
                    train_fraction,
                    validation_fraction,
                    test_fraction,
                    step_fraction,
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
                    portfolio_cfg,
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
                portfolio_cfg,
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
        min_lookback = TIMEFRAME_MIN_LOOKBACK.get(timeframe)
        if min_lookback is not None and int(lookback) < min_lookback:
            messagebox.showinfo("Invalid input", f"{timeframe} requires lookback >= {min_lookback} bars.")
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

    def _validate_walk_forward_inputs(self) -> tuple[float, float, float, float] | None:
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

        return train_fraction, validation_fraction, test_fraction, step_fraction

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
        train_fraction: float,
        validation_fraction: float,
        test_fraction: float,
        step_fraction: float,
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
        portfolio_cfg: dict[str, object],
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
                starting_capital=starting_capital,
                bet_sizing_mode=bet_sizing_mode,
                custom_bet_pct=custom_bet_pct,
                timeframe=timeframe,
                entry_signals=entry_signals,
                exit_signals=exit_signals,
                **portfolio_cfg,
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
        portfolio_cfg: dict[str, object],
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

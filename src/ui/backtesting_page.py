from __future__ import annotations

import threading
from datetime import date
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

from backtesting.cache_runner import (
    run_multi_signal_backtest,
    run_time_series_momentum_backtest,
    run_walk_forward_backtest,
)
from config import BACKTEST_CACHE_DIR, DEFAULT_BACKTEST_SETTINGS
from utils.parsing import normalize_cache_root, parse_date, parse_float

ENTRY_SIGNALS = ["ts_momentum", "ma_trend", "breakout"]
EXIT_SIGNALS = ["none", "momentum_flip", "trailing_stop", "max_hold"]
STRATEGIES = ["momentum", "xsmom"]
TIMEFRAMES = ["1m", "5m", "15m", "30m", "1h", "1d"]


class BacktestingPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller

        ttk.Label(self, text="Backtesting Parameters", font=("Arial", 18, "bold")).pack(pady=10)

        intro = (
            "Choose a strategy and configure its parameters. Shared settings stay visible, "
            "and strategy-specific controls appear only for the selected strategy."
        )
        ttk.Label(self, text=intro, wraplength=950, justify="center").pack(pady=5)

        content = ttk.Frame(self)
        content.pack(fill="both", expand=True, padx=30, pady=10)
        content.columnconfigure(0, weight=1)
        content.rowconfigure(1, weight=1)

        strategy_frame = ttk.LabelFrame(content, text="Strategy")
        strategy_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
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
        ttk.Entry(strategy_frame, textvariable=self.custom_bet_pct_var).grid(row=row, column=1, sticky="ew", padx=8, pady=6)

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
        self.use_walk_forward_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            strategy_frame,
            text="Use Walk-Forward (Momentum)",
            variable=self.use_walk_forward_var,
        ).grid(row=row, column=0, columnspan=2, sticky="w", padx=8, pady=6)

        row += 1
        ttk.Label(
            strategy_frame,
            text="Walk-forward tunes on train+validation, then evaluates only on out-of-sample test folds.",
            wraplength=520,
            justify="left",
        ).grid(row=row, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 6))

        row += 1
        self.walk_forward_frame = ttk.LabelFrame(strategy_frame, text="Walk-Forward Windows (bars)")
        self.walk_forward_frame.grid(row=row, column=0, columnspan=2, sticky="ew", padx=8, pady=6)
        self.walk_forward_frame.columnconfigure(1, weight=1)

        ttk.Label(self.walk_forward_frame, text="Train").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        self.wf_train_bars_var = tk.StringVar(value="3900")
        ttk.Entry(self.walk_forward_frame, textvariable=self.wf_train_bars_var).grid(row=0, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(self.walk_forward_frame, text="Validation").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        self.wf_validation_bars_var = tk.StringVar(value="780")
        ttk.Entry(self.walk_forward_frame, textvariable=self.wf_validation_bars_var).grid(row=1, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(self.walk_forward_frame, text="Test").grid(row=2, column=0, sticky="w", padx=8, pady=6)
        self.wf_test_bars_var = tk.StringVar(value="780")
        ttk.Entry(self.walk_forward_frame, textvariable=self.wf_test_bars_var).grid(row=2, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(self.walk_forward_frame, text="Step (blank = test size)").grid(row=3, column=0, sticky="w", padx=8, pady=6)
        self.wf_step_bars_var = tk.StringVar(value="")
        ttk.Entry(self.walk_forward_frame, textvariable=self.wf_step_bars_var).grid(row=3, column=1, sticky="ew", padx=8, pady=6)

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

        notes_frame = ttk.LabelFrame(content, text="Run Output")
        notes_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        notes_frame.columnconfigure(0, weight=1)
        self.notes_text = tk.Text(notes_frame, height=14)
        self.notes_text.grid(row=0, column=0, sticky="nsew", padx=8, pady=6)

        button_row = ttk.Frame(content)
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

        self._on_strategy_changed()

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

    def refresh(self) -> None:
        settings = dict(DEFAULT_BACKTEST_SETTINGS)
        settings.update(self.controller.state.backtest_settings)

        strategy = str(settings.get("strategy", "momentum"))
        if strategy not in STRATEGIES:
            strategy = "momentum"
        self.strategy_var.set(strategy)

        self.lookback_days_var.set(str(settings.get("lookback_days", "90")))
        self.skip_days_var.set(str(settings.get("skip_days", "5")))
        self.costs_bps_var.set(str(settings.get("costs_bps", "5")))
        self.starting_capital_var.set(str(settings.get("starting_capital", "100000")))
        self.bet_sizing_mode_var.set(str(settings.get("bet_sizing_mode", "half_kelly")))
        self.custom_bet_pct_var.set(str(settings.get("custom_bet_pct", "10")))
        timeframe = str(settings.get("timeframe", "1m"))
        self.timeframe_var.set(timeframe if timeframe in TIMEFRAMES else "1m")
        self.use_walk_forward_var.set(bool(settings.get("use_walk_forward", False)))
        self.wf_train_bars_var.set(str(settings.get("wf_train_bars", "3900")))
        self.wf_validation_bars_var.set(str(settings.get("wf_validation_bars", "780")))
        self.wf_test_bars_var.set(str(settings.get("wf_test_bars", "780")))
        self.wf_step_bars_var.set(str(settings.get("wf_step_bars", "")))

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

        self.notes_text.delete("1.0", tk.END)
        self.notes_text.insert("1.0", str(settings.get("notes", "")))
        self._on_strategy_changed()

    def save_settings(self) -> None:
        validated = self._validate_common_inputs()
        if validated is None:
            return

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
                return
            if not selected_exits:
                messagebox.showinfo("Invalid input", "Select at least one exit signal.")
                return
        else:
            xsmom_valid = self._validate_xsmom_inputs()
            if xsmom_valid is None:
                return
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
            "wf_train_bars": self.wf_train_bars_var.get().strip() or "3900",
            "wf_validation_bars": self.wf_validation_bars_var.get().strip() or "780",
            "wf_test_bars": self.wf_test_bars_var.get().strip() or "780",
            "wf_step_bars": self.wf_step_bars_var.get().strip(),
            "selected_entry_signals": ",".join(selected_entries),
            "selected_exit_signals": ",".join(selected_exits),
            "start_date": self.start_date_var.get().strip(),
            "end_date": self.end_date_var.get().strip(),
            "backtest_data_root": self.backtest_root_var.get().strip(),
            "notes": self.notes_text.get("1.0", tk.END).strip(),
            **xsmom_params,
        }
        self.controller.persist_state()
        messagebox.showinfo("Saved", "Backtesting parameters saved.")

    def run_backtest(self) -> None:
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
                train_bars, validation_bars, test_bars, step_bars = walk_forward_windows
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
                    train_bars,
                    validation_bars,
                    test_bars,
                    step_bars,
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
            )
            status_line = "Running cross-sectional momentum backtest...\n"

        self.run_button.config(state="disabled")
        self.notes_text.delete("1.0", tk.END)
        self.notes_text.insert("1.0", status_line)

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

    def _validate_walk_forward_inputs(self) -> tuple[int, int, int, int | None] | None:
        train_bars = parse_float(self.wf_train_bars_var.get())
        validation_bars = parse_float(self.wf_validation_bars_var.get())
        test_bars = parse_float(self.wf_test_bars_var.get())
        step_raw = self.wf_step_bars_var.get().strip()
        step_bars = parse_float(step_raw) if step_raw else None

        if train_bars is None or train_bars < 1 or int(train_bars) != train_bars:
            messagebox.showinfo("Invalid input", "Walk-forward train bars must be a positive integer.")
            return None
        if validation_bars is None or validation_bars < 1 or int(validation_bars) != validation_bars:
            messagebox.showinfo("Invalid input", "Walk-forward validation bars must be a positive integer.")
            return None
        if test_bars is None or test_bars < 1 or int(test_bars) != test_bars:
            messagebox.showinfo("Invalid input", "Walk-forward test bars must be a positive integer.")
            return None
        if step_bars is not None and (step_bars < 1 or int(step_bars) != step_bars):
            messagebox.showinfo("Invalid input", "Walk-forward step bars must be a positive integer when provided.")
            return None

        return int(train_bars), int(validation_bars), int(test_bars), None if step_bars is None else int(step_bars)

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
        train_bars: int,
        validation_bars: int,
        test_bars: int,
        step_bars: int | None,
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
                train_bars=train_bars,
                validation_bars=validation_bars,
                test_bars=test_bars,
                step_bars=step_bars,
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
            )
        except Exception as exc:
            output_text = f"Backtest failed: {exc}"
        self.after(0, lambda: self._finish_backtest_run(output_text))

    def _finish_backtest_run(self, output_text: str) -> None:
        self.notes_text.delete("1.0", tk.END)
        self.notes_text.insert("1.0", output_text)
        self.run_button.config(state="normal")

    def _split_csv_setting(self, raw: object) -> set[str]:
        return {part.strip() for part in str(raw).split(",") if part.strip()}

    def _selected_signal_names(self, signals: dict[str, tk.BooleanVar]) -> list[str]:
        return [name for name, var in signals.items() if bool(var.get())]

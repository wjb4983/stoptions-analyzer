from __future__ import annotations

import json
import threading
from datetime import date
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

from backtesting.cache_runner import run_parameter_sweep, run_time_series_momentum_backtest
from config import BACKTEST_CACHE_DIR, DEFAULT_BACKTEST_SETTINGS
from utils.parsing import normalize_cache_root, parse_date, parse_float


class BacktestingPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller

        ttk.Label(self, text="Backtesting Parameters", font=("Arial", 18, "bold")).pack(pady=10)

        intro = (
            "Run a single entry/exit backtest combo or perform a multi-signal parameter sweep. "
            "Sweep mode generates leaderboard and top-N report artifacts in backtest outputs."
        )
        ttk.Label(self, text=intro, wraplength=950, justify="center").pack(pady=5)

        content = ttk.Frame(self)
        content.pack(fill="both", expand=True, padx=30, pady=10)
        content.columnconfigure(0, weight=1)

        strategy_frame = ttk.LabelFrame(content, text="Strategy")
        strategy_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        strategy_frame.columnconfigure(1, weight=1)

        ttk.Label(strategy_frame, text="Run Mode").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        self.run_mode_var = tk.StringVar()
        ttk.Combobox(
            strategy_frame,
            textvariable=self.run_mode_var,
            state="readonly",
            values=["single", "sweep"],
        ).grid(row=0, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(strategy_frame, text="Lookback (bars)").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        self.lookback_days_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.lookback_days_var).grid(row=1, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(strategy_frame, text="Skip (bars)").grid(row=2, column=0, sticky="w", padx=8, pady=6)
        self.skip_days_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.skip_days_var).grid(row=2, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(strategy_frame, text="Entry Signal").grid(row=3, column=0, sticky="w", padx=8, pady=6)
        self.entry_signal_var = tk.StringVar()
        ttk.Combobox(
            strategy_frame,
            textvariable=self.entry_signal_var,
            state="readonly",
            values=["ts_momentum", "ma_trend", "breakout"],
        ).grid(row=3, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(strategy_frame, text="Entry Params (JSON)").grid(row=4, column=0, sticky="w", padx=8, pady=6)
        self.entry_signal_params_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.entry_signal_params_var).grid(row=4, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(strategy_frame, text="Exit Signal").grid(row=5, column=0, sticky="w", padx=8, pady=6)
        self.exit_signal_var = tk.StringVar()
        ttk.Combobox(
            strategy_frame,
            textvariable=self.exit_signal_var,
            state="readonly",
            values=["none", "momentum_flip", "trailing_stop", "max_hold"],
        ).grid(row=5, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(strategy_frame, text="Exit Params (JSON)").grid(row=6, column=0, sticky="w", padx=8, pady=6)
        self.exit_signal_params_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.exit_signal_params_var).grid(row=6, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(strategy_frame, text="Costs (bps)").grid(row=7, column=0, sticky="w", padx=8, pady=6)
        self.costs_bps_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.costs_bps_var).grid(row=7, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(strategy_frame, text="Sweep Entry Grid (JSON)").grid(row=8, column=0, sticky="w", padx=8, pady=6)
        self.sweep_entry_grid_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.sweep_entry_grid_var).grid(row=8, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(strategy_frame, text="Sweep Exit Grid (JSON)").grid(row=9, column=0, sticky="w", padx=8, pady=6)
        self.sweep_exit_grid_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.sweep_exit_grid_var).grid(row=9, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(strategy_frame, text="Sweep Core Grid (JSON)").grid(row=10, column=0, sticky="w", padx=8, pady=6)
        self.sweep_core_grid_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.sweep_core_grid_var).grid(row=10, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(strategy_frame, text="Sweep Seed").grid(row=11, column=0, sticky="w", padx=8, pady=6)
        self.sweep_seed_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.sweep_seed_var).grid(row=11, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(strategy_frame, text="Sweep Max Workers (blank=auto)").grid(row=12, column=0, sticky="w", padx=8, pady=6)
        self.sweep_max_workers_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.sweep_max_workers_var).grid(row=12, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(strategy_frame, text="Sweep Top N").grid(row=13, column=0, sticky="w", padx=8, pady=6)
        self.sweep_top_n_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.sweep_top_n_var).grid(row=13, column=1, sticky="ew", padx=8, pady=6)

        self.fail_fast_var = tk.BooleanVar()
        ttk.Checkbutton(strategy_frame, text="Fail fast on first combo error", variable=self.fail_fast_var).grid(
            row=14, column=0, columnspan=2, sticky="w", padx=8, pady=4
        )

        self.continue_on_error_var = tk.BooleanVar()
        ttk.Checkbutton(strategy_frame, text="Continue on combo errors", variable=self.continue_on_error_var).grid(
            row=15, column=0, columnspan=2, sticky="w", padx=8, pady=4
        )

        ttk.Label(strategy_frame, text="Start Date (YYYY-MM-DD)").grid(row=16, column=0, sticky="w", padx=8, pady=6)
        self.start_date_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.start_date_var).grid(row=16, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(strategy_frame, text="End Date (YYYY-MM-DD)").grid(row=17, column=0, sticky="w", padx=8, pady=6)
        self.end_date_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.end_date_var).grid(row=17, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(strategy_frame, text="Backtest Data Root").grid(row=18, column=0, sticky="w", padx=8, pady=6)
        self.backtest_root_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.backtest_root_var).grid(row=18, column=1, sticky="ew", padx=8, pady=6)

        notes_frame = ttk.LabelFrame(content, text="Run Output")
        notes_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        notes_frame.columnconfigure(0, weight=1)
        self.notes_text = tk.Text(notes_frame, height=14)
        self.notes_text.grid(row=0, column=0, sticky="nsew", padx=8, pady=6)

        button_row = ttk.Frame(self)
        button_row.pack(pady=10)

        ttk.Button(button_row, text="Save Parameters", command=self.save_settings).grid(row=0, column=0, padx=10)
        self.run_button = ttk.Button(button_row, text="Run Backtest", command=self.run_backtest)
        self.run_button.grid(row=0, column=1, padx=10)
        ttk.Button(
            button_row,
            text="Back to Main Menu",
            command=lambda: controller.show_frame("MainMenu"),
        ).grid(row=0, column=2, padx=10)

    def refresh(self) -> None:
        settings = dict(DEFAULT_BACKTEST_SETTINGS)
        settings.update(self.controller.state.backtest_settings)

        self.run_mode_var.set(str(settings.get("run_mode", "single")))
        self.lookback_days_var.set(str(settings.get("lookback_days", "90")))
        self.skip_days_var.set(str(settings.get("skip_days", "5")))
        self.entry_signal_var.set(str(settings.get("entry_signal", "ts_momentum")))
        self.entry_signal_params_var.set(str(settings.get("entry_signal_params", "{}")))
        self.exit_signal_var.set(str(settings.get("exit_signal", "none")))
        self.exit_signal_params_var.set(str(settings.get("exit_signal_params", "{}")))
        self.costs_bps_var.set(str(settings.get("costs_bps", "5")))
        self.sweep_entry_grid_var.set(str(settings.get("sweep_entry_grid", '{"ts_momentum": [{"lookback_days": 90, "skip_days": 5}]}')))
        self.sweep_exit_grid_var.set(str(settings.get("sweep_exit_grid", '{"none": [{}]}')))
        self.sweep_core_grid_var.set(str(settings.get("sweep_core_grid", '{"lookback_days": [90], "skip_days": [5], "costs_bps": [5.0]}')))
        self.sweep_seed_var.set(str(settings.get("sweep_seed", "42")))
        self.sweep_max_workers_var.set(str(settings.get("sweep_max_workers", "")))
        self.sweep_top_n_var.set(str(settings.get("sweep_top_n", "10")))
        self.fail_fast_var.set(bool(settings.get("sweep_fail_fast", False)))
        self.continue_on_error_var.set(bool(settings.get("sweep_continue_on_error", True)))
        self.start_date_var.set(str(settings.get("start_date", "")))
        self.end_date_var.set(str(settings.get("end_date", "")))
        self.backtest_root_var.set(str(settings.get("backtest_data_root", str(BACKTEST_CACHE_DIR))))

        self.notes_text.delete("1.0", tk.END)
        self.notes_text.insert("1.0", str(settings.get("notes", "")))

    def save_settings(self) -> None:
        lookback = parse_float(self.lookback_days_var.get())
        skip = parse_float(self.skip_days_var.get())
        costs_bps = parse_float(self.costs_bps_var.get())

        if lookback is None or lookback < 1 or int(lookback) != lookback:
            messagebox.showinfo("Invalid input", "Lookback must be a positive integer.")
            return
        if skip is None or skip < 0 or int(skip) != skip:
            messagebox.showinfo("Invalid input", "Skip must be a non-negative integer.")
            return
        if costs_bps is None or costs_bps < 0:
            messagebox.showinfo("Invalid input", "Costs must be zero or positive.")
            return

        self.controller.state.backtest_settings = {
            "strategy_name": "Time-Series Momentum",
            "run_mode": self.run_mode_var.get().strip() or "single",
            "lookback_days": str(int(lookback)),
            "skip_days": str(int(skip)),
            "entry_signal": self.entry_signal_var.get().strip() or "ts_momentum",
            "entry_signal_params": self.entry_signal_params_var.get().strip() or "{}",
            "exit_signal": self.exit_signal_var.get().strip() or "none",
            "exit_signal_params": self.exit_signal_params_var.get().strip() or "{}",
            "costs_bps": str(costs_bps),
            "sweep_entry_grid": self.sweep_entry_grid_var.get().strip() or '{"ts_momentum": [{"lookback_days": 90, "skip_days": 5}]}',
            "sweep_exit_grid": self.sweep_exit_grid_var.get().strip() or '{"none": [{}]}',
            "sweep_core_grid": self.sweep_core_grid_var.get().strip() or '{"lookback_days": [90], "skip_days": [5], "costs_bps": [5.0]}',
            "sweep_seed": self.sweep_seed_var.get().strip() or "42",
            "sweep_max_workers": self.sweep_max_workers_var.get().strip(),
            "sweep_top_n": self.sweep_top_n_var.get().strip() or "10",
            "sweep_fail_fast": bool(self.fail_fast_var.get()),
            "sweep_continue_on_error": bool(self.continue_on_error_var.get()),
            "start_date": self.start_date_var.get().strip(),
            "end_date": self.end_date_var.get().strip(),
            "backtest_data_root": self.backtest_root_var.get().strip(),
            "notes": self.notes_text.get("1.0", tk.END).strip(),
        }
        self.controller.persist_state()
        messagebox.showinfo("Saved", "Backtesting parameters saved.")

    def run_backtest(self) -> None:
        tickers = list(self.controller.state.tickers)
        if not tickers:
            messagebox.showinfo("No tickers", "Add tickers before running a backtest.")
            return

        start_date = parse_date(self.start_date_var.get())
        end_date = parse_date(self.end_date_var.get())
        if start_date is None or end_date is None:
            messagebox.showinfo("Invalid dates", "Both start and end dates must be valid YYYY-MM-DD values.")
            return
        if start_date >= end_date:
            messagebox.showinfo("Invalid dates", "Start date must be before end date.")
            return

        self.run_button.config(state="disabled")
        self.notes_text.delete("1.0", tk.END)
        self.notes_text.insert("1.0", f"Running {self.run_mode_var.get() or 'single'} backtest...\n")

        cache_root = normalize_cache_root(self.backtest_root_var.get())

        if (self.run_mode_var.get().strip() or "single") == "sweep":
            thread = threading.Thread(
                target=self._run_sweep_worker,
                args=(tickers, start_date, end_date, cache_root),
                daemon=True,
            )
        else:
            try:
                lookback = int(self.lookback_days_var.get().strip())
                skip = int(self.skip_days_var.get().strip())
                costs_bps = float(self.costs_bps_var.get().strip())
            except ValueError:
                self.run_button.config(state="normal")
                messagebox.showinfo("Invalid input", "Lookback, skip, and costs must be numeric.")
                return
            if lookback < 1 or skip < 0 or costs_bps < 0:
                self.run_button.config(state="normal")
                messagebox.showinfo("Invalid input", "Lookback must be >= 1, skip >= 0, and costs >= 0.")
                return
            thread = threading.Thread(
                target=self._run_single_worker,
                args=(
                    tickers,
                    start_date,
                    end_date,
                    cache_root,
                    lookback,
                    skip,
                    costs_bps,
                    self.entry_signal_var.get().strip() or "ts_momentum",
                    self.entry_signal_params_var.get().strip() or "{}",
                    self.exit_signal_var.get().strip() or "none",
                    self.exit_signal_params_var.get().strip() or "{}",
                ),
                daemon=True,
            )

        thread.start()

    def _run_single_worker(
        self,
        tickers: list[str],
        start_date: date,
        end_date: date,
        cache_root: Path,
        lookback: int,
        skip: int,
        costs_bps: float,
        entry_signal: str,
        entry_signal_params: str,
        exit_signal: str,
        exit_signal_params: str,
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
                entry_signal=entry_signal,
                entry_signal_params=self._parse_json_dict(entry_signal_params, "Entry Params"),
                exit_signal=exit_signal,
                exit_signal_params=self._parse_json_dict(exit_signal_params, "Exit Params"),
            )
        except Exception as exc:
            output_text = f"Backtest failed: {exc}"
        self.after(0, lambda: self._finish_backtest_run(output_text))

    def _run_sweep_worker(
        self,
        tickers: list[str],
        start_date: date,
        end_date: date,
        cache_root: Path,
    ) -> None:
        try:
            max_workers_raw = self.sweep_max_workers_var.get().strip()
            max_workers = int(max_workers_raw) if max_workers_raw else None
            output_text = run_parameter_sweep(
                tickers=tickers,
                start_date=start_date,
                end_date=end_date,
                cache_root=cache_root,
                entry_grid=self._parse_signal_grid_json(self.sweep_entry_grid_var.get(), "Sweep Entry Grid"),
                exit_grid=self._parse_signal_grid_json(self.sweep_exit_grid_var.get(), "Sweep Exit Grid"),
                core_grid=self._parse_core_grid_json(self.sweep_core_grid_var.get(), "Sweep Core Grid"),
                seed=int(self.sweep_seed_var.get().strip() or "42"),
                max_workers=max_workers,
                fail_fast=bool(self.fail_fast_var.get()),
                continue_on_error=bool(self.continue_on_error_var.get()),
                top_n=int(self.sweep_top_n_var.get().strip() or "10"),
            )
        except Exception as exc:
            output_text = f"Backtest sweep failed: {exc}"
        self.after(0, lambda: self._finish_backtest_run(output_text))

    def _parse_json_dict(self, raw: str, label: str) -> dict[str, object]:
        parsed = json.loads(raw)
        if not isinstance(parsed, dict):
            raise ValueError(f"{label} must be a JSON object")
        return parsed

    def _parse_signal_grid_json(self, raw: str, label: str) -> dict[str, list[dict[str, object]]]:
        payload = self._parse_json_dict(raw, label)
        normalized: dict[str, list[dict[str, object]]] = {}
        for key, value in payload.items():
            if not isinstance(value, list):
                raise ValueError(f"{label} values must be lists")
            entries: list[dict[str, object]] = []
            for item in value:
                if not isinstance(item, dict):
                    raise ValueError(f"{label} entries must be JSON objects")
                entries.append(item)
            normalized[str(key)] = entries
        return normalized

    def _parse_core_grid_json(self, raw: str, label: str) -> dict[str, list[object]]:
        payload = self._parse_json_dict(raw, label)
        normalized: dict[str, list[object]] = {}
        for key, value in payload.items():
            if not isinstance(value, list):
                raise ValueError(f"{label} values must be lists")
            normalized[str(key)] = value
        return normalized

    def _finish_backtest_run(self, output_text: str) -> None:
        self.notes_text.delete("1.0", tk.END)
        self.notes_text.insert("1.0", output_text)
        self.run_button.config(state="normal")

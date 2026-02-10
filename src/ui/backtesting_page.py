from __future__ import annotations

import threading
from datetime import date
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

from backtesting.cache_runner import run_time_series_momentum_backtest
from config import BACKTEST_CACHE_DIR, DEFAULT_BACKTEST_SETTINGS
from utils.parsing import normalize_cache_root, parse_date, parse_float


class BacktestingPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller

        ttk.Label(self, text="Backtesting Parameters", font=("Arial", 18, "bold")).pack(
            pady=10
        )

        intro = (
            "MVP strategy is fixed to Time-Series Momentum. "
            "Set only lookback, skip, costs, and date range."
        )
        ttk.Label(self, text=intro, wraplength=900, justify="center").pack(pady=5)

        content = ttk.Frame(self)
        content.pack(fill="both", expand=True, padx=30, pady=10)
        content.columnconfigure(0, weight=1)

        strategy_frame = ttk.LabelFrame(content, text="Strategy")
        strategy_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        strategy_frame.columnconfigure(1, weight=1)

        ttk.Label(strategy_frame, text="Strategy").grid(
            row=0, column=0, sticky="w", padx=8, pady=6
        )
        ttk.Label(strategy_frame, text="Time-Series Momentum").grid(
            row=0, column=1, sticky="w", padx=8, pady=6
        )

        ttk.Label(strategy_frame, text="Lookback (bars)").grid(
            row=1, column=0, sticky="w", padx=8, pady=6
        )
        self.lookback_days_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.lookback_days_var).grid(
            row=1, column=1, sticky="ew", padx=8, pady=6
        )

        ttk.Label(strategy_frame, text="Skip (bars)").grid(
            row=2, column=0, sticky="w", padx=8, pady=6
        )
        self.skip_days_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.skip_days_var).grid(
            row=2, column=1, sticky="ew", padx=8, pady=6
        )

        ttk.Label(strategy_frame, text="Costs (bps)").grid(
            row=3, column=0, sticky="w", padx=8, pady=6
        )
        self.costs_bps_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.costs_bps_var).grid(
            row=3, column=1, sticky="ew", padx=8, pady=6
        )

        ttk.Label(strategy_frame, text="Start Date (YYYY-MM-DD)").grid(
            row=4, column=0, sticky="w", padx=8, pady=6
        )
        self.start_date_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.start_date_var).grid(
            row=4, column=1, sticky="ew", padx=8, pady=6
        )

        ttk.Label(strategy_frame, text="End Date (YYYY-MM-DD)").grid(
            row=5, column=0, sticky="w", padx=8, pady=6
        )
        self.end_date_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.end_date_var).grid(
            row=5, column=1, sticky="ew", padx=8, pady=6
        )

        ttk.Label(strategy_frame, text="Backtest Data Root").grid(
            row=6, column=0, sticky="w", padx=8, pady=6
        )
        self.backtest_root_var = tk.StringVar()
        ttk.Entry(strategy_frame, textvariable=self.backtest_root_var).grid(
            row=6, column=1, sticky="ew", padx=8, pady=6
        )

        notes_frame = ttk.LabelFrame(content, text="Run Output")
        notes_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        notes_frame.columnconfigure(0, weight=1)
        self.notes_text = tk.Text(notes_frame, height=14)
        self.notes_text.grid(row=0, column=0, sticky="nsew", padx=8, pady=6)

        button_row = ttk.Frame(self)
        button_row.pack(pady=10)

        ttk.Button(button_row, text="Save Parameters", command=self.save_settings).grid(
            row=0, column=0, padx=10
        )
        self.run_button = ttk.Button(
            button_row,
            text="Run Backtest",
            command=self.run_backtest,
        )
        self.run_button.grid(row=0, column=1, padx=10)
        ttk.Button(
            button_row,
            text="Back to Main Menu",
            command=lambda: controller.show_frame("MainMenu"),
        ).grid(row=0, column=2, padx=10)

    def refresh(self) -> None:
        settings = dict(DEFAULT_BACKTEST_SETTINGS)
        settings.update(self.controller.state.backtest_settings)
        self.lookback_days_var.set(str(settings.get("lookback_days", "90")))
        self.skip_days_var.set(str(settings.get("skip_days", "5")))
        self.costs_bps_var.set(str(settings.get("costs_bps", "5")))
        self.start_date_var.set(settings.get("start_date", ""))
        self.end_date_var.set(settings.get("end_date", ""))
        self.backtest_root_var.set(
            settings.get("backtest_data_root", str(BACKTEST_CACHE_DIR))
        )
        self.notes_text.delete("1.0", tk.END)
        self.notes_text.insert("1.0", settings.get("notes", ""))

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
            "lookback_days": str(int(lookback)),
            "skip_days": str(int(skip)),
            "costs_bps": str(costs_bps),
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

        try:
            lookback = int(self.lookback_days_var.get().strip())
            skip = int(self.skip_days_var.get().strip())
            costs_bps = float(self.costs_bps_var.get().strip())
        except ValueError:
            messagebox.showinfo("Invalid input", "Lookback, skip, and costs must be numeric.")
            return
        if lookback < 1 or skip < 0 or costs_bps < 0:
            messagebox.showinfo(
                "Invalid input",
                "Lookback must be >= 1, skip must be >= 0, and costs must be >= 0.",
            )
            return

        start_date = parse_date(self.start_date_var.get())
        end_date = parse_date(self.end_date_var.get())
        if start_date is None or end_date is None:
            messagebox.showinfo(
                "Invalid dates", "Both start and end dates must be valid YYYY-MM-DD values."
            )
            return
        if start_date >= end_date:
            messagebox.showinfo("Invalid dates", "Start date must be before end date.")
            return

        self.run_button.config(state="disabled")
        self.notes_text.delete("1.0", tk.END)
        self.notes_text.insert("1.0", "Running Time-Series Momentum backtest...\n")

        cache_root = normalize_cache_root(self.backtest_root_var.get())
        thread = threading.Thread(
            target=self._run_backtest_worker,
            args=(tickers, start_date, end_date, cache_root, lookback, skip, costs_bps),
            daemon=True,
        )
        thread.start()

    def _run_backtest_worker(
        self,
        tickers: list[str],
        start_date: date,
        end_date: date,
        cache_root: Path,
        lookback: int,
        skip: int,
        costs_bps: float,
    ) -> None:
        output_text = run_time_series_momentum_backtest(
            tickers=tickers,
            start_date=start_date,
            end_date=end_date,
            cache_root=cache_root,
            lookback_days=lookback,
            skip_days=skip,
            costs_bps=costs_bps,
        )
        self.after(0, lambda: self._finish_backtest_run(output_text))

    def _finish_backtest_run(self, output_text: str) -> None:
        self.notes_text.delete("1.0", tk.END)
        self.notes_text.insert("1.0", output_text)
        self.run_button.config(state="normal")

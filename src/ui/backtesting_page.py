from __future__ import annotations

import threading
from datetime import date, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

from backtesting.cache_runner import run_backtest_cache
from config import BACKTEST_CACHE_DIR, DEFAULT_BACKTEST_SETTINGS
from utils.parsing import (
    normalize_cache_root,
    parse_date,
    parse_float,
)


class BacktestingPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller

        ttk.Label(self, text="Backtesting Parameters", font=("Arial", 18, "bold")).pack(
            pady=10
        )

        intro = (
            "Configure strategy selection, data, and execution realism. "
            "Use intraday data for realistic options backtests; EOD is cheaper and "
            "better for swing-style research."
        )
        ttk.Label(self, text=intro, wraplength=900, justify="center").pack(pady=5)

        content = ttk.Frame(self)
        content.pack(fill="both", expand=True, padx=30, pady=10)
        content.columnconfigure(0, weight=1)
        content.columnconfigure(1, weight=1)

        strategy_frame = ttk.LabelFrame(content, text="Strategy / Model")
        strategy_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        strategy_frame.columnconfigure(1, weight=1)

        ttk.Label(strategy_frame, text="Strategy Type").grid(
            row=0, column=0, sticky="w", padx=8, pady=6
        )
        self.strategy_type_var = tk.StringVar()
        self.strategy_type_combo = ttk.Combobox(
            strategy_frame,
            textvariable=self.strategy_type_var,
            values=["Rule-based", "ML signals", "Hybrid"],
            state="readonly",
        )
        self.strategy_type_combo.grid(row=0, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(strategy_frame, text="AI Model").grid(
            row=1, column=0, sticky="w", padx=8, pady=6
        )
        self.ai_model_var = tk.StringVar()
        self.ai_model_combo = ttk.Combobox(
            strategy_frame,
            textvariable=self.ai_model_var,
            values=[
                "Baseline (No ML)",
                "Random Forest",
                "XGBoost",
                "LSTM",
                "Custom",
            ],
            state="readonly",
        )
        self.ai_model_combo.grid(row=1, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(strategy_frame, text="Stat Method").grid(
            row=2, column=0, sticky="w", padx=8, pady=6
        )
        self.stat_method_var = tk.StringVar()
        self.stat_method_combo = ttk.Combobox(
            strategy_frame,
            textvariable=self.stat_method_var,
            values=[
                "Momentum",
                "Mean Reversion",
                "Volatility Breakout",
                "Pairs Trading",
                "Custom",
            ],
            state="readonly",
        )
        self.stat_method_combo.grid(row=2, column=1, sticky="ew", padx=8, pady=6)

        data_frame = ttk.LabelFrame(content, text="Data")
        data_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        data_frame.columnconfigure(1, weight=1)

        ttk.Label(data_frame, text="Data Source").grid(
            row=0, column=0, sticky="w", padx=8, pady=6
        )
        self.data_source_var = tk.StringVar()
        self.data_source_combo = ttk.Combobox(
            data_frame,
            textvariable=self.data_source_var,
            values=["Massive API", "Discount Option Data", "Custom CSV"],
            state="readonly",
        )
        self.data_source_combo.grid(row=0, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(data_frame, text="Granularity").grid(
            row=1, column=0, sticky="w", padx=8, pady=6
        )
        self.data_granularity_var = tk.StringVar()
        self.data_granularity_combo = ttk.Combobox(
            data_frame,
            textvariable=self.data_granularity_var,
            values=["Daily (EOD)", "60m", "30m", "15m", "5m", "1m"],
            state="readonly",
        )
        self.data_granularity_combo.grid(row=1, column=1, sticky="ew", padx=8, pady=6)

        ttk.Label(data_frame, text="Start Date (YYYY-MM-DD)").grid(
            row=2, column=0, sticky="w", padx=8, pady=6
        )
        self.start_date_var = tk.StringVar()
        ttk.Entry(data_frame, textvariable=self.start_date_var).grid(
            row=2, column=1, sticky="ew", padx=8, pady=6
        )

        ttk.Label(data_frame, text="End Date (YYYY-MM-DD)").grid(
            row=3, column=0, sticky="w", padx=8, pady=6
        )
        self.end_date_var = tk.StringVar()
        ttk.Entry(data_frame, textvariable=self.end_date_var).grid(
            row=3, column=1, sticky="ew", padx=8, pady=6
        )

        ttk.Label(data_frame, text="Training Split").grid(
            row=4, column=0, sticky="w", padx=8, pady=6
        )
        self.training_split_var = tk.StringVar()
        ttk.Entry(data_frame, textvariable=self.training_split_var).grid(
            row=4, column=1, sticky="ew", padx=8, pady=6
        )

        ttk.Label(data_frame, text="Backtest Data Root").grid(
            row=5, column=0, sticky="w", padx=8, pady=6
        )
        self.backtest_root_var = tk.StringVar()
        ttk.Entry(data_frame, textvariable=self.backtest_root_var).grid(
            row=5, column=1, sticky="ew", padx=8, pady=6
        )

        realism_frame = ttk.LabelFrame(content, text="Execution Realism")
        realism_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        realism_frame.columnconfigure(1, weight=1)

        ttk.Label(realism_frame, text="Slippage (bps)").grid(
            row=0, column=0, sticky="w", padx=8, pady=6
        )
        self.slippage_bps_var = tk.StringVar()
        ttk.Entry(realism_frame, textvariable=self.slippage_bps_var).grid(
            row=0, column=1, sticky="ew", padx=8, pady=6
        )

        ttk.Label(realism_frame, text="Commission per Contract ($)").grid(
            row=1, column=0, sticky="w", padx=8, pady=6
        )
        self.commission_var = tk.StringVar()
        ttk.Entry(realism_frame, textvariable=self.commission_var).grid(
            row=1, column=1, sticky="ew", padx=8, pady=6
        )

        ttk.Label(realism_frame, text="Fill Probability (0-1)").grid(
            row=2, column=0, sticky="w", padx=8, pady=6
        )
        self.fill_probability_var = tk.StringVar()
        ttk.Entry(realism_frame, textvariable=self.fill_probability_var).grid(
            row=2, column=1, sticky="ew", padx=8, pady=6
        )

        self.use_bid_ask_var = tk.BooleanVar()
        ttk.Checkbutton(
            realism_frame, text="Use bid/ask spread", variable=self.use_bid_ask_var
        ).grid(row=3, column=0, columnspan=2, sticky="w", padx=8, pady=4)

        self.walk_forward_var = tk.BooleanVar()
        ttk.Checkbutton(
            realism_frame, text="Walk-forward retraining", variable=self.walk_forward_var
        ).grid(row=4, column=0, columnspan=2, sticky="w", padx=8, pady=4)

        ttk.Label(realism_frame, text="Max Position Size (%)").grid(
            row=5, column=0, sticky="w", padx=8, pady=6
        )
        self.max_position_pct_var = tk.StringVar()
        ttk.Entry(realism_frame, textvariable=self.max_position_pct_var).grid(
            row=5, column=1, sticky="ew", padx=8, pady=6
        )

        notes_frame = ttk.LabelFrame(content, text="Notes / Metrics")
        notes_frame.grid(row=1, column=1, sticky="nsew", padx=10, pady=10)
        notes_frame.columnconfigure(0, weight=1)
        self.notes_text = tk.Text(notes_frame, height=8)
        self.notes_text.grid(row=0, column=0, sticky="nsew", padx=8, pady=6)

        button_row = ttk.Frame(self)
        button_row.pack(pady=10)

        ttk.Button(button_row, text="Save Parameters", command=self.save_settings).grid(
            row=0, column=0, padx=10
        )
        self.run_button = ttk.Button(
            button_row,
            text="Run Backtest (Cache Data)",
            command=self.run_backtest_cache,
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
        self.strategy_type_var.set(settings.get("strategy_type", "Rule-based"))
        self.ai_model_var.set(settings.get("ai_model", "Baseline (No ML)"))
        self.stat_method_var.set(settings.get("stat_method", "Momentum"))
        self.data_source_var.set(settings.get("data_source", "Massive API"))
        self.data_granularity_var.set(settings.get("data_granularity", "Daily (EOD)"))
        self.start_date_var.set(settings.get("start_date", ""))
        self.end_date_var.set(settings.get("end_date", ""))
        self.training_split_var.set(settings.get("training_split", "70/30"))
        self.backtest_root_var.set(
            settings.get("backtest_data_root", str(BACKTEST_CACHE_DIR))
        )
        self.slippage_bps_var.set(settings.get("slippage_bps", "5"))
        self.commission_var.set(settings.get("commission_per_contract", "0.65"))
        self.fill_probability_var.set(settings.get("fill_probability", "0.9"))
        self.use_bid_ask_var.set(bool(settings.get("use_bid_ask", True)))
        self.walk_forward_var.set(bool(settings.get("model_walk_forward", False)))
        self.max_position_pct_var.set(settings.get("max_position_pct", "10"))
        self.notes_text.delete("1.0", tk.END)
        self.notes_text.insert("1.0", settings.get("notes", ""))

    def save_settings(self) -> None:
        slippage_bps = parse_float(self.slippage_bps_var.get())
        commission = parse_float(self.commission_var.get())
        fill_probability = parse_float(self.fill_probability_var.get())
        max_position_pct = parse_float(self.max_position_pct_var.get())

        if slippage_bps is None or slippage_bps < 0:
            messagebox.showinfo("Invalid input", "Slippage must be zero or positive.")
            return
        if commission is None or commission < 0:
            messagebox.showinfo("Invalid input", "Commission must be zero or positive.")
            return
        if fill_probability is None or not (0 <= fill_probability <= 1):
            messagebox.showinfo(
                "Invalid input", "Fill probability must be between 0 and 1."
            )
            return
        if max_position_pct is None or not (0 < max_position_pct <= 100):
            messagebox.showinfo(
                "Invalid input", "Max position size must be between 0 and 100."
            )
            return

        self.controller.state.backtest_settings = {
            "strategy_type": self.strategy_type_var.get(),
            "ai_model": self.ai_model_var.get(),
            "stat_method": self.stat_method_var.get(),
            "data_source": self.data_source_var.get(),
            "data_granularity": self.data_granularity_var.get(),
            "start_date": self.start_date_var.get().strip(),
            "end_date": self.end_date_var.get().strip(),
            "training_split": self.training_split_var.get().strip(),
            "backtest_data_root": self.backtest_root_var.get().strip(),
            "slippage_bps": str(slippage_bps),
            "commission_per_contract": str(commission),
            "fill_probability": str(fill_probability),
            "use_bid_ask": self.use_bid_ask_var.get(),
            "model_walk_forward": self.walk_forward_var.get(),
            "max_position_pct": str(max_position_pct),
            "notes": self.notes_text.get("1.0", tk.END).strip(),
        }
        self.controller.persist_state()
        messagebox.showinfo("Saved", "Backtesting parameters saved.")

    def run_backtest_cache(self) -> None:
        if not self.controller.api_key:
            messagebox.showinfo("Missing key", "Enter a Massive API key first.")
            return
        tickers = list(self.controller.state.tickers)
        if not tickers:
            messagebox.showinfo("No tickers", "Add tickers before running a backtest.")
            return

        start_date = parse_date(self.start_date_var.get())
        end_date = parse_date(self.end_date_var.get())
        if end_date is None:
            end_date = date.today()
        five_year_start = end_date - timedelta(days=365 * 5)
        if start_date is None or start_date != five_year_start:
            start_date = five_year_start
        if start_date >= end_date:
            messagebox.showinfo("Invalid dates", "Start date must be before end date.")
            return

        self.run_button.config(state="disabled")
        self.notes_text.delete("1.0", tk.END)
        self.notes_text.insert(
            "1.0",
            "Running backtest data cache (5 years @ 1-minute)...\n"
            "This may take a while for 1-minute data.\n",
        )

        cache_root = normalize_cache_root(self.backtest_root_var.get())
        thread = threading.Thread(
            target=self._run_backtest_worker,
            args=(tickers, start_date, end_date, cache_root),
            daemon=True,
        )
        thread.start()

    def _run_backtest_worker(
        self, tickers: list[str], start_date: date, end_date: date, cache_root: Path
    ) -> None:
        output_text = run_backtest_cache(
            tickers=tickers,
            start_date=start_date,
            end_date=end_date,
            cache_root=cache_root,
            api_key=self.controller.api_key,
        )
        self.after(0, lambda: self._finish_backtest_run(output_text))

    def _finish_backtest_run(self, output_text: str) -> None:
        self.notes_text.delete("1.0", tk.END)
        self.notes_text.insert("1.0", output_text)
        self.run_button.config(state="normal")

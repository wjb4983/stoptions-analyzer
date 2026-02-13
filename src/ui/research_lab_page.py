from __future__ import annotations

import threading
from datetime import date, timedelta
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Any, Callable

from backtesting.cache_runner import (
    run_multi_signal_backtest,
    run_strategy_optimization,
    run_walk_forward_backtest,
)
from config import DEFAULT_BACKTEST_SETTINGS
from utils.parsing import normalize_cache_root, parse_date


class ResearchLabPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller
        self._is_running = False

        ttk.Label(self, text="Research Lab", font=("Arial", 18, "bold")).pack(pady=(12, 8))
        ttk.Label(
            self,
            text=(
                "Advanced orchestration workflows for validating hypotheses, tuning parameters, "
                "stress testing, and governing experiments."
            ),
            wraplength=900,
            justify="center",
        ).pack(pady=(0, 14), padx=20)

        tiles = ttk.Frame(self)
        tiles.pack(fill="x", padx=40)
        tiles.columnconfigure(0, weight=1)
        tiles.columnconfigure(1, weight=1)

        self._build_tile(
            tiles,
            row=0,
            column=0,
            title="Walk-forward validation",
            description="Run out-of-sample fold validation using cache_runner walk-forward entry points.",
            action_label="Run Validation",
            action=self.run_walk_forward,
        )
        self._build_tile(
            tiles,
            row=0,
            column=1,
            title="Parameter optimization",
            description="Launch multi-objective optimization against entry/exit signal combinations.",
            action_label="Run Optimization",
            action=self.run_optimization,
        )
        self._build_tile(
            tiles,
            row=1,
            column=0,
            title="Stress/scenario tests",
            description="Execute multi-signal backtests for scenario-aware robustness checks.",
            action_label="Run Stress Test",
            action=self.run_stress_tests,
        )
        self._build_tile(
            tiles,
            row=1,
            column=1,
            title="Experiment comparison and governance",
            description="Open the full backtesting workstation to compare runs and review governance gates.",
            action_label="Open Governance Workspace",
            action=self.open_governance_workspace,
        )

        button_row = ttk.Frame(self)
        button_row.pack(pady=(16, 8))
        ttk.Button(
            button_row,
            text="Back to Main Menu",
            command=lambda: self.controller.show_frame("MainMenu"),
        ).pack()

        output_frame = ttk.LabelFrame(self, text="Research Lab Output")
        output_frame.pack(fill="both", expand=True, padx=40, pady=(4, 20))
        self.output_text = tk.Text(output_frame, height=12, wrap="word")
        self.output_text.pack(fill="both", expand=True, padx=8, pady=8)

    def _build_tile(
        self,
        parent: ttk.Frame,
        *,
        row: int,
        column: int,
        title: str,
        description: str,
        action_label: str,
        action: Callable[[], None],
    ) -> None:
        tile = ttk.LabelFrame(parent, text=title)
        tile.grid(row=row, column=column, sticky="nsew", padx=8, pady=8)
        tile.columnconfigure(0, weight=1)

        ttk.Label(tile, text=description, wraplength=420, justify="left").grid(
            row=0, column=0, sticky="w", padx=10, pady=(8, 10)
        )
        ttk.Button(tile, text=action_label, command=action).grid(
            row=1, column=0, sticky="w", padx=10, pady=(0, 10)
        )

    def _append_output(self, text: str) -> None:
        self.output_text.insert(tk.END, f"{text}\n")
        self.output_text.see(tk.END)

    def _start_worker(self, target: Callable[[dict[str, Any]], str], label: str) -> None:
        if self._is_running:
            messagebox.showinfo("Run in progress", "Wait for the current Research Lab run to finish.")
            return

        context = self._build_common_context()
        if context is None:
            return

        self._is_running = True
        self._append_output(f"Starting: {label}")

        def worker() -> None:
            try:
                output = target(context)
            except Exception as exc:
                output = f"Research workflow failed: {exc}"
            self.after(0, lambda: self._finish_worker(output))

        threading.Thread(target=worker, daemon=True).start()

    def _finish_worker(self, output: str) -> None:
        self._is_running = False
        self._append_output(output)

    def _build_common_context(self) -> dict[str, Any] | None:
        tickers = list(self.controller.state.tickers)
        if not tickers:
            messagebox.showinfo("No tickers", "Add tickers before launching Research Lab workflows.")
            return None

        default_start = date.today() - timedelta(days=365)
        default_end = date.today()
        configured_start = parse_date(str(DEFAULT_BACKTEST_SETTINGS.get("start_date", "")))
        configured_end = parse_date(str(DEFAULT_BACKTEST_SETTINGS.get("end_date", "")))

        start_date = configured_start or default_start
        end_date = configured_end or default_end
        if start_date >= end_date:
            start_date = default_start
            end_date = default_end

        cache_root = normalize_cache_root(str(DEFAULT_BACKTEST_SETTINGS.get("backtest_data_root", "")))

        lookback = int(DEFAULT_BACKTEST_SETTINGS.get("lookback_days", "90"))
        skip = int(DEFAULT_BACKTEST_SETTINGS.get("skip_days", "5"))
        costs_bps = float(DEFAULT_BACKTEST_SETTINGS.get("costs_bps", "5"))

        return {
            "tickers": tickers,
            "start_date": start_date,
            "end_date": end_date,
            "cache_root": cache_root,
            "lookback": lookback,
            "skip": skip,
            "costs_bps": costs_bps,
        }

    def run_walk_forward(self) -> None:
        self._start_worker(self._run_walk_forward_workflow, "Walk-forward validation")

    def run_optimization(self) -> None:
        self._start_worker(self._run_optimization_workflow, "Parameter optimization")

    def run_stress_tests(self) -> None:
        self._start_worker(self._run_stress_workflow, "Stress/scenario tests")

    def open_governance_workspace(self) -> None:
        self.controller.show_frame("BacktestingPage")

    def _build_signal_grids(self, context: dict[str, Any]) -> tuple[dict[str, list[dict[str, object]]], dict[str, list[dict[str, object]]], dict[str, list[object]]]:
        entry_grid = {"ts_momentum": [{}], "breakout": [{}]}
        exit_grid = {"none": [{}], "momentum_flip": [{}]}
        core_grid = {
            "lookback_days": [int(context["lookback"])],
            "skip_days": [int(context["skip"])],
            "costs_bps": [float(context["costs_bps"])],
        }
        return entry_grid, exit_grid, core_grid

    def _run_walk_forward_workflow(self, context: dict[str, Any]) -> str:
        entry_grid, exit_grid, core_grid = self._build_signal_grids(context)
        return run_walk_forward_backtest(
            tickers=list(context["tickers"]),
            start_date=context["start_date"],
            end_date=context["end_date"],
            cache_root=context["cache_root"],
            entry_grid=entry_grid,
            exit_grid=exit_grid,
            core_grid=core_grid,
            train_fraction=0.70,
            validation_fraction=0.15,
            test_fraction=0.15,
            step_fraction=0.15,
            governance_metadata={"promotion_state": "research", "approval_status": "pending"},
        )

    def _run_optimization_workflow(self, context: dict[str, Any]) -> str:
        entry_grid, exit_grid, core_grid = self._build_signal_grids(context)
        return run_strategy_optimization(
            tickers=list(context["tickers"]),
            start_date=context["start_date"],
            end_date=context["end_date"],
            cache_root=context["cache_root"],
            entry_grid=entry_grid,
            exit_grid=exit_grid,
            core_grid=core_grid,
            seed=42,
            n_trials=20,
            sampler_name="tpe",
            partial_period_fractions=[0.5, 1.0],
            governance_metadata={"promotion_state": "research", "approval_status": "pending"},
        )

    def _run_stress_workflow(self, context: dict[str, Any]) -> str:
        return run_multi_signal_backtest(
            tickers=list(context["tickers"]),
            start_date=context["start_date"],
            end_date=context["end_date"],
            cache_root=context["cache_root"],
            lookback_days=int(context["lookback"]),
            skip_days=int(context["skip"]),
            costs_bps=float(context["costs_bps"]),
            timeframe=str(DEFAULT_BACKTEST_SETTINGS.get("timeframe", "1m")),
            entry_signals=["ts_momentum", "breakout"],
            exit_signals=["none", "momentum_flip", "trailing_stop"],
            governance_metadata={"promotion_state": "research", "approval_status": "pending"},
        )

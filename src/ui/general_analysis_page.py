from __future__ import annotations

import json
import math
import os
import random
import socket
import time
import threading
import re
import webbrowser
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, time as dt_time, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk
from urllib.error import HTTPError, URLError
from zoneinfo import ZoneInfo

import numpy as np

from analysis.cross_sectional import (
    CrossSectionalSettings,
    MomentumSettings,
    STRATEGY_REGISTRY,
    compute_cross_sectional_momentum,
)
from analysis.prompt_pack import write_prompt_pack
from analysis.reporting import format_cross_sectional_report, format_time_series_report
from analysis.time_series import (
    TIME_SERIES_STRATEGY_REGISTRY,
    TimeSeriesMomentumSettings,
    TimeSeriesSettings,
    compute_time_series_momentum,
)
from execution import JOB_ANALYSIS_CALLABLE
from config import (
    ANALYSIS_OUTPUT_DIR,
    API_KEY_PATH,
    BACKTEST_CACHE_DIR,
    BACKTEST_OUTPUT_DIR,
    CONFIG_DIR,
    DATA_DIR,
    DEFAULT_BACKTEST_SETTINGS,
    DEFAULT_GENERAL_ANALYSIS_SETTINGS,
    HORIZON_CONFIGS,
)
from data_access.api_client import MassiveApiClient
from data_access.cache import _safe_ticker_name, load_cached_market_data, save_cached_market_data
from data_access.option_loader import load_option_records
from ui.helpers import load_api_key, save_api_key
from utils.parsing import (
    _coerce_number,
    _get_nested_value,
    _has_fundamentals_data,
    _parse_iso_date,
    build_npz_payload,
    chunk_results_by_year,
    combine_greeks,
    effective_market_date,
    extract_greeks,
    format_http_error_detail,
    format_strike,
    normalize_cache_root,
    normalize_contract_type,
    normalize_likelihood_threshold,
    normalize_option_records,
    option_likelihood,
    option_mid_price,
    parse_date,
    parse_float,
    parse_int,
)


class GeneralAnalysisPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller
        self.api_client: MassiveApiClient | None = None
        self._rate_lock = threading.Lock()
        self._last_request_time = 0.0
        self._min_request_interval = 0.05
        self._grouped_ticker_pattern = re.compile(r"^[A-Z0-9]+$")
        self._latest_combined_report = ""
        self._latest_report_path = ""
        self._run_history: list[dict[str, str]] = []

        header = ttk.Label(self, text="General Analysis", font=("Arial", 18, "bold"))
        header.pack(pady=10)

        description = ttk.Label(
            self,
            text=(
                "Run cross-sectional, time-series, and factor analysis across your current "
                "stock universe. Tune parameters and export results to a text file."
            ),
            wraplength=700,
            justify="center",
        )
        description.pack(pady=(0, 15))

        form_frame = ttk.LabelFrame(self, text="Analysis Settings")
        form_frame.pack(padx=40, pady=10, fill="x")
        form_frame.columnconfigure(1, weight=1)

        ttk.Label(form_frame, text="Analysis Type").grid(
            row=0, column=0, padx=10, pady=6, sticky="w"
        )
        self.analysis_type_var = tk.StringVar()
        self.analysis_type_dropdown = ttk.Combobox(
            form_frame,
            textvariable=self.analysis_type_var,
            values=[
                "Cross-Sectional",
                "Time-Series",
                "Cross-Sectional + Time-Series",
            ],
            state="readonly",
            width=30,
        )
        self.analysis_type_dropdown.grid(row=0, column=1, padx=10, pady=6, sticky="ew")

        ttk.Label(form_frame, text="Cross-Sectional Strategy").grid(
            row=1, column=0, padx=10, pady=6, sticky="w"
        )
        self.cross_sectional_var = tk.StringVar()
        working_strategies = ["Momentum", "Low Volatility", "Liquidity", "Size"]
        non_working = [
            "Value",
            "Quality",
            "Investment",
            "Earnings Momentum",
            "Carry / Yield",
        ]
        self.strategy_label_to_key: dict[str, str] = {
            name: name for name in working_strategies
        }
        self.strategy_label_to_key.update(
            {f"{name} (no work)": name for name in non_working}
        )
        self.strategy_key_to_label = {
            key: label for label, key in self.strategy_label_to_key.items()
        }
        dropdown_values = list(self.strategy_label_to_key.keys())
        self.cross_sectional_dropdown = ttk.Combobox(
            form_frame,
            textvariable=self.cross_sectional_var,
            values=dropdown_values,
            state="readonly",
            width=30,
        )
        self.cross_sectional_dropdown.grid(row=1, column=1, padx=10, pady=6, sticky="ew")
        self.cross_sectional_dropdown.bind("<<ComboboxSelected>>", self._on_strategy_change)

        self.strategy_detail_var = tk.StringVar(value="Required data: prices")
        self.strategy_detail_label = ttk.Label(form_frame, textvariable=self.strategy_detail_var)
        self.strategy_detail_label.grid(row=2, column=0, columnspan=2, padx=10, pady=2, sticky="w")

        ttk.Label(form_frame, text="Lookback (days)").grid(
            row=3, column=0, padx=10, pady=6, sticky="w"
        )
        self.lookback_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.lookback_var).grid(
            row=3, column=1, padx=10, pady=6, sticky="ew"
        )

        ttk.Label(form_frame, text="Skip (days)").grid(
            row=4, column=0, padx=10, pady=6, sticky="w"
        )
        self.skip_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.skip_var).grid(
            row=4, column=1, padx=10, pady=6, sticky="ew"
        )

        ttk.Label(form_frame, text="Top Quantile (0-1)").grid(
            row=5, column=0, padx=10, pady=6, sticky="w"
        )
        self.top_quantile_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.top_quantile_var).grid(
            row=5, column=1, padx=10, pady=6, sticky="ew"
        )

        ttk.Label(form_frame, text="Bottom Quantile (0-1)").grid(
            row=6, column=0, padx=10, pady=6, sticky="w"
        )
        self.bottom_quantile_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.bottom_quantile_var).grid(
            row=6, column=1, padx=10, pady=6, sticky="ew"
        )

        self.momentum_label = ttk.Label(form_frame, text="Momentum Toggles")
        self.momentum_label.grid(row=7, column=0, padx=10, pady=6, sticky="w")
        self.momentum_toggles_frame = ttk.Frame(form_frame)
        self.momentum_toggles_frame.grid(row=7, column=1, padx=10, pady=6, sticky="w")
        self.momentum_volatility_var = tk.BooleanVar()
        self.momentum_residual_var = tk.BooleanVar()
        self.momentum_multi_horizon_var = tk.BooleanVar()
        self.momentum_volatility_check = ttk.Checkbutton(
            self.momentum_toggles_frame,
            text="Volatility Scaling",
            variable=self.momentum_volatility_var,
        )
        self.momentum_residual_check = ttk.Checkbutton(
            self.momentum_toggles_frame,
            text="Residual Momentum",
            variable=self.momentum_residual_var,
        )
        self.momentum_multi_horizon_check = ttk.Checkbutton(
            self.momentum_toggles_frame,
            text="Multi-Horizon",
            variable=self.momentum_multi_horizon_var,
        )
        self.momentum_volatility_check.grid(row=0, column=0, padx=5, sticky="w")
        self.momentum_residual_check.grid(row=0, column=1, padx=5, sticky="w")
        self.momentum_multi_horizon_check.grid(row=0, column=2, padx=5, sticky="w")

        ttk.Label(form_frame, text="Output Directory").grid(
            row=8, column=0, padx=10, pady=6, sticky="w"
        )
        self.output_dir_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.output_dir_var).grid(
            row=8, column=1, padx=10, pady=6, sticky="ew"
        )

        button_row = ttk.Frame(self)
        button_row.pack(pady=15)

        ttk.Button(button_row, text="Run Analysis", command=self.run_analysis).grid(
            row=0, column=0, padx=10
        )
        ttk.Button(button_row, text="Export Prompt Pack", command=self.export_prompt_pack).grid(
            row=0, column=1, padx=10
        )
        ttk.Button(
            button_row,
            text="Back to Main Menu",
            command=lambda: controller.show_frame("MainMenu"),
        ).grid(row=0, column=2, padx=10)

        output_frame = ttk.LabelFrame(self, text="Run Timeline")
        output_frame.pack(padx=40, pady=(5, 15), fill="both", expand=True)
        output_frame.rowconfigure(0, weight=1)
        output_frame.columnconfigure(0, weight=1)

        columns = ("run_type", "start_time", "end_time", "status", "key_metrics", "artifact_path")
        self.run_tree = ttk.Treeview(output_frame, columns=columns, show="headings", height=12)
        self.run_tree.heading("run_type", text="Run Type")
        self.run_tree.heading("start_time", text="Start")
        self.run_tree.heading("end_time", text="End")
        self.run_tree.heading("status", text="Status")
        self.run_tree.heading("key_metrics", text="Key Metrics")
        self.run_tree.heading("artifact_path", text="Artifact Path")
        self.run_tree.column("run_type", width=180, anchor="w")
        self.run_tree.column("start_time", width=170, anchor="w")
        self.run_tree.column("end_time", width=170, anchor="w")
        self.run_tree.column("status", width=90, anchor="center")
        self.run_tree.column("key_metrics", width=320, anchor="w")
        self.run_tree.column("artifact_path", width=350, anchor="w")
        self.run_tree.bind("<<TreeviewSelect>>", lambda _event: self._update_run_actions_state())
        self.run_tree.bind("<Double-1>", lambda _event: self._open_selected_artifact_directory())

        output_scrollbar = ttk.Scrollbar(
            output_frame, orient="vertical", command=self.run_tree.yview
        )
        self.run_tree.configure(yscrollcommand=output_scrollbar.set)
        self.run_tree.grid(row=0, column=0, sticky="nsew")
        output_scrollbar.grid(row=0, column=1, sticky="ns")

        action_row = ttk.Frame(output_frame)
        action_row.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        action_row.columnconfigure(2, weight=1)
        self.open_artifact_button = ttk.Button(
            action_row,
            text="Open Artifact Directory",
            command=self._open_selected_artifact_directory,
            state="disabled",
        )
        self.open_artifact_button.grid(row=0, column=0, padx=(0, 8), sticky="w")
        self.copy_artifact_button = ttk.Button(
            action_row,
            text="Copy Artifact Path",
            command=self._copy_selected_artifact_path,
            state="disabled",
        )
        self.copy_artifact_button.grid(row=0, column=1, padx=(0, 8), sticky="w")
        self.run_actions_var = tk.StringVar(value="Select a run to open or copy its artifact path.")
        ttk.Label(action_row, textvariable=self.run_actions_var).grid(row=0, column=2, sticky="w")

    def refresh(self) -> None:
        settings = dict(DEFAULT_GENERAL_ANALYSIS_SETTINGS)
        settings.update(self.controller.state.general_analysis_settings or {})
        self.analysis_type_var.set(settings.get("analysis_type", "Cross-Sectional"))
        strategy_key = settings.get("cross_sectional_strategy", "Momentum")
        self.cross_sectional_var.set(
            self.strategy_key_to_label.get(strategy_key, "Momentum")
        )
        self._update_strategy_detail()
        self.lookback_var.set(str(settings.get("lookback_days", 90)))
        self.skip_var.set(str(settings.get("skip_days", 5)))
        self.top_quantile_var.set(str(settings.get("top_quantile", 0.2)))
        self.bottom_quantile_var.set(str(settings.get("bottom_quantile", 0.2)))
        self.momentum_volatility_var.set(settings.get("momentum_use_volatility_scaling", False))
        self.momentum_residual_var.set(settings.get("momentum_use_residual", False))
        self.momentum_multi_horizon_var.set(settings.get("momentum_use_multi_horizon", False))
        self.output_dir_var.set(settings.get("output_dir", str(ANALYSIS_OUTPUT_DIR)))

    def _on_strategy_change(self, _event: object) -> None:
        self._update_strategy_detail()

    def _update_strategy_detail(self) -> None:
        strategy = self._selected_strategy_key()
        if strategy == "Momentum":
            self.strategy_detail_var.set("Required data: prices (with optional volume)")
            self._set_momentum_toggles_state(enabled=True)
            return
        spec = STRATEGY_REGISTRY.get(strategy)
        if spec:
            required = ", ".join(spec.required_data)
            self.strategy_detail_var.set(f"Required data: {required}")
        else:
            self.strategy_detail_var.set("Required data: prices")
        self._set_momentum_toggles_state(enabled=False)

    def _set_momentum_toggles_state(self, *, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        if not enabled:
            self.momentum_volatility_var.set(False)
            self.momentum_residual_var.set(False)
            self.momentum_multi_horizon_var.set(False)
            self.momentum_label.grid_remove()
            self.momentum_toggles_frame.grid_remove()
        else:
            self.momentum_label.grid()
            self.momentum_toggles_frame.grid()
        for widget in (
            self.momentum_volatility_check,
            self.momentum_residual_check,
            self.momentum_multi_horizon_check,
        ):
            widget.configure(state=state)

    def run_analysis(self) -> None:
        run_started_at = datetime.now()
        run_type = self.analysis_type_var.get().strip() or "Unknown"
        if not self.controller.api_key:
            messagebox.showinfo("Missing key", "Enter a Massive API key first.")
            self._record_run(
                run_type=run_type,
                start_time=run_started_at,
                end_time=datetime.now(),
                status="failed",
                key_metrics="Missing API key",
                artifact_path="",
            )
            return
        if not self.controller.state.tickers:
            messagebox.showinfo("Missing universe", "Please add tickers first.")
            self._record_run(
                run_type=run_type,
                start_time=run_started_at,
                end_time=datetime.now(),
                status="failed",
                key_metrics="Missing ticker universe",
                artifact_path="",
            )
            return
        lookback_days = parse_int(self.lookback_var.get())
        skip_days = parse_int(self.skip_var.get())
        top_quantile = parse_float(self.top_quantile_var.get())
        bottom_quantile = parse_float(self.bottom_quantile_var.get())
        if lookback_days is None or lookback_days <= 0:
            messagebox.showinfo("Invalid input", "Lookback days must be a positive integer.")
            return
        if skip_days is None or skip_days < 0:
            messagebox.showinfo("Invalid input", "Skip days must be zero or a positive integer.")
            return
        if skip_days >= lookback_days:
            messagebox.showinfo(
                "Invalid input",
                "Skip days must be less than lookback days to compute momentum.",
            )
            return
        if top_quantile is None or not (0 < top_quantile <= 1):
            messagebox.showinfo("Invalid input", "Top quantile must be between 0 and 1.")
            return
        if bottom_quantile is None or not (0 < bottom_quantile <= 1):
            messagebox.showinfo("Invalid input", "Bottom quantile must be between 0 and 1.")
            return
        output_dir = self.output_dir_var.get().strip() or str(ANALYSIS_OUTPUT_DIR)

        strategy = self._selected_strategy_key()
        settings_payload = {
            "analysis_type": self.analysis_type_var.get(),
            "cross_sectional_strategy": strategy,
            "time_series_strategy": strategy,
            "lookback_days": lookback_days,
            "skip_days": skip_days,
            "top_quantile": top_quantile,
            "bottom_quantile": bottom_quantile,
            "momentum_use_volatility_scaling": self.momentum_volatility_var.get(),
            "momentum_use_residual": self.momentum_residual_var.get(),
            "momentum_use_multi_horizon": self.momentum_multi_horizon_var.get(),
            "time_series_use_volatility_scaling": self.momentum_volatility_var.get(),
            "time_series_use_residual": self.momentum_residual_var.get(),
            "time_series_use_multi_horizon": self.momentum_multi_horizon_var.get(),
            "time_series_use_zscore": False,
            "time_series_winsorize_sigma": None,
            "output_dir": output_dir,
        }
        self.controller.state.general_analysis_settings = settings_payload
        self.controller.persist_state()

        if self.api_client is None:
            self.api_client = MassiveApiClient(self.controller.api_key)

        as_of = effective_market_date().isoformat()
        universe = list(self.controller.state.tickers)
        price_history, fetch_skipped = self._collect_price_history(
            universe, lookback_days, skip_days
        )
        fundamentals_by_ticker, fundamentals_skipped = self._collect_fundamentals(
            universe, as_of
        )

        analysis_type = settings_payload["analysis_type"]
        reports: list[str] = []
        report_labels: list[str] = []
        if analysis_type in {"Cross-Sectional", "Cross-Sectional + Time-Series"}:
            cross_result = self._run_cross_sectional_analysis(
                strategy,
                price_history,
                fundamentals_by_ticker,
                fetch_skipped,
                fundamentals_skipped,
                lookback_days,
                skip_days,
                top_quantile,
                bottom_quantile,
                settings_payload,
                universe,
                as_of,
            )
            if cross_result is None:
                self._record_run(
                    run_type=analysis_type,
                    start_time=run_started_at,
                    end_time=datetime.now(),
                    status="failed",
                    key_metrics="Cross-sectional run failed",
                    artifact_path="",
                )
                return
            reports.append(cross_result)
            report_labels.append("cross_sectional")
        if analysis_type in {"Time-Series", "Cross-Sectional + Time-Series"}:
            time_series_result = self._run_time_series_analysis(
                strategy,
                price_history,
                fundamentals_by_ticker,
                fetch_skipped,
                fundamentals_skipped,
                lookback_days,
                skip_days,
                top_quantile,
                bottom_quantile,
                settings_payload,
                universe,
                as_of,
            )
            if time_series_result is None:
                self._record_run(
                    run_type=analysis_type,
                    start_time=run_started_at,
                    end_time=datetime.now(),
                    status="failed",
                    key_metrics="Time-series run failed",
                    artifact_path="",
                )
                return
            reports.append(time_series_result)
            report_labels.append("time_series")

        combined_report = "\n\n".join(reports)

        analysis_slug = "combined" if len(report_labels) > 1 else report_labels[0]
        output_path = self._write_report(combined_report, output_dir, strategy, analysis_slug)
        self._latest_combined_report = combined_report
        self._latest_report_path = str(output_path)
        self._record_run(
            run_type=analysis_type,
            start_time=run_started_at,
            end_time=datetime.now(),
            status="success",
            key_metrics=self._summarize_key_metrics(combined_report),
            artifact_path=str(output_path),
        )
        messagebox.showinfo(
            "Analysis complete",
            f"{analysis_type} {strategy.lower()} results written to:\n{output_path}",
        )


    def _run_backend_callable(self, fn: object, **kwargs: object) -> object:
        backend = self.controller.execution_backend
        job_id = backend.submit_job(JOB_ANALYSIS_CALLABLE, {"callable": fn, "kwargs": kwargs})
        while True:
            status = backend.get_status(job_id)
            if status in {"succeeded", "failed"}:
                break
            threading.Event().wait(0.05)
        if status == "failed":
            logs = backend.stream_logs(job_id)
            raise RuntimeError(logs[-1] if logs else "analysis backend job failed")
        return backend.get_result(job_id) if hasattr(backend, "get_result") else None

    def _summarize_key_metrics(self, combined_report: str) -> str:
        metrics: list[str] = []
        metric_patterns = [
            ("Sharpe", r"Sharpe(?:\s+Ratio)?\s*:\s*([^\n]+)"),
            ("CAGR", r"CAGR\s*:\s*([^\n]+)"),
            ("Total Return", r"Total\s+Return\s*:\s*([^\n]+)"),
            ("Max DD", r"Max\s+Drawdown\s*:\s*([^\n]+)"),
            ("Win Rate", r"Win\s+Rate\s*:\s*([^\n]+)"),
        ]
        for label, pattern in metric_patterns:
            match = re.search(pattern, combined_report, flags=re.IGNORECASE)
            if match:
                metrics.append(f"{label} {match.group(1).strip()}")
            if len(metrics) == 3:
                break
        if not metrics:
            first_line = next((line.strip() for line in combined_report.splitlines() if line.strip()), "")
            return first_line[:120] if first_line else "Report generated"
        return " | ".join(metrics)

    def _record_run(
        self,
        *,
        run_type: str,
        start_time: datetime,
        end_time: datetime,
        status: str,
        key_metrics: str,
        artifact_path: str,
    ) -> None:
        row = {
            "run_type": run_type,
            "start_time": start_time.strftime("%Y-%m-%d %H:%M:%S"),
            "end_time": end_time.strftime("%Y-%m-%d %H:%M:%S"),
            "status": status,
            "key_metrics": key_metrics,
            "artifact_path": artifact_path,
        }
        self._run_history.insert(0, row)
        self._refresh_run_tree()

    def _refresh_run_tree(self) -> None:
        for item_id in self.run_tree.get_children():
            self.run_tree.delete(item_id)
        for row in self._run_history:
            self.run_tree.insert(
                "",
                "end",
                values=(
                    row["run_type"],
                    row["start_time"],
                    row["end_time"],
                    row["status"],
                    row["key_metrics"],
                    row["artifact_path"],
                ),
            )
        self._update_run_actions_state()

    def _selected_artifact_path(self) -> str:
        selected = self.run_tree.selection()
        if not selected:
            return ""
        values = self.run_tree.item(selected[0], "values")
        return str(values[5]).strip() if len(values) > 5 else ""

    def _update_run_actions_state(self) -> None:
        artifact_path = self._selected_artifact_path()
        if artifact_path:
            self.open_artifact_button.configure(state="normal")
            self.copy_artifact_button.configure(state="normal")
            self.run_actions_var.set("Double-click a run row to open its artifact directory.")
            return
        self.open_artifact_button.configure(state="disabled")
        self.copy_artifact_button.configure(state="disabled")
        self.run_actions_var.set("Select a run to open or copy its artifact path.")

    def _open_selected_artifact_directory(self) -> None:
        artifact_path = self._selected_artifact_path()
        if not artifact_path:
            return
        artifact = Path(artifact_path).expanduser().resolve()
        target_dir = artifact if artifact.is_dir() else artifact.parent
        if not target_dir.exists():
            messagebox.showerror("Artifact missing", f"Artifact directory not found:\n{target_dir}")
            return
        webbrowser.open(f"file://{target_dir}")

    def _copy_selected_artifact_path(self) -> None:
        artifact_path = self._selected_artifact_path()
        if not artifact_path:
            return
        self.clipboard_clear()
        self.clipboard_append(artifact_path)
        self.run_actions_var.set("Artifact path copied to clipboard.")

    def export_prompt_pack(self) -> None:
        output_dir = self.output_dir_var.get().strip() or str(ANALYSIS_OUTPUT_DIR)
        prompt_dir = Path(output_dir) / "prompt_packs"
        settings_payload = dict(DEFAULT_GENERAL_ANALYSIS_SETTINGS)
        settings_payload.update(self.controller.state.general_analysis_settings or {})
        recent_outputs = {
            "latest_report_path": self._latest_report_path,
            "latest_report_excerpt": self._latest_combined_report[:4000],
            "tickers": list(self.controller.state.tickers),
        }
        output_path = write_prompt_pack(
            output_dir=prompt_dir,
            file_stem="analysis_prompt_pack",
            title="General Analysis Prompt Pack",
            config=settings_payload,
            recent_outputs=recent_outputs,
        )
        messagebox.showinfo("Prompt pack exported", f"Saved prompt pack to:\n{output_path}")

    def _selected_strategy_key(self) -> str:
        label = self.cross_sectional_var.get()
        return self.strategy_label_to_key.get(label, "Momentum")

    def _missing_data_guidance(self, missing_requirements: list[str]) -> str:
        suggestions: list[str] = []
        if "fundamentals" in missing_requirements:
            suggestions.append(
                "Fundamentals: import a fundamentals file (CSV/JSON) with fields like book value, earnings, dividends, and market cap."
            )
        if "market_cap" in missing_requirements:
            suggestions.append(
                "Market cap: compute from price * shares outstanding if shares data is available, or ingest a market-cap file."
            )
        if "earnings" in missing_requirements or "analyst_revisions" in missing_requirements:
            suggestions.append(
                "Earnings/revisions: ingest analyst estimates or earnings surprise data from a data vendor or your own CSV export."
            )
        if not suggestions:
            suggestions.append("Provide the missing data in a local file and wire it into the analysis pipeline.")
        return "\n".join(f"- {item}" for item in suggestions)

    def _run_cross_sectional_analysis(
        self,
        strategy: str,
        price_history: dict[str, list[dict]],
        fundamentals_by_ticker: dict[str, dict],
        fetch_skipped: dict[str, str],
        fundamentals_skipped: dict[str, str],
        lookback_days: int,
        skip_days: int,
        top_quantile: float,
        bottom_quantile: float,
        settings_payload: dict[str, object],
        universe: list[str],
        as_of: str,
    ) -> str | None:
        if strategy == "Momentum":
            momentum_settings = MomentumSettings(
                lookback_days=lookback_days,
                skip_days=skip_days,
                top_quantile=top_quantile,
                bottom_quantile=bottom_quantile,
                use_volatility_scaling=settings_payload["momentum_use_volatility_scaling"],
                use_residual=settings_payload["momentum_use_residual"],
                use_multi_horizon=settings_payload["momentum_use_multi_horizon"],
            )
            result = self._run_backend_callable(
                compute_cross_sectional_momentum,
                price_history=price_history,
                fundamentals_by_ticker=fundamentals_by_ticker,
                settings=momentum_settings,
            )
        else:
            spec = STRATEGY_REGISTRY.get(strategy, STRATEGY_REGISTRY["Value"])
            data_availability = {
                "fundamentals": bool(fundamentals_by_ticker),
                "market_cap": any(
                    payload.get("market_cap") is not None
                    for payload in fundamentals_by_ticker.values()
                ),
                "earnings": any(
                    payload.get("earnings_actual") is not None
                    for payload in fundamentals_by_ticker.values()
                ),
                "analyst_revisions": False,
            }
            missing_requirements = [
                requirement
                for requirement in spec.required_data
                if requirement not in {"prices", "volume"}
                and not data_availability.get(requirement, False)
            ]
            if missing_requirements:
                guidance = self._missing_data_guidance(missing_requirements)
                messagebox.showinfo(
                    "Missing data",
                    "This strategy requires additional data sources that are not yet wired: "
                    f"{', '.join(missing_requirements)}.\n\n{guidance}",
                )
                return None
            factor_settings = CrossSectionalSettings(
                top_quantile=top_quantile, bottom_quantile=bottom_quantile
            )
            result = self._run_backend_callable(
                spec.compute,
                price_history=price_history,
                fundamentals_by_ticker=fundamentals_by_ticker,
                settings=factor_settings,
            )
        result.skipped.update(fetch_skipped)
        result.skipped.update(fundamentals_skipped)

        report_title = f"Cross-Sectional {strategy} Report"
        return format_cross_sectional_report(
            title=report_title,
            as_of=as_of,
            universe=universe,
            settings=settings_payload,
            result=result,
        )

    def _run_time_series_analysis(
        self,
        strategy: str,
        price_history: dict[str, list[dict]],
        fundamentals_by_ticker: dict[str, dict],
        fetch_skipped: dict[str, str],
        fundamentals_skipped: dict[str, str],
        lookback_days: int,
        skip_days: int,
        top_quantile: float,
        bottom_quantile: float,
        settings_payload: dict[str, object],
        universe: list[str],
        as_of: str,
    ) -> str | None:
        if strategy == "Momentum":
            momentum_settings = TimeSeriesMomentumSettings(
                lookback_days=lookback_days,
                skip_days=skip_days,
                top_quantile=top_quantile,
                bottom_quantile=bottom_quantile,
                use_volatility_scaling=settings_payload["time_series_use_volatility_scaling"],
                use_residual=settings_payload["time_series_use_residual"],
                use_multi_horizon=settings_payload["time_series_use_multi_horizon"],
                use_zscore=settings_payload["time_series_use_zscore"],
                winsorize_sigma=settings_payload["time_series_winsorize_sigma"],
            )
            result = self._run_backend_callable(
                compute_time_series_momentum,
                price_history=price_history,
                fundamentals_by_ticker=fundamentals_by_ticker,
                settings=momentum_settings,
            )
        else:
            spec = TIME_SERIES_STRATEGY_REGISTRY.get(
                strategy, TIME_SERIES_STRATEGY_REGISTRY["Value"]
            )
            data_availability = {
                "fundamentals": bool(fundamentals_by_ticker),
                "market_cap": any(
                    payload.get("market_cap") is not None
                    for payload in fundamentals_by_ticker.values()
                ),
                "earnings": any(
                    payload.get("earnings_actual") is not None
                    for payload in fundamentals_by_ticker.values()
                ),
                "analyst_revisions": False,
            }
            missing_requirements = [
                requirement
                for requirement in spec.required_data
                if requirement not in {"prices", "volume"}
                and not data_availability.get(requirement, False)
            ]
            if missing_requirements:
                guidance = self._missing_data_guidance(missing_requirements)
                messagebox.showinfo(
                    "Missing data",
                    "This strategy requires additional data sources that are not yet wired: "
                    f"{', '.join(missing_requirements)}.\n\n{guidance}",
                )
                return None
            factor_settings = TimeSeriesSettings(
                lookback_days=lookback_days,
                skip_days=skip_days,
                top_quantile=top_quantile,
                bottom_quantile=bottom_quantile,
                use_volatility_scaling=settings_payload["time_series_use_volatility_scaling"],
                use_residual=settings_payload["time_series_use_residual"],
                use_multi_horizon=settings_payload["time_series_use_multi_horizon"],
                use_zscore=settings_payload["time_series_use_zscore"],
                winsorize_sigma=settings_payload["time_series_winsorize_sigma"],
            )
            result = self._run_backend_callable(
                spec.compute,
                price_history=price_history,
                fundamentals_by_ticker=fundamentals_by_ticker,
                settings=factor_settings,
            )
        result.skipped.update(fetch_skipped)
        result.skipped.update(fundamentals_skipped)

        report_title = f"Time-Series {strategy} Report"
        return format_time_series_report(
            title=report_title,
            as_of=as_of,
            universe=universe,
            settings=settings_payload,
            result=result,
        )

    def _collect_fundamentals(
        self, tickers: list[str], as_of: str
    ) -> tuple[dict[str, dict], dict[str, str]]:
        fundamentals_by_ticker: dict[str, dict] = {}
        skipped: dict[str, str] = {}
        if not tickers:
            return fundamentals_by_ticker, skipped
        max_workers = min(8, max(1, len(tickers)))
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(self._load_fundamentals, ticker, as_of) for ticker in tickers]
            for future in as_completed(futures):
                ticker, payload, reason = future.result()
                if payload and _has_fundamentals_data(payload):
                    fundamentals_by_ticker[ticker] = payload
                if reason:
                    skipped[ticker] = reason
                elif not payload or not _has_fundamentals_data(payload):
                    skipped[ticker] = "missing_fundamentals"
        return fundamentals_by_ticker, skipped

    def _load_fundamentals(
        self, ticker: str, as_of: str
    ) -> tuple[str, dict | None, str | None]:
        cache_payload = load_cached_market_data(ticker) or {}
        fundamentals_cache = cache_payload.get("fundamentals") or {}
        cached_as_of = fundamentals_cache.get("as_of")
        cached_payload = fundamentals_cache.get("payload")
        if cached_as_of == as_of and isinstance(cached_payload, dict):
            return ticker, cached_payload, None

        if self.api_client is None:
            self.api_client = MassiveApiClient(self.controller.api_key)

        errors: list[str] = []
        ticker_details: dict = {}
        financials: list[dict] = []
        dividends: list[dict] = []
        earnings: list[dict] = []

        try:
            self._throttle_request()
            ticker_details = self.api_client.fetch_ticker_details(ticker)
        except HTTPError as exc:
            errors.append(f"ticker_details_http_error_{exc.code}")
        except (TimeoutError, socket.timeout):
            errors.append("ticker_details_timeout")
        except URLError as exc:
            errors.append(f"ticker_details_url_error_{exc.reason}")

        try:
            self._throttle_request()
            financials = self.api_client.fetch_financials(ticker, period="annual")
        except HTTPError as exc:
            errors.append(f"financials_http_error_{exc.code}")
        except (TimeoutError, socket.timeout):
            errors.append("financials_timeout")
        except URLError as exc:
            errors.append(f"financials_url_error_{exc.reason}")

        try:
            self._throttle_request()
            dividends = self.api_client.fetch_dividends(ticker)
        except HTTPError as exc:
            errors.append(f"dividends_http_error_{exc.code}")
        except (TimeoutError, socket.timeout):
            errors.append("dividends_timeout")
        except URLError as exc:
            errors.append(f"dividends_url_error_{exc.reason}")

        try:
            self._throttle_request()
            earnings = self.api_client.fetch_earnings(ticker)
        except HTTPError as exc:
            if exc.code not in {403, 404}:
                errors.append(f"earnings_http_error_{exc.code}")
        except (TimeoutError, socket.timeout):
            errors.append("earnings_timeout")
        except URLError as exc:
            errors.append(f"earnings_url_error_{exc.reason}")

        payload = self._build_fundamentals_payload(
            ticker_details, financials, dividends, earnings
        )
        cache_payload["fundamentals"] = {
            "as_of": as_of,
            "payload": payload,
        }
        save_cached_market_data(ticker, cache_payload)
        reason = errors[0] if errors else None
        if reason is None and not _has_fundamentals_data(payload):
            reason = "missing_fundamentals"
        return ticker, payload, reason

    def _build_fundamentals_payload(
        self,
        ticker_details: dict,
        financials: list[dict],
        dividends: list[dict],
        earnings: list[dict],
    ) -> dict:
        financials_record = self._select_latest_by_date(
            financials,
            ["fiscal_period_end_date", "reporting_date", "filing_date", "calendar_date"],
        )
        financials_record = financials_record or {}

        book_value = _get_nested_value(
            financials_record,
            [
                ("financials", "balance_sheet", "equity"),
                ("financials", "balance_sheet", "stockholders_equity"),
                ("financials", "balance_sheet", "total_equity"),
                ("financials", "balance_sheet", "shareholders_equity"),
            ],
        )
        total_assets = _get_nested_value(
            financials_record,
            [
                ("financials", "balance_sheet", "assets"),
                ("financials", "balance_sheet", "total_assets"),
            ],
        )
        total_equity = _get_nested_value(
            financials_record,
            [
                ("financials", "balance_sheet", "equity"),
                ("financials", "balance_sheet", "total_equity"),
                ("financials", "balance_sheet", "stockholders_equity"),
            ],
        )
        net_income = _get_nested_value(
            financials_record,
            [
                ("financials", "income_statement", "net_income_loss"),
                ("financials", "income_statement", "net_income"),
            ],
        )
        eps = _get_nested_value(
            financials_record,
            [
                ("financials", "income_statement", "basic_earnings_per_share"),
                ("financials", "income_statement", "diluted_earnings_per_share"),
                ("financials", "income_statement", "eps"),
            ],
        )
        shares_outstanding = (
            _coerce_number(ticker_details.get("weighted_shares_outstanding"))
            or _coerce_number(ticker_details.get("share_class_shares_outstanding"))
            or _get_nested_value(
                financials_record,
                [
                    ("financials", "income_statement", "weighted_average_shares_outstanding"),
                    ("financials", "income_statement", "weighted_average_shares_outstanding_diluted"),
                ],
            )
        )
        market_cap = _coerce_number(ticker_details.get("market_cap"))

        dividends_ttm, last_dividend_date, last_dividend_amount = self._summarize_dividends(
            dividends
        )
        earnings_record = self._select_latest_by_date(
            earnings,
            ["reporting_date", "fiscal_period_end_date", "publish_date", "date"],
        )
        earnings_record = earnings_record or {}
        earnings_actual = _coerce_number(
            earnings_record.get("actual_eps")
            or earnings_record.get("eps_actual")
            or earnings_record.get("eps")
        )
        earnings_estimate = _coerce_number(
            earnings_record.get("estimated_eps") or earnings_record.get("eps_estimate")
        )
        if earnings_actual is None:
            earnings_actual = eps
        earnings_surprise = (
            earnings_actual - earnings_estimate
            if earnings_actual is not None and earnings_estimate is not None
            else None
        )

        return {
            "market_cap": market_cap,
            "book_value": book_value or total_equity,
            "total_assets": total_assets,
            "total_equity": total_equity,
            "net_income": net_income,
            "eps": eps,
            "dividends_ttm": dividends_ttm,
            "shares_outstanding": shares_outstanding,
            "earnings_actual": earnings_actual,
            "earnings_estimate": earnings_estimate,
            "earnings_surprise": earnings_surprise,
            "last_ex_dividend_date": last_dividend_date.isoformat()
            if last_dividend_date
            else None,
            "last_dividend_amount": last_dividend_amount,
        }

    def _select_latest_by_date(
        self, records: list[dict], date_fields: list[str]
    ) -> dict | None:
        best_record: dict | None = None
        best_date: date | None = None
        for record in records:
            record_date: date | None = None
            for field in date_fields:
                record_date = _parse_iso_date(record.get(field))
                if record_date:
                    break
            if record_date and (best_date is None or record_date > best_date):
                best_date = record_date
                best_record = record
        if best_record is None and records:
            return records[0]
        return best_record

    def _summarize_dividends(
        self, dividends: list[dict]
    ) -> tuple[float | None, date | None, float | None]:
        as_of_date = effective_market_date()
        ttm_start = as_of_date - timedelta(days=365)
        total = 0.0
        total_found = False
        last_date: date | None = None
        last_amount: float | None = None
        for record in dividends:
            ex_date = _parse_iso_date(
                record.get("ex_dividend_date")
                or record.get("ex_date")
                or record.get("pay_date")
                or record.get("payment_date")
            )
            amount = _coerce_number(
                record.get("cash_amount") or record.get("dividend") or record.get("amount")
            )
            if ex_date and amount is not None:
                if ex_date >= ttm_start:
                    total += amount
                    total_found = True
                if last_date is None or ex_date > last_date:
                    last_date = ex_date
                    last_amount = amount
        dividends_ttm = total if total_found else None
        return dividends_ttm, last_date, last_amount

    def _collect_price_history(
        self, tickers: list[str], lookback_days: int, skip_days: int
    ) -> tuple[dict[str, list[float]], dict[str, str]]:
        min_points = lookback_days + skip_days + 1
        buffer_days = max(10, int(min_points * 1.5))
        days_back = min_points + buffer_days
        as_of = effective_market_date().isoformat()

        prices_by_ticker: dict[str, list[float]] = {}
        skipped: dict[str, str] = {}

        if len(tickers) >= 200:
            if self.api_client is None:
                self.api_client = MassiveApiClient(self.controller.api_key)
            grouped = [ticker for ticker in tickers if self._is_grouped_eligible(ticker)]
            special = [ticker for ticker in tickers if ticker not in grouped]
            prices_by_ticker, skipped = self._collect_grouped_history(grouped, min_points)
            if special:
                special_prices, special_skipped = self._collect_ticker_history(
                    special, days_back, min_points, as_of
                )
                prices_by_ticker.update(special_prices)
                skipped.update(special_skipped)
            fallback_tickers = [
                ticker
                for ticker, reason in skipped.items()
                if reason == "insufficient_history" and ticker not in special
            ]
            if fallback_tickers:
                fallback_prices, fallback_skipped = self._collect_ticker_history(
                    fallback_tickers, days_back, min_points, as_of
                )
                prices_by_ticker.update(fallback_prices)
                for ticker in fallback_tickers:
                    skipped.pop(ticker, None)
                skipped.update(fallback_skipped)
            return prices_by_ticker, skipped

        return self._collect_ticker_history(tickers, days_back, min_points, as_of)

    def _collect_ticker_history(
        self, tickers: list[str], days_back: int, min_points: int, as_of: str
    ) -> tuple[dict[str, list[float]], dict[str, str]]:
        prices_by_ticker: dict[str, list[float]] = {}
        skipped: dict[str, str] = {}
        max_workers = min(12, max(1, len(tickers)))
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [
                executor.submit(
                    self._load_daily_closes, ticker, days_back, min_points, as_of
                )
                for ticker in tickers
            ]
            for future in as_completed(futures):
                ticker, prices, reason = future.result()
                if prices:
                    prices_by_ticker[ticker] = prices
                else:
                    skipped[ticker] = reason or "no_data"
        return prices_by_ticker, skipped

    def _collect_grouped_history(
        self, tickers: list[str], min_points: int
    ) -> tuple[dict[str, list[float]], dict[str, str]]:
        if self.api_client is None:
            self.api_client = MassiveApiClient(self.controller.api_key)
        universe = set(tickers)
        closes: dict[str, list[float]] = {ticker: [] for ticker in tickers}
        skipped: dict[str, str] = {}
        pending = set(tickers)

        end_date = effective_market_date()
        cache_dir = DATA_DIR / "grouped_cache" / end_date.isoformat()
        cache_dir.mkdir(parents=True, exist_ok=True)
        max_calendar_days = max(30, min_points * 4)
        current_date = end_date
        days_checked = 0

        retry_count = 0
        max_retries = 4
        backoff_seconds = 1.0

        while pending and days_checked < max_calendar_days:
            if current_date.weekday() >= 5:
                current_date -= timedelta(days=1)
                days_checked += 1
                continue

            try:
                day_cache_path = cache_dir / f"{current_date}.json"
                if day_cache_path.exists():
                    day_payload = json.loads(day_cache_path.read_text())
                    if day_payload:
                        sample_value = next(iter(day_payload.values()))
                        if isinstance(sample_value, (int, float)):
                            day_payload = {
                                ticker: {"close": float(value), "volume": None}
                                for ticker, value in day_payload.items()
                            }
                else:
                    self._throttle_request()
                    aggregates = self.api_client.fetch_grouped_daily_aggregates(current_date)
                    day_payload = {}
                    for item in aggregates:
                        ticker = item.get("T")
                        if ticker not in universe:
                            continue
                        close_value = item.get("c")
                        volume_value = item.get("v")
                        if isinstance(close_value, (int, float)) and close_value > 0:
                            day_payload[ticker] = {
                                "close": float(close_value),
                                "volume": float(volume_value)
                                if isinstance(volume_value, (int, float))
                                else None,
                            }
                    day_cache_path.write_text(json.dumps(day_payload))
            except HTTPError as exc:
                if exc.code == 429 and retry_count < max_retries:
                    retry_after = exc.headers.get("Retry-After") if exc.headers else None
                    wait_seconds = backoff_seconds
                    if retry_after:
                        try:
                            wait_seconds = max(wait_seconds, float(retry_after))
                        except ValueError:
                            pass
                    time.sleep(wait_seconds)
                    retry_count += 1
                    backoff_seconds *= 2
                    continue
                for ticker in pending:
                    skipped[ticker] = f"http_error_{exc.code}"
                break
            except (TimeoutError, socket.timeout):
                if retry_count < max_retries:
                    time.sleep(backoff_seconds)
                    retry_count += 1
                    backoff_seconds *= 2
                    continue
                for ticker in pending:
                    skipped[ticker] = "timeout"
                break
            except URLError as exc:
                for ticker in pending:
                    skipped[ticker] = f"url_error_{exc.reason}"
                break

            retry_count = 0
            backoff_seconds = 3.0
            for ticker, payload in day_payload.items():
                if ticker in pending:
                    closes[ticker].append(payload)
                    if len(closes[ticker]) >= min_points:
                        pending.discard(ticker)

            current_date -= timedelta(days=1)
            days_checked += 1

        prices_by_ticker: dict[str, list[float]] = {}
        for ticker, values in closes.items():
            if len(values) >= min_points:
                prices_by_ticker[ticker] = list(reversed(values))
            elif ticker not in skipped:
                skipped[ticker] = "insufficient_history"

        return prices_by_ticker, skipped

    def _is_grouped_eligible(self, ticker: str) -> bool:
        return bool(self._grouped_ticker_pattern.fullmatch(ticker))

    def _load_daily_closes(
        self, ticker: str, days_back: int, min_points: int, as_of: str
    ) -> tuple[str, list[float] | None, str | None]:
        cache_payload = load_cached_market_data(ticker) or {}
        daily_cache = cache_payload.get("daily_closes") or {}
        cached_as_of = daily_cache.get("as_of")
        cached_prices = daily_cache.get("prices")
        if (
            cached_as_of == as_of
            and isinstance(cached_prices, list)
            and len(cached_prices) >= min_points
        ):
            if cached_prices and isinstance(cached_prices[0], dict):
                return ticker, cached_prices, None
            return ticker, [{"close": float(value), "volume": None} for value in cached_prices], None

        if self.api_client is None:
            self.api_client = MassiveApiClient(self.controller.api_key)

        current_days_back = days_back
        max_days_back = max(days_back * 3, min_points * 4)
        closes: list[dict] = []

        retry_count = 0
        max_retries = 4
        backoff_seconds = 1.0
        while current_days_back <= max_days_back:
            self._throttle_request()
            try:
                aggregates = self.api_client.fetch_daily_aggregates(
                    ticker, days_back=current_days_back
                )
            except HTTPError as exc:
                if exc.code == 429 and retry_count < max_retries:
                    retry_after = exc.headers.get("Retry-After") if exc.headers else None
                    wait_seconds = backoff_seconds
                    if retry_after:
                        try:
                            wait_seconds = max(wait_seconds, float(retry_after))
                        except ValueError:
                            pass
                    time.sleep(wait_seconds)
                    retry_count += 1
                    backoff_seconds *= 2
                    continue
                return ticker, None, f"http_error_{exc.code}"
            except (TimeoutError, socket.timeout):
                if retry_count < max_retries:
                    time.sleep(backoff_seconds)
                    retry_count += 1
                    backoff_seconds *= 2
                    continue
                return ticker, None, "timeout"
            except URLError as exc:
                return ticker, None, f"url_error_{exc.reason}"

            closes = []
            for item in sorted(aggregates, key=lambda row: row.get("t", 0)):
                close_value = item.get("c")
                volume_value = item.get("v")
                if isinstance(close_value, (int, float)) and close_value > 0:
                    closes.append(
                        {
                            "close": float(close_value),
                            "volume": float(volume_value)
                            if isinstance(volume_value, (int, float))
                            else None,
                        }
                    )

            if len(closes) >= min_points:
                break
            current_days_back = int(current_days_back * 1.5) + 1

        if closes:
            cache_payload["daily_closes"] = {
                "as_of": as_of,
                "prices": closes,
                "days_back": current_days_back,
            }
            save_cached_market_data(ticker, cache_payload)

        if len(closes) < min_points:
            return ticker, None, "insufficient_history"
        return ticker, closes, None

    def _throttle_request(self) -> None:
        with self._rate_lock:
            now = time.monotonic()
            elapsed = now - self._last_request_time
            if elapsed < self._min_request_interval:
                time.sleep(self._min_request_interval - elapsed)
            self._last_request_time = time.monotonic()

    def _write_report(
        self, report: str, output_dir: str, strategy: str, analysis_slug: str
    ) -> Path:
        directory = Path(output_dir)
        directory.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        strategy_slug = re.sub(r"[^a-z0-9]+", "_", strategy.strip().lower()).strip("_")
        output_path = directory / f"{analysis_slug}_{strategy_slug}_{timestamp}.txt"
        output_path.write_text(report)
        return output_path

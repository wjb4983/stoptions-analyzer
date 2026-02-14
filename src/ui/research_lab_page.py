from __future__ import annotations

import threading
import shutil
import json
import hashlib
import re
import uuid
import statistics
import webbrowser
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable

from backtesting.cache_runner import (
    CancellationToken,
    TaskCancellationError,
    run_multi_signal_backtest,
    run_strategy_optimization,
    run_walk_forward_backtest,
)
from backtesting.chain_runner import build_default_research_execution_chain
from config import (
    BACKTEST_OUTPUT_DIR,
    CONFIG_DIR,
    DEFAULT_BACKTEST_SETTINGS,
    DEFAULT_HYPOTHESIS_RUBRIC_TEMPLATES,
    HYPOTHESIS_RUBRIC_TEMPLATES_PATH,
    RESEARCH_LAB_PRESETS_PATH,
)
from ui.workflow_preset_validator import validate_workflow_preset_payload
from utils.parsing import normalize_cache_root, parse_date, parse_float


@dataclass
class ResearchWorkflowConfig:
    preset_name: str
    entry_signals: list[str]
    exit_signals: list[str]
    optimization_n_trials: int
    optimization_sampler: str
    optimization_enable_pruning: bool
    optimization_prune_on_constraint: bool
    optimization_prune_on_lcb: bool
    optimization_min_completed_for_pruning: int
    optimization_staged_budgets: list[dict[str, object]]
    train_fraction: float
    validation_fraction: float
    test_fraction: float
    step_fraction: float
    walk_forward_split_policy: str
    stress_controls: dict[str, object]
    benchmark_selection: list[str]


DEFAULT_RESEARCH_WORKFLOW_PRESETS: dict[str, Any] = {
    "default_preset": "balanced_baseline",
    "presets": {
        "balanced_baseline": {
            "label": "Balanced Baseline",
            "description": "General-purpose baseline with diversified signals and balanced stress settings.",
            "entry_signals": ["ts_momentum", "breakout"],
            "exit_signals": ["none", "momentum_flip"],
            "optimization": {"n_trials": 20, "sampler": "tpe", "enable_pruning": True, "prune_on_constraint": True, "prune_on_lcb": True, "min_completed_for_pruning": 5, "staged_budgets": [{"label": "coarse", "n_trials": 12, "sampler": "random", "partial_period_fractions": [0.33, 0.66]}, {"label": "fine", "n_trials": 20, "sampler": "tpe", "partial_period_fractions": [0.5, 1.0]}]},
            "walk_forward": {
                "train_fraction": 0.70,
                "validation_fraction": 0.15,
                "test_fraction": 0.15,
                "step_fraction": 0.15,
                "split_policy": "calendar-based",
            },
            "benchmark_selection": ["buy_hold", "equal_weight_momentum", "volatility_parity"],
            "stress_controls": {
                "enable_historical_replay_regimes": True,
                "historical_window_fraction": 0.20,
                "historical_replay_window_bars": 20,
                "synthetic_jump_magnitude": 0.02,
                "synthetic_jump_interval": 7,
                "synthetic_vol_cluster_multiplier": 1.6,
                "overlay_spread_multiplier": 2.5,
                "overlay_liquidity_multiplier": 0.4,
            },
        },
    },
}


WIZARD_STATE_SCHEMA_VERSION = 2


@dataclass
class ResearchTask:
    task_id: str
    label: str
    target: Callable[[dict[str, Any], ResearchWorkflowConfig, CancellationToken], str]
    context: dict[str, Any]
    config: ResearchWorkflowConfig
    state: str = "queued"
    logs: list[str] | None = None
    cancel_requested: bool = False
    cancellation_token: CancellationToken | None = None
    cancellation_reason: str | None = None
    cancellation_confirmed: bool = False
    research_pack_path: str | None = None
    workflow_output: str | None = None
    priority: int = 1
    enqueue_order: int = 0

    def __post_init__(self) -> None:
        if self.logs is None:
            self.logs = []
        if self.cancellation_token is None:
            self.cancellation_token = CancellationToken()


class ResearchLabPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller
        self._task_queue: list[ResearchTask] = []
        self._active_task_ids: set[str] = set()
        self._task_enqueue_counter = 0
        self._max_concurrent_jobs = 1
        self._research_lab_dir = BACKTEST_OUTPUT_DIR / "research_lab"
        self._sampler_options = ("tpe", "cma-es", "random", "grid")
        self._signal_options = ("ts_momentum", "ma_trend", "breakout")
        self._exit_signal_options = ("none", "momentum_flip", "trailing_stop", "max_hold")
        self._benchmark_options = ("buy_hold", "equal_weight_momentum", "volatility_parity")
        self._wizard_state_path = self._research_lab_dir / "wizard_state.json"
        self._hypothesis_rubric_templates_path = HYPOTHESIS_RUBRIC_TEMPLATES_PATH
        self._rubric_templates = self._load_rubric_templates()
        self._workflow_presets = self._load_workflow_presets()
        self._wizard_comments: list[dict[str, str]] = []
        self._wizard_history: list[dict[str, str]] = []


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
        self._build_tile(
            tiles,
            row=2,
            column=0,
            title="Hypothesis intake pipeline",
            description=(
                "Capture idea intake, rationale template, data requirements, test design, results "
                "review, and promotion/rejection decisions with rubric scoring and funnel metrics."
            ),
            action_label="Run Intake Pipeline",
            action=self.run_hypothesis_pipeline,
        )

        self._build_workflow_controls()
        self._build_wizard_mode()
        self._build_governance_mini_dashboard()
        self._build_funnel_kpi_dashboard()
        self._refresh_funnel_kpi_dashboard()

        button_row = ttk.Frame(self)
        button_row.pack(pady=(16, 8))
        ttk.Button(
            button_row,
            text="Back to Main Menu",
            command=lambda: self.controller.show_frame("MainMenu"),
        ).pack()

        self._build_task_queue_controls()

        output_frame = ttk.LabelFrame(self, text="Research Lab Output")
        output_frame.pack(fill="both", expand=True, padx=40, pady=(4, 20))
        output_controls = ttk.Frame(output_frame)
        output_controls.pack(fill="x", padx=8, pady=(8, 0))
        ttk.Button(
            output_controls,
            text="Export Research Pack",
            command=self.export_research_pack,
        ).pack(side="right")
        self.output_text = tk.Text(output_frame, height=8, wrap="word")
        self.output_text.pack(fill="both", expand=True, padx=8, pady=8)

        task_logs_frame = ttk.LabelFrame(self, text="Selected Task Logs")
        task_logs_frame.pack(fill="both", expand=True, padx=40, pady=(0, 20))
        self.task_logs_text = tk.Text(task_logs_frame, height=8, wrap="word")
        self.task_logs_text.pack(fill="both", expand=True, padx=8, pady=8)

    def _build_governance_mini_dashboard(self) -> None:
        frame = ttk.LabelFrame(self, text="Governance Mini-Dashboard")
        frame.pack(fill="x", padx=40, pady=(4, 8))
        frame.columnconfigure(1, weight=1)

        self._governance_gate_counts_var = tk.StringVar(value="No run manifest loaded.")
        self._governance_missing_checks_var = tk.StringVar(value="No run manifest loaded.")
        self._governance_promotion_ready_var = tk.StringVar(value="No run manifest loaded.")
        self._governance_approval_status_var = tk.StringVar(value="No run manifest loaded.")

        ttk.Label(frame, text="Gate pass/fail counts").grid(row=0, column=0, sticky="w", padx=8, pady=4)
        ttk.Label(frame, textvariable=self._governance_gate_counts_var).grid(row=0, column=1, sticky="w", padx=8, pady=4)

        ttk.Label(frame, text="Missing required checks").grid(row=1, column=0, sticky="w", padx=8, pady=4)
        ttk.Label(frame, textvariable=self._governance_missing_checks_var, wraplength=780, justify="left").grid(
            row=1,
            column=1,
            sticky="w",
            padx=8,
            pady=4,
        )

        ttk.Label(frame, text="Promotion readiness").grid(row=2, column=0, sticky="w", padx=8, pady=4)
        ttk.Label(frame, textvariable=self._governance_promotion_ready_var).grid(row=2, column=1, sticky="w", padx=8, pady=4)

        ttk.Label(frame, text="Approval status").grid(row=3, column=0, sticky="w", padx=8, pady=(4, 8))
        ttk.Label(frame, textvariable=self._governance_approval_status_var).grid(row=3, column=1, sticky="w", padx=8, pady=(4, 8))

        ttk.Separator(frame, orient="horizontal").grid(row=4, column=0, columnspan=2, sticky="ew", padx=8, pady=(4, 8))
        ttk.Label(frame, text="Pipeline graph").grid(row=5, column=0, sticky="nw", padx=8, pady=(0, 6))

        graph_frame = ttk.Frame(frame)
        graph_frame.grid(row=5, column=1, sticky="ew", padx=8, pady=(0, 8))
        graph_frame.columnconfigure(0, weight=1)

        self._pipeline_graph_items_var = tk.StringVar(value=[])
        self._pipeline_graph_listbox = tk.Listbox(
            graph_frame,
            listvariable=self._pipeline_graph_items_var,
            height=6,
            exportselection=False,
        )
        self._pipeline_graph_listbox.grid(row=0, column=0, columnspan=2, sticky="ew")
        self._pipeline_graph_listbox.bind("<<ListboxSelect>>", lambda _event: self._refresh_pipeline_graph_node_details())

        self._pipeline_graph_node_details_var = tk.StringVar(value="No lineage graph loaded.")
        ttk.Label(graph_frame, textvariable=self._pipeline_graph_node_details_var, wraplength=680, justify="left").grid(
            row=1,
            column=0,
            columnspan=2,
            sticky="w",
            pady=(4, 6),
        )
        ttk.Button(graph_frame, text="Open Artifact", command=lambda: self._open_pipeline_graph_reference("artifact_path")).grid(
            row=2,
            column=0,
            sticky="w",
        )
        ttk.Button(graph_frame, text="Open Logs", command=lambda: self._open_pipeline_graph_reference("logs_path")).grid(
            row=2,
            column=1,
            sticky="w",
            padx=(6, 0),
        )

        self._pipeline_graph_payload: dict[str, Any] = {"nodes": [], "edges": []}
        self._pipeline_graph_nodes_by_id: dict[str, dict[str, Any]] = {}

    def _build_funnel_kpi_dashboard(self) -> None:
        frame = ttk.LabelFrame(self, text="Idea Funnel KPI Dashboard")
        frame.pack(fill="x", padx=40, pady=(0, 8))
        frame.columnconfigure(1, weight=1)

        self._funnel_acceptance_var = tk.StringVar(value="Acceptance rate: n/a")
        self._funnel_median_time_var = tk.StringVar(value="Median time-to-decision: n/a")
        self._funnel_false_positive_var = tk.StringVar(value="False-positive proxy: n/a")
        self._funnel_strategy_rates_var = tk.StringVar(value="No strategy-family funnel data yet.")
        self._funnel_monthly_conversion_var = tk.StringVar(value="No monthly promotion conversion data yet.")

        ttk.Label(frame, textvariable=self._funnel_acceptance_var).grid(row=0, column=0, sticky="w", padx=8, pady=(6, 4))
        ttk.Label(frame, textvariable=self._funnel_median_time_var).grid(row=0, column=1, sticky="w", padx=8, pady=(6, 4))
        ttk.Label(frame, textvariable=self._funnel_false_positive_var).grid(row=1, column=0, sticky="w", padx=8, pady=(0, 6))

        ttk.Label(frame, text="Pass rates by strategy family").grid(row=2, column=0, sticky="nw", padx=8, pady=(0, 4))
        ttk.Label(frame, textvariable=self._funnel_strategy_rates_var, justify="left").grid(
            row=2,
            column=1,
            sticky="w",
            padx=8,
            pady=(0, 4),
        )

        ttk.Label(frame, text="Promotion conversion by month").grid(row=3, column=0, sticky="nw", padx=8, pady=(0, 8))
        ttk.Label(frame, textvariable=self._funnel_monthly_conversion_var, justify="left").grid(
            row=3,
            column=1,
            sticky="w",
            padx=8,
            pady=(0, 8),
        )

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

    def _build_task_queue_controls(self) -> None:
        frame = ttk.LabelFrame(self, text="Task Queue")
        frame.pack(fill="x", padx=40, pady=(2, 8))
        frame.columnconfigure(0, weight=1)

        self._task_list_var = tk.StringVar(value=[])
        self._task_listbox = tk.Listbox(frame, listvariable=self._task_list_var, height=6, exportselection=False)
        self._task_listbox.grid(row=0, column=0, sticky="ew", padx=8, pady=8)
        self._task_listbox.bind("<<ListboxSelect>>", lambda _event: self._refresh_selected_task_logs())

        controls = ttk.Frame(frame)
        controls.grid(row=0, column=1, sticky="ns", padx=(0, 8), pady=8)
        self._cancel_task_button = ttk.Button(controls, text="Cancel Task", command=self._cancel_selected_task)
        self._cancel_task_button.pack(fill="x", pady=(0, 6))
        self._retry_task_button = ttk.Button(controls, text="Retry Task", command=self._retry_selected_task)
        self._retry_task_button.pack(fill="x")

        ttk.Label(controls, text="Priority").pack(anchor="w", pady=(8, 0))
        self._task_priority_var = tk.StringVar(value="normal")
        ttk.Combobox(
            controls,
            textvariable=self._task_priority_var,
            values=("high", "normal", "low"),
            state="readonly",
            width=10,
        ).pack(fill="x", pady=(2, 6))

        ttk.Label(controls, text="Max concurrent").pack(anchor="w")
        self._max_concurrent_jobs_var = tk.IntVar(value=1)
        self._max_concurrent_jobs_var.trace_add("write", lambda *_: self._on_max_concurrent_jobs_changed())
        ttk.Spinbox(controls, from_=1, to=8, textvariable=self._max_concurrent_jobs_var, width=8).pack(fill="x")

        self._refresh_task_queue_ui()

    def _get_task_by_id(self, task_id: str | None) -> ResearchTask | None:
        if task_id is None:
            return None
        for task in self._task_queue:
            if task.task_id == task_id:
                return task
        return None

    def _selected_task(self) -> ResearchTask | None:
        selection = self._task_listbox.curselection()
        if not selection:
            return None
        index = int(selection[0])
        if index < 0 or index >= len(self._task_queue):
            return None
        return self._task_queue[index]

    def _refresh_task_queue_ui(self) -> None:
        rows = [f"[{task.state}|p{task.priority}] {task.label}" for task in self._task_queue]
        self._task_list_var.set(rows)
        task = self._selected_task()
        self._cancel_task_button.configure(state="normal" if task and task.state in {"queued", "running", "canceling"} else "disabled")
        self._retry_task_button.configure(state="normal" if task and task.state in {"failed", "canceled"} else "disabled")
        self._refresh_selected_task_logs()
        if hasattr(self, "_wizard_next_button"):
            self._wizard_refresh_nav_state()

    def _refresh_selected_task_logs(self) -> None:
        if not hasattr(self, "task_logs_text"):
            return
        task = self._selected_task()
        self.task_logs_text.delete("1.0", tk.END)
        if not task or task.logs is None:
            return
        for line in task.logs:
            self.task_logs_text.insert(tk.END, f"{line}\n")
        self.task_logs_text.see(tk.END)

    def _task_log(self, task: ResearchTask, message: str) -> None:
        task.logs.append(message)
        self._append_output(f"[{task.label}] {message}")
        self._refresh_task_queue_ui()

    def _schedule_ui_update(self, callback: Callable[[], None]) -> None:
        self.after(0, callback)

    def _enqueue_task(
        self,
        *,
        label: str,
        target: Callable[[dict[str, Any], ResearchWorkflowConfig, CancellationToken], str],
        context: dict[str, Any],
        config: ResearchWorkflowConfig,
    ) -> None:
        priority = self._priority_value_from_label(getattr(self, "_task_priority_var", None).get() if hasattr(self, "_task_priority_var") else "normal")
        self._task_enqueue_counter += 1
        task_id = uuid.uuid4().hex
        task_context = dict(context)
        cache_root = Path(task_context.get("cache_root", BACKTEST_OUTPUT_DIR))
        task_context["cache_root"] = cache_root / "research_lab_tasks" / task_id
        task_context["run_namespace"] = f"task_{task_id}"
        task = ResearchTask(
            task_id=task_id,
            label=label,
            target=target,
            context=task_context,
            config=config,
            priority=priority,
            enqueue_order=self._task_enqueue_counter,
        )
        self._task_queue.append(task)
        self._refresh_task_queue_ui()
        self._schedule_tasks()

    def _priority_value_from_label(self, label: str) -> int:
        normalized = str(label).strip().lower()
        return {"low": 0, "normal": 1, "high": 2}.get(normalized, 1)

    def _on_max_concurrent_jobs_changed(self) -> None:
        if hasattr(self, "_max_concurrent_jobs_var"):
            self._max_concurrent_jobs = max(1, int(self._max_concurrent_jobs_var.get() or 1))
        self._schedule_tasks()

    def _schedule_tasks(self) -> None:
        available_slots = max(0, int(self._max_concurrent_jobs) - len(self._active_task_ids))
        if available_slots <= 0:
            return
        queued_tasks = sorted(
            [task for task in self._task_queue if task.state == "queued"],
            key=lambda task: (-int(task.priority), int(task.enqueue_order)),
        )
        for next_task in queued_tasks[:available_slots]:
            self._start_task_worker(next_task)

    def _start_task_worker(self, next_task: ResearchTask) -> None:
        self._active_task_ids.add(next_task.task_id)
        next_task.state = "running"
        self._task_log(next_task, "Task started.")

        def worker(task: ResearchTask) -> None:
            if task.cancel_requested:
                self._schedule_ui_update(lambda: self._finish_task(task.task_id, "", canceled=True))
                return
            try:
                output = task.target(task.context, task.config, task.cancellation_token or CancellationToken())
                self._schedule_ui_update(lambda: self._finish_task(task.task_id, output))
            except TaskCancellationError as exc:
                self._schedule_ui_update(lambda: self._finish_task(task.task_id, str(exc), canceled=True))
            except Exception as exc:
                self._schedule_ui_update(lambda: self._finish_task(task.task_id, f"Research workflow failed: {exc}", failed=True))

        threading.Thread(target=worker, args=(next_task,), daemon=True).start()

    def _finish_task(self, task_id: str, output: str, *, failed: bool = False, canceled: bool = False) -> None:
        task = self._get_task_by_id(task_id)
        if task is None:
            return
        output_text = str(output)
        if canceled or task.cancel_requested:
            task.state = "canceled"
            task.cancellation_confirmed = True
            task.logs.append("Task canceled.")
            self._append_output(f"[{task.label}] Task canceled.")
        elif failed:
            task.state = "failed"
            task.logs.append(output_text)
            self._append_output(f"[{task.label}] Failed: {output_text}")
        else:
            task.state = "succeeded"
            task.logs.append(output_text)
            task.workflow_output = output_text
            self._append_output(f"[{task.label}] Succeeded.")
            self._refresh_governance_dashboard_from_output(output_text)
            self._refresh_pipeline_graph_from_output(output_text)
            self._append_explainability_cards(task, output_text)
            pack_path = self._emit_research_pack(task, output_text)
            if pack_path is not None:
                task.research_pack_path = str(pack_path)
                message = f"Research pack exported: {pack_path}"
                task.logs.append(message)
                self._append_output(f"[{task.label}] {message}")
        self._active_task_ids.discard(task_id)
        self._refresh_task_queue_ui()
        self._schedule_tasks()

    def _append_explainability_cards(self, task: ResearchTask, output_text: str) -> None:
        manifest_path = self._extract_manifest_path_from_output(output_text)
        run_dir = manifest_path.parent if manifest_path is not None else None
        if run_dir is None:
            return

        explain_path = run_dir / "trade_explainability.json"
        if not explain_path.exists():
            return
        try:
            explain_rows = json.loads(explain_path.read_text(encoding="utf-8"))
        except Exception:
            return
        if not isinstance(explain_rows, list) or not explain_rows:
            return

        regime_by_timestamp = self._load_regime_by_timestamp(run_dir)
        loss_card = self._build_top_loss_contributors_card(explain_rows)
        failure_card = self._build_regime_failure_windows_card(explain_rows, regime_by_timestamp)
        slippage_card = self._build_slippage_sensitivity_card(explain_rows)

        cards = [
            f"[{task.label}] Explainability highlights:",
            f"• {loss_card}",
            f"• {failure_card}",
            f"• {slippage_card}",
        ]
        for line in cards:
            self._append_output(line)
            task.logs.append(line)

    def _load_regime_by_timestamp(self, run_dir: Path) -> dict[str, str]:
        regimes_path = run_dir / "regimes.json"
        if not regimes_path.exists():
            return {}
        try:
            payload = json.loads(regimes_path.read_text(encoding="utf-8"))
        except Exception:
            return {}
        if not isinstance(payload, list):
            return {}
        mapping: dict[str, str] = {}
        for row in payload:
            if not isinstance(row, dict):
                continue
            ts = str(row.get("timestamp", "")).strip()
            regime = str(row.get("regime", "")).strip() or "unknown"
            if ts:
                mapping[ts] = regime
        return mapping

    def _build_top_loss_contributors_card(self, explain_rows: list[dict[str, Any]]) -> str:
        contributors: dict[str, float] = {}
        for row in explain_rows:
            top_drivers = row.get("top_drivers", [])
            if not isinstance(top_drivers, list):
                continue
            for driver in top_drivers:
                if not isinstance(driver, dict):
                    continue
                feature = str(driver.get("feature", "unknown")).strip() or "unknown"
                score = float(driver.get("attribution_score", 0.0))
                if score < 0.0:
                    contributors[feature] = contributors.get(feature, 0.0) + abs(score)
        if not contributors:
            return "Top loss contributors: none detected from negative attribution drivers."
        ranked = sorted(contributors.items(), key=lambda item: item[1], reverse=True)[:3]
        summary = ", ".join(f"{feature} ({value:.2f})" for feature, value in ranked)
        return f"Top loss contributors: {summary}."

    def _build_regime_failure_windows_card(
        self,
        explain_rows: list[dict[str, Any]],
        regime_by_timestamp: dict[str, str],
    ) -> str:
        windows_by_regime: dict[str, list[tuple[str, str, int]]] = {}
        current_regime: str | None = None
        start_ts = ""
        end_ts = ""
        current_count = 0

        def flush() -> None:
            nonlocal current_regime, start_ts, end_ts, current_count
            if current_regime is None or current_count <= 0:
                return
            windows_by_regime.setdefault(current_regime, []).append((start_ts, end_ts, current_count))
            current_regime = None
            start_ts = ""
            end_ts = ""
            current_count = 0

        for row in explain_rows:
            flags = row.get("red_flags", [])
            ts = str(row.get("timestamp", "")).strip()
            regime = regime_by_timestamp.get(ts, "unknown")
            flagged = isinstance(flags, list) and len(flags) > 0
            if not flagged:
                flush()
                continue
            if current_regime == regime:
                end_ts = ts
                current_count += 1
            else:
                flush()
                current_regime = regime
                start_ts = ts
                end_ts = ts
                current_count = 1
        flush()

        if not windows_by_regime:
            return "Regime-specific failure windows: none (no red-flag windows)."
        top_windows = []
        for regime, windows in windows_by_regime.items():
            largest = max(windows, key=lambda item: item[2])
            top_windows.append((regime, largest[0], largest[1], largest[2]))
        ranked = sorted(top_windows, key=lambda item: item[3], reverse=True)[:2]
        desc = "; ".join(
            f"{regime}: {count} flagged trades ({start} → {end})"
            for regime, start, end, count in ranked
        )
        return f"Regime-specific failure windows: {desc}."

    def _build_slippage_sensitivity_card(self, explain_rows: list[dict[str, Any]]) -> str:
        slices: dict[str, list[float]] = {"low": [], "mid": [], "high": []}
        for row in explain_rows:
            fills = row.get("fill_context", [])
            if not isinstance(fills, list):
                continue
            for fill in fills:
                if not isinstance(fill, dict):
                    continue
                requested = abs(float(fill.get("requested_size", 0.0)))
                if requested <= 1e-9:
                    continue
                residual = abs(float(fill.get("residual_size", 0.0)))
                participation = float(fill.get("participation_rate", 0.0))
                residual_ratio = residual / requested
                if participation < 0.25:
                    slices["low"].append(residual_ratio)
                elif participation < 0.6:
                    slices["mid"].append(residual_ratio)
                else:
                    slices["high"].append(residual_ratio)
        populated = {k: v for k, v in slices.items() if v}
        if not populated:
            return "Slippage sensitivity slices: unavailable (no fill context)."
        parts = [
            f"{bucket} participation residual={sum(vals) / len(vals):.2%}"
            for bucket, vals in populated.items()
        ]
        return "Slippage sensitivity slices: " + ", ".join(parts) + "."

    def export_research_pack(self) -> None:
        task = self._selected_task()
        if task is None:
            task = next((item for item in reversed(self._task_queue) if item.state == "succeeded"), None)
        if task is None:
            messagebox.showinfo("Export Research Pack", "Run a workflow first to export a research pack.")
            return
        if task.state != "succeeded" or not task.logs:
            messagebox.showinfo("Export Research Pack", "Selected task has not completed successfully.")
            return

        output_text = task.workflow_output or (str(task.logs[-1]) if task.logs else "")
        pack_path = self._emit_research_pack(task, output_text)
        if pack_path is None:
            messagebox.showinfo("Export Research Pack", "Unable to export research pack for this task.")
            return
        task.research_pack_path = str(pack_path)
        messagebox.showinfo("Export Research Pack", f"Research pack exported to:\n{pack_path}")

    def _emit_research_pack(self, task: ResearchTask, output_text: str) -> Path | None:
        manifest_path = self._extract_manifest_path_from_output(output_text)
        run_dir = manifest_path.parent if manifest_path is not None else None
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_label = re.sub(r"[^a-z0-9]+", "_", task.label.lower()).strip("_") or "workflow"
        if run_dir is not None:
            pack_dir = run_dir / f"research_pack_{timestamp}"
        else:
            pack_dir = self._research_lab_dir / "research_packs" / f"{safe_label}_{timestamp}"
        pack_dir.mkdir(parents=True, exist_ok=True)

        source_manifest: dict[str, Any] = {}
        if manifest_path is not None and manifest_path.exists():
            try:
                source_manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            except Exception:
                source_manifest = {}

        metrics_tables = self._build_metrics_tables_payload(run_dir, source_manifest)
        stress_outputs = self._build_stress_outputs_payload(run_dir)
        parameter_set = self._build_parameter_set_payload(source_manifest)
        pipeline_graph = self._build_pipeline_graph_payload(source_manifest, manifest_path=manifest_path) if manifest_path is not None and source_manifest else {"nodes": [], "edges": []}

        packaged_manifest = {
            "task": {
                "id": task.task_id,
                "label": task.label,
                "state": task.state,
                "cancel_requested": bool(task.cancel_requested),
                "cancellation_reason": task.cancellation_reason,
                "cancellation_confirmed": bool(task.cancellation_confirmed),
                "cancellation": task.cancellation_token.snapshot() if task.cancellation_token else {},
            },
            "generated_at": datetime.now().isoformat(),
            "source_manifest_path": str(manifest_path) if manifest_path else None,
            "source_run_dir": str(run_dir) if run_dir else None,
            "included_files": [
                "manifest.json",
                "metrics_tables.json",
                "stress_outputs.json",
                "parameter_set.json",
                "pipeline_graph.json",
                "summary.md",
            ],
        }

        (pack_dir / "manifest.json").write_text(json.dumps(packaged_manifest, indent=2), encoding="utf-8")
        (pack_dir / "metrics_tables.json").write_text(json.dumps(metrics_tables, indent=2), encoding="utf-8")
        (pack_dir / "stress_outputs.json").write_text(json.dumps(stress_outputs, indent=2), encoding="utf-8")
        (pack_dir / "parameter_set.json").write_text(json.dumps(parameter_set, indent=2), encoding="utf-8")
        (pack_dir / "pipeline_graph.json").write_text(json.dumps(pipeline_graph, indent=2), encoding="utf-8")
        (pack_dir / "summary.md").write_text(
            self._build_research_pack_summary(task, output_text, metrics_tables, stress_outputs, parameter_set, pipeline_graph),
            encoding="utf-8",
        )

        if manifest_path is not None and manifest_path.exists():
            shutil.copy2(manifest_path, pack_dir / "source_manifest.json")
        return pack_dir

    def _build_metrics_tables_payload(self, run_dir: Path | None, manifest_payload: dict[str, Any]) -> dict[str, Any]:
        payload: dict[str, Any] = {"tables": {}, "manifest_metrics": manifest_payload.get("metrics", {})}
        if run_dir is None:
            return payload
        for file_name in ("metrics.json", "aggregate_metrics.json", "leaderboard.json", "fold_summary.json"):
            path = run_dir / file_name
            if not path.exists():
                continue
            try:
                payload["tables"][file_name] = json.loads(path.read_text(encoding="utf-8"))
            except Exception:
                payload["tables"][file_name] = {"error": f"failed to parse {file_name}"}
        return payload

    def _build_stress_outputs_payload(self, run_dir: Path | None) -> dict[str, Any]:
        payload: dict[str, Any] = {"stress_outputs": {}}
        if run_dir is None:
            return payload
        for file_name in ("stress_scenarios.json", "risk_diagnostics.json", "robustness_frontier.json"):
            path = run_dir / file_name
            if not path.exists():
                continue
            try:
                payload["stress_outputs"][file_name] = json.loads(path.read_text(encoding="utf-8"))
            except Exception:
                payload["stress_outputs"][file_name] = {"error": f"failed to parse {file_name}"}
        return payload

    def _build_parameter_set_payload(self, manifest_payload: dict[str, Any]) -> dict[str, Any]:
        if not manifest_payload:
            return {"parameters": {}}
        parameters = manifest_payload.get("parameters", {})
        return {
            "parameters": parameters if isinstance(parameters, dict) else {},
            "strategy": manifest_payload.get("strategy"),
            "timeframe": manifest_payload.get("timeframe"),
        }

    def _build_research_pack_summary(
        self,
        task: ResearchTask,
        output_text: str,
        metrics_tables: dict[str, Any],
        stress_outputs: dict[str, Any],
        parameter_set: dict[str, Any],
        pipeline_graph: dict[str, Any],
    ) -> str:
        metric_files = sorted(metrics_tables.get("tables", {}).keys()) if isinstance(metrics_tables.get("tables"), dict) else []
        stress_files = sorted(stress_outputs.get("stress_outputs", {}).keys()) if isinstance(stress_outputs.get("stress_outputs"), dict) else []
        parameter_keys = sorted((parameter_set.get("parameters") or {}).keys()) if isinstance(parameter_set.get("parameters"), dict) else []
        graph_nodes = pipeline_graph.get("nodes", []) if isinstance(pipeline_graph.get("nodes"), list) else []
        graph_edges = pipeline_graph.get("edges", []) if isinstance(pipeline_graph.get("edges"), list) else []
        lines = [
            f"# Research Pack Summary: {task.label}",
            "",
            "## Overview",
            f"- Task ID: `{task.task_id}`",
            f"- Status: `{task.state}`",
            f"- Cancellation requested: `{bool(task.cancel_requested)}`",
            f"- Generated: `{datetime.now().isoformat()}`",
            "",
            "## Included Components",
            "- manifest.json",
            "- metrics_tables.json",
            "- stress_outputs.json",
            "- parameter_set.json",
            "- pipeline_graph.json",
            "- summary.md",
            "",
            "## Metrics Tables",
            f"- Files captured: {', '.join(metric_files) if metric_files else 'None detected'}",
            "",
            "## Stress Outputs",
            f"- Files captured: {', '.join(stress_files) if stress_files else 'None detected'}",
            "",
            "## Parameter Set",
            f"- Parameter keys: {', '.join(parameter_keys[:20]) if parameter_keys else 'None detected'}",
            "",
            "## Pipeline Graph",
            f"- Nodes: {len(graph_nodes)}",
            f"- Edges: {len(graph_edges)}",
            "",
            "## Plain-English Result Summary",
            f"The workflow '{task.label}' completed and produced the output below.\n\n{output_text.strip()}",
            "",
        ]
        return "\n".join(lines)

    def _refresh_governance_dashboard_from_output(self, output_text: str) -> None:
        governance_payload = self._load_governance_payload_from_output(output_text)
        if not governance_payload:
            return

        gate_checks = governance_payload.get("gate_checks", {})
        gate_map = gate_checks if isinstance(gate_checks, dict) else {}
        total_gates = len(gate_map)
        pass_count = sum(1 for value in gate_map.values() if bool(value))
        fail_count = total_gates - pass_count

        missing_required_checks = governance_payload.get("missing_required_checks", [])
        missing_checks = (
            [str(item).strip() for item in missing_required_checks if str(item).strip()]
            if isinstance(missing_required_checks, list)
            else []
        )
        readiness = bool(governance_payload.get("is_promotion_ready", False))
        approval_status = str(governance_payload.get("approval_status", "pending")).strip() or "pending"
        promotion_state = str(governance_payload.get("promotion_state", "unknown")).strip() or "unknown"

        self._governance_gate_counts_var.set(f"{pass_count} passed / {fail_count} failed ({total_gates} total)")
        self._governance_missing_checks_var.set(
            ", ".join(missing_checks) if missing_checks else "None"
        )
        self._governance_promotion_ready_var.set("Ready" if readiness else "Not ready")
        self._governance_approval_status_var.set(f"{approval_status} ({promotion_state})")

    def _refresh_pipeline_graph_from_output(self, output_text: str) -> None:
        manifest_path = self._extract_lineage_manifest_path_from_output(output_text)
        if manifest_path is None or not manifest_path.exists():
            return
        try:
            manifest_payload = json.loads(manifest_path.read_text(encoding="utf-8"))
        except Exception:
            return
        if not isinstance(manifest_payload, dict):
            return
        graph_payload = self._build_pipeline_graph_payload(manifest_payload, manifest_path=manifest_path)
        self._set_pipeline_graph_payload(graph_payload)

    def _set_pipeline_graph_payload(self, graph_payload: dict[str, Any]) -> None:
        nodes = graph_payload.get("nodes", []) if isinstance(graph_payload.get("nodes"), list) else []
        rows = [f"{node.get('label', node.get('id', 'node'))} [{node.get('status', 'pending')}]" for node in nodes]
        self._pipeline_graph_payload = graph_payload
        self._pipeline_graph_nodes_by_id = {
            str(node.get("id")): node for node in nodes if isinstance(node, dict) and str(node.get("id", "")).strip()
        }
        if hasattr(self, "_pipeline_graph_items_var"):
            self._pipeline_graph_items_var.set(rows)
        if hasattr(self, "_pipeline_graph_listbox") and rows:
            self._pipeline_graph_listbox.selection_clear(0, tk.END)
            self._pipeline_graph_listbox.selection_set(0)
        self._refresh_pipeline_graph_node_details()

    def _selected_pipeline_graph_node(self) -> dict[str, Any] | None:
        if not hasattr(self, "_pipeline_graph_listbox"):
            return None
        selection = self._pipeline_graph_listbox.curselection()
        if not selection:
            return None
        index = int(selection[0])
        nodes = self._pipeline_graph_payload.get("nodes", []) if isinstance(self._pipeline_graph_payload, dict) else []
        if not isinstance(nodes, list) or index < 0 or index >= len(nodes):
            return None
        node = nodes[index]
        return node if isinstance(node, dict) else None

    def _refresh_pipeline_graph_node_details(self) -> None:
        if not hasattr(self, "_pipeline_graph_node_details_var"):
            return
        node = self._selected_pipeline_graph_node()
        if not node:
            self._pipeline_graph_node_details_var.set("No lineage graph loaded.")
            return
        details = [
            f"Node: {node.get('label', node.get('id', 'unknown'))}",
            f"Status: {node.get('status', 'pending')}",
            f"Artifact: {node.get('artifact_path') or 'n/a'}",
            f"Logs: {node.get('logs_path') or 'n/a'}",
        ]
        self._pipeline_graph_node_details_var.set("\n".join(details))

    def _open_pipeline_graph_reference(self, key: str) -> None:
        node = self._selected_pipeline_graph_node()
        if not node:
            return
        raw_path = str(node.get(key, "")).strip()
        if not raw_path:
            self._append_output(f"No {key} available for selected node.")
            return
        path = Path(raw_path)
        if not path.exists():
            self._append_output(f"{key} path does not exist: {path}")
            return
        self._append_output(f"Opening {key}: {path}")
        try:
            webbrowser.open(path.resolve().as_uri())
        except Exception:
            pass

    def _build_pipeline_graph_payload(self, manifest_payload: dict[str, Any], *, manifest_path: Path) -> dict[str, Any]:
        lineage = manifest_payload.get("lineage", {})
        run_references_raw = manifest_payload.get("run_references", [])
        gate_outcomes_raw = manifest_payload.get("gate_outcomes", [])
        promotion_history_raw = manifest_payload.get("promotion_history", [])
        decision_log_raw = manifest_payload.get("decision_log", [])
        run_references = [item for item in run_references_raw if isinstance(item, dict)] if isinstance(run_references_raw, list) else []
        gate_outcomes = [item for item in gate_outcomes_raw if isinstance(item, dict)] if isinstance(gate_outcomes_raw, list) else []
        promotion_history = [item for item in promotion_history_raw if isinstance(item, dict)] if isinstance(promotion_history_raw, list) else []
        decision_log = [item for item in decision_log_raw if isinstance(item, dict)] if isinstance(decision_log_raw, list) else []
        lineage_map = lineage if isinstance(lineage, dict) else {}

        nodes: list[dict[str, Any]] = []
        edges: list[dict[str, str]] = []
        run_reference_by_id = {
            str(item.get("run_id")): item
            for item in run_references
            if str(item.get("run_id", "")).strip()
        }

        def _append_node(node_id: str, *, label: str, node_type: str, status: str = "pending") -> None:
            if any(existing.get("id") == node_id for existing in nodes):
                return
            run_ref = run_reference_by_id.get(node_id, {})
            nodes.append(
                {
                    "id": node_id,
                    "label": label,
                    "type": node_type,
                    "status": str(run_ref.get("status", status)),
                    "artifact_path": run_ref.get("artifact_path"),
                    "logs_path": run_ref.get("logs_path"),
                    "step": run_ref.get("step"),
                }
            )

        hypothesis_id = str(lineage_map.get("hypothesis_id") or manifest_payload.get("hypothesis", {}).get("hypothesis_id") or "hypothesis")
        _append_node(hypothesis_id, label="Hypothesis", node_type="hypothesis", status="accepted")
        nodes[-1]["artifact_path"] = str(manifest_path)
        nodes[-1]["logs_path"] = str(manifest_path.parent / "logs.txt")

        stage_specs = [
            ("optimization_run_id", "Optimization", "optimization"),
            ("walk_forward_run_id", "Walk-forward", "walk_forward"),
            ("stress_run_id", "Stress", "stress"),
        ]
        for key, label, node_type in stage_specs:
            stage_id = str(lineage_map.get(key, "")).strip()
            if stage_id:
                _append_node(stage_id, label=label, node_type=node_type)

        gate_node_id = f"gate_checks_{hypothesis_id}"
        gate_status = "passed" if all(str(item.get("status", "")).lower() in {"accept", "pass", "passed"} for item in gate_outcomes) else "pending"
        _append_node(gate_node_id, label="Gate checks", node_type="gate_checks", status=gate_status)
        nodes[-1]["artifact_path"] = str(manifest_path)

        latest_promotion = promotion_history[-1] if promotion_history else {}
        latest_decision = decision_log[-1] if decision_log else {}
        promotion_status = str(latest_promotion.get("state") or latest_decision.get("decision") or "pending")
        promotion_node_id = f"promotion_decision_{hypothesis_id}"
        _append_node(promotion_node_id, label="Promotion decision", node_type="promotion_decision", status=promotion_status)
        nodes[-1]["artifact_path"] = str(manifest_path)

        links_payload = lineage_map.get("links", {}) if isinstance(lineage_map.get("links"), dict) else {}
        for link in links_payload.values():
            if not isinstance(link, dict):
                continue
            parent_id = str(link.get("parent_id", "")).strip()
            child_id = str(link.get("child_id", "")).strip()
            if parent_id and child_id:
                edges.append({"source": parent_id, "target": child_id})

        for run_ref in run_references:
            parent_id = str(run_ref.get("parent_id", "")).strip()
            run_id = str(run_ref.get("run_id", "")).strip()
            if parent_id and run_id:
                edges.append({"source": parent_id, "target": run_id})
        stress_id = str(lineage_map.get("stress_run_id", "")).strip()
        if stress_id:
            edges.append({"source": stress_id, "target": gate_node_id})
        edges.append({"source": gate_node_id, "target": promotion_node_id})

        deduped_edges: list[dict[str, str]] = []
        seen_pairs: set[tuple[str, str]] = set()
        for edge in edges:
            key = (edge["source"], edge["target"])
            if key in seen_pairs:
                continue
            seen_pairs.add(key)
            deduped_edges.append(edge)
        return {
            "schema_version": "1.0",
            "manifest_path": str(manifest_path),
            "nodes": nodes,
            "edges": deduped_edges,
        }

    def _refresh_funnel_kpi_dashboard(self) -> None:
        funnel = self._compute_funnel_metrics()
        self._funnel_acceptance_var.set(
            f"Acceptance rate: {funnel['acceptance_rate_pct']:.1f}% ({int(funnel['accepted_ideas'])}/{int(funnel['total_ideas'])})"
        )
        self._funnel_median_time_var.set(
            f"Median time-to-decision: {funnel['median_time_to_decision_days']:.1f} days"
        )
        self._funnel_false_positive_var.set(
            "False-positive proxy: "
            f"{funnel['false_positive_rate_pct']:.1f}% (accepted then later rejected)"
        )

        strategy_lines = []
        for row in funnel.get("pass_rates_by_strategy_family", []):
            strategy_lines.append(
                f"{row['strategy_family']}: {row['acceptance_rate_pct']:.1f}% "
                f"({row['accepted']}/{row['total']})"
            )
        self._funnel_strategy_rates_var.set("\n".join(strategy_lines) if strategy_lines else "No strategy-family funnel data yet.")

        month_lines = []
        for row in funnel.get("promotion_conversion_by_month", []):
            month_lines.append(
                f"{row['month']}: {row['promotion_conversion_pct']:.1f}% "
                f"({row['promoted']}/{row['total']})"
            )
        self._funnel_monthly_conversion_var.set(
            "\n".join(month_lines) if month_lines else "No monthly promotion conversion data yet."
        )

    def _load_governance_payload_from_output(self, output_text: str) -> dict[str, Any]:
        manifest_path = self._extract_manifest_path_from_output(output_text)
        if manifest_path is None or not manifest_path.exists():
            return {}
        try:
            manifest_payload = json.loads(manifest_path.read_text(encoding="utf-8"))
        except Exception:
            return {}
        if not isinstance(manifest_payload, dict):
            return {}
        governance_payload = manifest_payload.get("governance", {})
        if isinstance(governance_payload, dict) and governance_payload:
            return governance_payload
        parameters = manifest_payload.get("parameters", {})
        if isinstance(parameters, dict):
            governance_metadata = parameters.get("governance_metadata", {})
            if isinstance(governance_metadata, dict):
                return governance_metadata
        return {}

    def _extract_manifest_path_from_output(self, output_text: str) -> Path | None:
        candidates: list[Path] = []
        direct_path = Path(output_text.strip())
        if output_text.strip() and direct_path.exists():
            candidates.append(direct_path)

        for match in re.findall(r"Saved outputs to:\s*(.+)", output_text):
            raw = match.strip().splitlines()[0].strip()
            if raw:
                candidates.append(Path(raw))

        for candidate in reversed(candidates):
            if candidate.is_file() and candidate.name == "manifest.json":
                return candidate
            if candidate.is_dir():
                manifest_path = candidate / "manifest.json"
                if manifest_path.exists():
                    return manifest_path
        return None

    def _extract_lineage_manifest_path_from_output(self, output_text: str) -> Path | None:
        manifest_path = self._extract_manifest_path_from_output(output_text)
        if manifest_path is not None:
            return manifest_path
        for match in re.findall(r"Generated experiment skeleton:\s*(.+)", output_text):
            raw = match.strip().splitlines()[0].strip()
            if not raw:
                continue
            candidate = Path(raw)
            if candidate.is_file() and candidate.name == "research_project.json":
                return candidate
        direct_path = Path(output_text.strip())
        if direct_path.exists() and direct_path.is_file() and direct_path.suffix == ".json":
            return direct_path
        return None

    def _cancel_selected_task(self) -> None:
        task = self._selected_task()
        if task is None:
            return
        task.cancel_requested = True
        task.cancellation_reason = "Canceled from Research Lab UI"
        if task.cancellation_token is not None:
            task.cancellation_token.cancel(task.cancellation_reason)
        if task.state == "queued":
            task.state = "canceled"
            task.cancellation_confirmed = True
            task.logs.append("Task canceled before execution.")
        elif task.state == "running":
            task.state = "canceling"
            task.logs.append("Cancellation requested. Waiting for workflow cancellation checkpoint.")
        elif task.state == "retrying":
            task.state = "canceled"
            task.cancellation_confirmed = True
            task.logs.append("Retry canceled before task re-entered queue.")
        else:
            task.logs.append("Cancellation already requested.")
        self._refresh_task_queue_ui()

    def _retry_selected_task(self) -> None:
        task = self._selected_task()
        if task is None or task.state not in {"failed", "canceled"}:
            return
        task.state = "retrying"
        task.logs.append("Retry requested.")
        task.state = "queued"
        task.cancel_requested = False
        task.cancellation_reason = None
        task.cancellation_confirmed = False
        task.cancellation_token = CancellationToken()
        task.logs.append("Task re-queued for retry.")
        self._refresh_task_queue_ui()
        self._schedule_tasks()

    def _build_workflow_controls(self) -> None:
        controls = ttk.LabelFrame(self, text="Workflow Controls")
        controls.pack(fill="x", padx=40, pady=(6, 8))
        controls.columnconfigure(1, weight=1)
        controls.columnconfigure(2, weight=0)

        self.entry_signals_var = tk.StringVar(value="ts_momentum, breakout")
        self.exit_signals_var = tk.StringVar(value="none, momentum_flip")
        self.benchmark_selection_var = tk.StringVar(value="buy_hold, equal_weight_momentum, volatility_parity")
        self.optimization_trials_var = tk.StringVar(value="20")
        self.optimization_sampler_var = tk.StringVar(value="tpe")
        self.optimization_enable_pruning_var = tk.BooleanVar(value=True)
        self.optimization_prune_constraint_var = tk.BooleanVar(value=True)
        self.optimization_prune_lcb_var = tk.BooleanVar(value=True)
        self.optimization_min_completed_var = tk.StringVar(value="5")
        self.optimization_staged_budgets_var = tk.StringVar(value='[{"label":"coarse","n_trials":12,"sampler":"random","partial_period_fractions":[0.33,0.66]},{"label":"fine","n_trials":20,"sampler":"tpe","partial_period_fractions":[0.5,1.0]}]')
        self.wf_train_fraction_var = tk.StringVar(value="0.70")
        self.wf_validation_fraction_var = tk.StringVar(value="0.15")
        self.wf_test_fraction_var = tk.StringVar(value="0.15")
        self.wf_step_fraction_var = tk.StringVar(value="0.15")
        self.wf_split_policy_var = tk.StringVar(value="calendar-based")

        self.stress_enable_historical_replay_var = tk.BooleanVar(value=True)
        self.stress_historical_window_fraction_var = tk.StringVar(value="0.20")
        self.stress_historical_replay_window_bars_var = tk.StringVar(value="20")
        self.stress_synthetic_jump_magnitude_var = tk.StringVar(value="0.02")
        self.stress_synthetic_jump_interval_var = tk.StringVar(value="7")
        self.stress_synthetic_vol_cluster_multiplier_var = tk.StringVar(value="1.6")
        self.stress_overlay_spread_multiplier_var = tk.StringVar(value="2.5")
        self.stress_overlay_liquidity_multiplier_var = tk.StringVar(value="0.4")

        profile_options = tuple(self._rubric_templates["profiles"].keys())
        default_profile = str(self._rubric_templates.get("default_profile") or profile_options[0]) if profile_options else "intraday_alpha"
        self.hypothesis_rubric_profile_var = tk.StringVar(value=default_profile)
        default_scores = self._resolve_profile_defaults(default_profile)
        self.hypothesis_novelty_var = tk.StringVar(value=f"{default_scores['novelty']:.2f}")
        self.hypothesis_plausibility_var = tk.StringVar(value=f"{default_scores['plausibility']:.2f}")
        self.hypothesis_implementation_complexity_var = tk.StringVar(value=f"{default_scores['implementation_complexity']:.2f}")
        self.hypothesis_expected_capacity_var = tk.StringVar(value=f"{default_scores['expected_capacity']:.2f}")
        self.hypothesis_robustness_var = tk.StringVar(value=f"{default_scores['robustness']:.2f}")

        preset_options = tuple(self._workflow_presets.get("presets", {}).keys())
        default_preset = str(self._workflow_presets.get("default_preset", "custom"))
        if default_preset not in preset_options:
            default_preset = preset_options[0] if preset_options else "custom"
        self.workflow_preset_var = tk.StringVar(value=default_preset)
        self.easy_mode_var = tk.BooleanVar(value=True)
        self._workflow_validation_var = tk.StringVar(value="")
        self._preset_validation_warnings_var = tk.StringVar(value=self._format_preset_warning_text())
        self._advanced_workflow_widgets: list[tk.Widget] = []

        help_text = {
            "preset": "Preset bundles entry/exit signal sets plus optimization, walk-forward, and stress defaults.",
            "easy": "Easy Mode applies robust defaults and locks advanced tuning fields.",
            "entry": "Select one or more entry signal IDs: ts_momentum, ma_trend, breakout.",
            "exit": "Select one or more exit signal IDs: none, momentum_flip, trailing_stop, max_hold.",
            "opt": "n_trials controls search breadth. sampler controls candidate generation.",
            "prune": "Enable early stopping for weak trials; set minimum completed trials for pruning.",
            "staged": "Optional JSON list of optimization stages for coarse-to-fine search.",
            "wf_fracs": "Train/val/test must be >0 and sum to 1.0; step is walk-forward shift.",
            "wf_split": "Split policy controls fold construction for walk-forward validation.",
            "stress_toggle": "Turn historical replay stress on/off.",
            "stress_hist": "Replay window fraction and bars per replayed segment.",
            "stress_jump": "Synthetic jump magnitude, jump interval, and volatility clustering multiplier.",
            "stress_overlay": "Spread and liquidity multipliers for adverse execution overlays.",
        }

        def add_help(row: int, text: str) -> None:
            icon = ttk.Label(controls, text="ⓘ", foreground="#2a4f7a")
            icon.grid(row=row, column=2, sticky="w", padx=(0, 6))
            self._attach_tooltip(icon, text)

        row = 0
        ttk.Label(controls, text="Workflow preset").grid(row=row, column=0, sticky="w", padx=8, pady=5)
        preset_row = ttk.Frame(controls)
        preset_row.grid(row=row, column=1, sticky="ew", padx=8, pady=5)
        preset_row.columnconfigure(0, weight=1)
        ttk.Combobox(
            preset_row,
            textvariable=self.workflow_preset_var,
            values=preset_options,
            state="readonly",
            width=24,
        ).grid(row=0, column=0, sticky="w")
        ttk.Button(preset_row, text="Apply preset", command=self._apply_selected_workflow_preset).grid(row=0, column=1, sticky="w", padx=(8, 0))
        add_help(row, help_text["preset"])

        row += 1
        ttk.Label(controls, textvariable=self._preset_validation_warnings_var, foreground="#8a5a00", justify="left", wraplength=780).grid(
            row=row,
            column=0,
            columnspan=3,
            sticky="w",
            padx=8,
            pady=(0, 4),
        )

        row += 1
        ttk.Checkbutton(
            controls,
            text="Easy Mode (lock advanced fields + robust defaults)",
            variable=self.easy_mode_var,
            command=self._on_easy_mode_toggle,
        ).grid(row=row, column=0, columnspan=2, sticky="w", padx=8, pady=(2, 5))
        add_help(row, help_text["easy"])

        row += 1
        ttk.Label(controls, text="Entry signals").grid(row=row, column=0, sticky="nw", padx=8, pady=5)
        entry_row = ttk.Frame(controls)
        entry_row.grid(row=row, column=1, sticky="ew", padx=8, pady=5)
        entry_row.columnconfigure(0, weight=1)
        self.entry_signals_listbox = tk.Listbox(entry_row, selectmode="multiple", exportselection=False, height=len(self._signal_options))
        self.entry_signals_listbox.grid(row=0, column=0, sticky="ew")
        for option in self._signal_options:
            self.entry_signals_listbox.insert("end", option)
        self._set_listbox_selection(self.entry_signals_listbox, ["ts_momentum", "breakout"], valid_options=self._signal_options)
        self.entry_signals_listbox.bind("<<ListboxSelect>>", lambda _event: self._on_structured_selection_change("entry"))
        ttk.Button(entry_row, text="Select recommended set", command=lambda: self._select_recommended_set("entry")).grid(row=1, column=0, sticky="w", pady=(4, 0))
        add_help(row, help_text["entry"])

        row += 1
        ttk.Label(controls, text="Exit signals").grid(row=row, column=0, sticky="nw", padx=8, pady=5)
        exit_row = ttk.Frame(controls)
        exit_row.grid(row=row, column=1, sticky="ew", padx=8, pady=5)
        exit_row.columnconfigure(0, weight=1)
        self.exit_signals_listbox = tk.Listbox(exit_row, selectmode="multiple", exportselection=False, height=len(self._exit_signal_options))
        self.exit_signals_listbox.grid(row=0, column=0, sticky="ew")
        for option in self._exit_signal_options:
            self.exit_signals_listbox.insert("end", option)
        self._set_listbox_selection(self.exit_signals_listbox, ["none", "momentum_flip"], valid_options=self._exit_signal_options)
        self.exit_signals_listbox.bind("<<ListboxSelect>>", lambda _event: self._on_structured_selection_change("exit"))
        ttk.Button(exit_row, text="Select recommended set", command=lambda: self._select_recommended_set("exit")).grid(row=1, column=0, sticky="w", pady=(4, 0))
        add_help(row, help_text["exit"])

        row += 1
        ttk.Label(controls, text="Benchmarks").grid(row=row, column=0, sticky="nw", padx=8, pady=5)
        benchmark_row = ttk.Frame(controls)
        benchmark_row.grid(row=row, column=1, sticky="ew", padx=8, pady=5)
        benchmark_row.columnconfigure(0, weight=1)
        self.benchmark_selection_listbox = tk.Listbox(benchmark_row, selectmode="multiple", exportselection=False, height=len(self._benchmark_options))
        self.benchmark_selection_listbox.grid(row=0, column=0, sticky="ew")
        for option in self._benchmark_options:
            self.benchmark_selection_listbox.insert("end", option)
        self._set_listbox_selection(
            self.benchmark_selection_listbox,
            ["buy_hold", "equal_weight_momentum", "volatility_parity"],
            valid_options=self._benchmark_options,
        )
        self.benchmark_selection_listbox.bind("<<ListboxSelect>>", lambda _event: self._on_structured_selection_change("benchmark"))
        ttk.Button(
            benchmark_row,
            text="Select recommended set",
            command=lambda: self._select_recommended_set("benchmark"),
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        row += 1
        ttk.Label(controls, text="Optimization trials / sampler").grid(row=row, column=0, sticky="w", padx=8, pady=5)
        optimization_row = ttk.Frame(controls)
        optimization_row.grid(row=row, column=1, sticky="w", padx=8, pady=5)
        ttk.Entry(optimization_row, textvariable=self.optimization_trials_var, width=8).pack(side="left")
        ttk.Label(optimization_row, text=" / ").pack(side="left")
        ttk.Combobox(optimization_row, textvariable=self.optimization_sampler_var, values=self._sampler_options, state="readonly", width=12).pack(side="left")
        self._advanced_workflow_widgets.extend(optimization_row.winfo_children())
        add_help(row, help_text["opt"])

        row += 1
        ttk.Label(controls, text="Pruning enable/constraint/lcb/min").grid(row=row, column=0, sticky="w", padx=8, pady=5)
        prune_row = ttk.Frame(controls)
        prune_row.grid(row=row, column=1, sticky="w", padx=8, pady=5)
        ttk.Checkbutton(prune_row, text="Enable", variable=self.optimization_enable_pruning_var).pack(side="left")
        ttk.Checkbutton(prune_row, text="Constraint", variable=self.optimization_prune_constraint_var).pack(side="left", padx=(8, 0))
        ttk.Checkbutton(prune_row, text="LCB", variable=self.optimization_prune_lcb_var).pack(side="left", padx=(8, 0))
        ttk.Entry(prune_row, textvariable=self.optimization_min_completed_var, width=5).pack(side="left", padx=(8, 0))
        self._advanced_workflow_widgets.extend(prune_row.winfo_children())
        add_help(row, help_text["prune"])

        row += 1
        ttk.Label(controls, text="Staged budgets (JSON)").grid(row=row, column=0, sticky="w", padx=8, pady=5)
        staged_entry = ttk.Entry(controls, textvariable=self.optimization_staged_budgets_var)
        staged_entry.grid(row=row, column=1, sticky="ew", padx=8, pady=5)
        self._advanced_workflow_widgets.append(staged_entry)
        add_help(row, help_text["staged"])

        row += 1
        ttk.Label(controls, text="Walk-forward train/val/test/step").grid(row=row, column=0, sticky="w", padx=8, pady=5)
        wf_row = ttk.Frame(controls)
        wf_row.grid(row=row, column=1, sticky="w", padx=8, pady=5)
        ttk.Entry(wf_row, textvariable=self.wf_train_fraction_var, width=6).pack(side="left")
        ttk.Label(wf_row, text=" / ").pack(side="left")
        ttk.Entry(wf_row, textvariable=self.wf_validation_fraction_var, width=6).pack(side="left")
        ttk.Label(wf_row, text=" / ").pack(side="left")
        ttk.Entry(wf_row, textvariable=self.wf_test_fraction_var, width=6).pack(side="left")
        ttk.Label(wf_row, text=" / ").pack(side="left")
        ttk.Entry(wf_row, textvariable=self.wf_step_fraction_var, width=6).pack(side="left")
        add_help(row, help_text["wf_fracs"])

        row += 1
        ttk.Label(controls, text="Walk-forward split policy").grid(row=row, column=0, sticky="w", padx=8, pady=5)
        split_combo = ttk.Combobox(
            controls,
            textvariable=self.wf_split_policy_var,
            values=("calendar-based", "volatility-regime-stratified", "event-exclusion windows"),
            state="readonly",
            width=32,
        )
        split_combo.grid(row=row, column=1, sticky="w", padx=8, pady=5)
        self._advanced_workflow_widgets.append(split_combo)
        add_help(row, help_text["wf_split"])

        row += 1
        ttk.Label(controls, text="Stress: Replay/Jump/Overlay").grid(row=row, column=0, sticky="w", padx=8, pady=5)
        stress_row = ttk.Frame(controls)
        stress_row.grid(row=row, column=1, sticky="w", padx=8, pady=5)
        ttk.Checkbutton(stress_row, text="Historical replay regimes", variable=self.stress_enable_historical_replay_var).pack(side="left")
        add_help(row, help_text["stress_toggle"])

        row += 1
        ttk.Label(controls, text="Stress window frac / replay bars").grid(row=row, column=0, sticky="w", padx=8, pady=5)
        stress_row2 = ttk.Frame(controls)
        stress_row2.grid(row=row, column=1, sticky="w", padx=8, pady=5)
        ttk.Entry(stress_row2, textvariable=self.stress_historical_window_fraction_var, width=8).pack(side="left")
        ttk.Label(stress_row2, text=" / ").pack(side="left")
        ttk.Entry(stress_row2, textvariable=self.stress_historical_replay_window_bars_var, width=8).pack(side="left")
        self._advanced_workflow_widgets.extend(stress_row2.winfo_children())
        add_help(row, help_text["stress_hist"])

        row += 1
        ttk.Label(controls, text="Stress jump mag / interval / vol cluster").grid(row=row, column=0, sticky="w", padx=8, pady=5)
        stress_row3 = ttk.Frame(controls)
        stress_row3.grid(row=row, column=1, sticky="w", padx=8, pady=5)
        ttk.Entry(stress_row3, textvariable=self.stress_synthetic_jump_magnitude_var, width=8).pack(side="left")
        ttk.Label(stress_row3, text=" / ").pack(side="left")
        ttk.Entry(stress_row3, textvariable=self.stress_synthetic_jump_interval_var, width=8).pack(side="left")
        ttk.Label(stress_row3, text=" / ").pack(side="left")
        ttk.Entry(stress_row3, textvariable=self.stress_synthetic_vol_cluster_multiplier_var, width=8).pack(side="left")
        self._advanced_workflow_widgets.extend(stress_row3.winfo_children())
        add_help(row, help_text["stress_jump"])

        row += 1
        ttk.Label(controls, text="Stress overlay spread / liquidity").grid(row=row, column=0, sticky="w", padx=8, pady=5)
        stress_row4 = ttk.Frame(controls)
        stress_row4.grid(row=row, column=1, sticky="w", padx=8, pady=5)
        ttk.Entry(stress_row4, textvariable=self.stress_overlay_spread_multiplier_var, width=8).pack(side="left")
        ttk.Label(stress_row4, text=" / ").pack(side="left")
        ttk.Entry(stress_row4, textvariable=self.stress_overlay_liquidity_multiplier_var, width=8).pack(side="left")
        self._advanced_workflow_widgets.extend(stress_row4.winfo_children())
        add_help(row, help_text["stress_overlay"])

        row += 1
        ttk.Label(controls, text="Hypothesis rubric profile").grid(row=row, column=0, sticky="w", padx=8, pady=5)
        ttk.Combobox(controls, textvariable=self.hypothesis_rubric_profile_var, values=profile_options, state="readonly", width=24).grid(row=row, column=1, sticky="w", padx=8, pady=5)

        row += 1
        ttk.Label(controls, text="Hypothesis scores N/P/C/CAP/R").grid(row=row, column=0, sticky="w", padx=8, pady=5)
        hypothesis_row = ttk.Frame(controls)
        hypothesis_row.grid(row=row, column=1, sticky="w", padx=8, pady=5)
        ttk.Entry(hypothesis_row, textvariable=self.hypothesis_novelty_var, width=6).pack(side="left")
        ttk.Label(hypothesis_row, text=" / ").pack(side="left")
        ttk.Entry(hypothesis_row, textvariable=self.hypothesis_plausibility_var, width=6).pack(side="left")
        ttk.Label(hypothesis_row, text=" / ").pack(side="left")
        ttk.Entry(hypothesis_row, textvariable=self.hypothesis_implementation_complexity_var, width=6).pack(side="left")
        ttk.Label(hypothesis_row, text=" / ").pack(side="left")
        ttk.Entry(hypothesis_row, textvariable=self.hypothesis_expected_capacity_var, width=6).pack(side="left")
        ttk.Label(hypothesis_row, text=" / ").pack(side="left")
        ttk.Entry(hypothesis_row, textvariable=self.hypothesis_robustness_var, width=6).pack(side="left")

        row += 1
        ttk.Label(controls, textvariable=self._workflow_validation_var, foreground="#8a2d2d", justify="left", wraplength=780).grid(
            row=row,
            column=0,
            columnspan=3,
            sticky="w",
            padx=8,
            pady=(0, 6),
        )

        self.workflow_preset_var.trace_add("write", self._on_workflow_preset_change)
        self._attach_workflow_validation_traces()
        self._sync_signal_selection_vars()
        self._apply_selected_workflow_preset()
        self._on_easy_mode_toggle()

    def _build_wizard_mode(self) -> None:
        self._wizard_step_index = 0
        self._wizard_steps = [
            "idea capture",
            "data universe + period",
            "test plan",
            "run workflows",
            "review + promote/reject",
        ]

        self.wizard_idea_name_var = tk.StringVar(value="")
        self.wizard_idea_thesis_var = tk.StringVar(value="")
        self.wizard_idea_owner_var = tk.StringVar(value="research_lab_ui")
        self.wizard_data_universe_var = tk.StringVar(value=", ".join(self.controller.state.tickers))
        self.wizard_period_start_var = tk.StringVar(value=str(DEFAULT_BACKTEST_SETTINGS.get("start_date", "")))
        self.wizard_period_end_var = tk.StringVar(value=str(DEFAULT_BACKTEST_SETTINGS.get("end_date", "")))
        self.wizard_sector_include_var = tk.StringVar(value="")
        self.wizard_sector_exclude_var = tk.StringVar(value="")
        self.wizard_adv_threshold_var = tk.StringVar(value="")
        self.wizard_liquidity_threshold_var = tk.StringVar(value="")
        self.wizard_price_min_var = tk.StringVar(value="")
        self.wizard_price_max_var = tk.StringVar(value="")
        self.wizard_market_cap_min_var = tk.StringVar(value="")
        self.wizard_market_cap_max_var = tk.StringVar(value="")
        self.wizard_min_option_oi_var = tk.StringVar(value="")
        self.wizard_min_option_volume_var = tk.StringVar(value="")
        self.wizard_min_option_dte_var = tk.StringVar(value="")
        self.wizard_require_weeklies_var = tk.BooleanVar(value=False)
        self.wizard_test_plan_var = tk.StringVar(value="walk_forward")
        self.wizard_acceptance_var = tk.StringVar(value="Sharpe >= 0.8 and drawdown >= -0.25")
        self.wizard_run_validation_var = tk.BooleanVar(value=True)
        self.wizard_run_optimization_var = tk.BooleanVar(value=True)
        self.wizard_run_stress_var = tk.BooleanVar(value=True)
        self.wizard_review_notes_var = tk.StringVar(value="")
        self.wizard_promotion_decision_var = tk.StringVar(value="pending")
        self.wizard_session_label_var = tk.StringVar(value="")

        wizard = ttk.LabelFrame(self, text="Wizard Mode — Hypothesis Lifecycle")
        wizard.pack(fill="x", padx=40, pady=(8, 8))
        wizard.columnconfigure(0, weight=1)

        self._wizard_section_label = ttk.Label(wizard, text="", font=("Arial", 11, "bold"))
        self._wizard_section_label.grid(row=0, column=0, sticky="w", padx=10, pady=(8, 4))

        self._wizard_validation_label = ttk.Label(wizard, text="", foreground="#8a2d2d")
        self._wizard_validation_label.grid(row=1, column=0, sticky="w", padx=10, pady=(0, 6))

        self._wizard_content_frame = ttk.Frame(wizard)
        self._wizard_content_frame.grid(row=2, column=0, sticky="ew", padx=10)
        self._wizard_content_frame.columnconfigure(0, weight=1)

        self._wizard_step_frames: list[ttk.Frame] = []
        self._wizard_build_step_idea_capture()
        self._wizard_build_step_universe_period()
        self._wizard_build_step_test_plan()
        self._wizard_build_step_run_workflows()
        self._wizard_build_step_review()

        nav = ttk.Frame(wizard)
        nav.grid(row=3, column=0, sticky="ew", padx=10, pady=(8, 10))
        session_tools = ttk.Frame(nav)
        session_tools.pack(side="left")
        ttk.Label(session_tools, text="Session label").pack(side="left")
        ttk.Entry(session_tools, textvariable=self.wizard_session_label_var, width=20).pack(side="left", padx=(6, 6))
        ttk.Button(session_tools, text="Export Session", command=self._wizard_export_session).pack(side="left")
        ttk.Button(session_tools, text="Import Session", command=self._wizard_import_session).pack(side="left", padx=(6, 0))
        self._wizard_back_button = ttk.Button(nav, text="Back", command=self._wizard_go_back)
        self._wizard_back_button.pack(side="left", padx=(12, 0))
        self._wizard_next_button = ttk.Button(nav, text="Next", command=self._wizard_go_next)
        self._wizard_next_button.pack(side="right")

        self._wizard_attach_state_traces()
        self._wizard_load_state()
        self._wizard_show_step(self._wizard_step_index)

    def _wizard_add_step_frame(self) -> ttk.Frame:
        frame = ttk.Frame(self._wizard_content_frame)
        frame.grid(row=0, column=0, sticky="ew")
        frame.columnconfigure(0, weight=1)
        self._wizard_step_frames.append(frame)
        return frame

    def _wizard_build_step_idea_capture(self) -> None:
        frame = self._wizard_add_step_frame()
        ttk.Label(frame, text="[STEP 1] IDEA CAPTURE", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 6))
        ttk.Label(frame, text="Idea name").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.wizard_idea_name_var).grid(row=2, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(frame, text="Core thesis").grid(row=3, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.wizard_idea_thesis_var).grid(row=4, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(frame, text="Owner").grid(row=5, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.wizard_idea_owner_var).grid(row=6, column=0, sticky="ew")

    def _wizard_build_step_universe_period(self) -> None:
        frame = self._wizard_add_step_frame()
        ttk.Label(frame, text="[STEP 2] DATA UNIVERSE + PERIOD", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 6))
        ttk.Label(frame, text="Universe tickers (comma-separated)").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.wizard_data_universe_var).grid(row=2, column=0, sticky="ew", pady=(0, 6))
        period_row = ttk.Frame(frame)
        period_row.grid(row=3, column=0, sticky="w")
        ttk.Label(period_row, text="Period start").pack(side="left")
        ttk.Entry(period_row, textvariable=self.wizard_period_start_var, width=14).pack(side="left", padx=(8, 12))
        ttk.Label(period_row, text="Period end").pack(side="left")
        ttk.Entry(period_row, textvariable=self.wizard_period_end_var, width=14).pack(side="left", padx=(8, 0))

        panel = ttk.LabelFrame(frame, text="Universe Builder")
        panel.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        panel.columnconfigure(1, weight=1)

        row = 0
        ttk.Label(panel, text="Sector include / exclude").grid(row=row, column=0, sticky="w", padx=8, pady=4)
        sector_row = ttk.Frame(panel)
        sector_row.grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        sector_row.columnconfigure(0, weight=1)
        sector_row.columnconfigure(1, weight=1)
        ttk.Entry(sector_row, textvariable=self.wizard_sector_include_var).grid(row=0, column=0, sticky="ew")
        ttk.Entry(sector_row, textvariable=self.wizard_sector_exclude_var).grid(row=0, column=1, sticky="ew", padx=(6, 0))

        row += 1
        ttk.Label(panel, text="Min ADV / Min liquidity").grid(row=row, column=0, sticky="w", padx=8, pady=4)
        adv_row = ttk.Frame(panel)
        adv_row.grid(row=row, column=1, sticky="w", padx=8, pady=4)
        ttk.Entry(adv_row, textvariable=self.wizard_adv_threshold_var, width=14).pack(side="left")
        ttk.Entry(adv_row, textvariable=self.wizard_liquidity_threshold_var, width=14).pack(side="left", padx=(6, 0))

        row += 1
        ttk.Label(panel, text="Price min / max").grid(row=row, column=0, sticky="w", padx=8, pady=4)
        price_row = ttk.Frame(panel)
        price_row.grid(row=row, column=1, sticky="w", padx=8, pady=4)
        ttk.Entry(price_row, textvariable=self.wizard_price_min_var, width=14).pack(side="left")
        ttk.Entry(price_row, textvariable=self.wizard_price_max_var, width=14).pack(side="left", padx=(6, 0))

        row += 1
        ttk.Label(panel, text="Market-cap min / max").grid(row=row, column=0, sticky="w", padx=8, pady=4)
        mcap_row = ttk.Frame(panel)
        mcap_row.grid(row=row, column=1, sticky="w", padx=8, pady=4)
        ttk.Entry(mcap_row, textvariable=self.wizard_market_cap_min_var, width=14).pack(side="left")
        ttk.Entry(mcap_row, textvariable=self.wizard_market_cap_max_var, width=14).pack(side="left", padx=(6, 0))

        row += 1
        ttk.Label(panel, text="Min option OI / volume").grid(row=row, column=0, sticky="w", padx=8, pady=4)
        options_row = ttk.Frame(panel)
        options_row.grid(row=row, column=1, sticky="w", padx=8, pady=4)
        ttk.Entry(options_row, textvariable=self.wizard_min_option_oi_var, width=14).pack(side="left")
        ttk.Entry(options_row, textvariable=self.wizard_min_option_volume_var, width=14).pack(side="left", padx=(6, 0))

        row += 1
        ttk.Label(panel, text="Min option DTE").grid(row=row, column=0, sticky="w", padx=8, pady=4)
        ttk.Entry(panel, textvariable=self.wizard_min_option_dte_var, width=14).grid(row=row, column=1, sticky="w", padx=8, pady=4)

        row += 1
        ttk.Checkbutton(panel, text="Require weekly-listed options", variable=self.wizard_require_weeklies_var).grid(
            row=row, column=0, columnspan=2, sticky="w", padx=8, pady=(4, 8)
        )

    def _wizard_build_step_test_plan(self) -> None:
        frame = self._wizard_add_step_frame()
        ttk.Label(frame, text="[STEP 3] TEST PLAN", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 6))
        ttk.Label(frame, text="Primary test").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.wizard_test_plan_var).grid(row=2, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(frame, text="Acceptance criteria").grid(row=3, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.wizard_acceptance_var).grid(row=4, column=0, sticky="ew")

    def _wizard_build_step_run_workflows(self) -> None:
        frame = self._wizard_add_step_frame()
        ttk.Label(frame, text="[STEP 4] RUN WORKFLOWS", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 6))
        check_row = ttk.Frame(frame)
        check_row.grid(row=1, column=0, sticky="w")
        ttk.Checkbutton(check_row, text="Walk-forward validation", variable=self.wizard_run_validation_var).pack(side="left", padx=(0, 10))
        ttk.Checkbutton(check_row, text="Optimization", variable=self.wizard_run_optimization_var).pack(side="left", padx=(0, 10))
        ttk.Checkbutton(check_row, text="Stress tests", variable=self.wizard_run_stress_var).pack(side="left")
        self._wizard_run_button = ttk.Button(frame, text="Run Selected Workflows", command=self._wizard_run_selected_workflows)
        self._wizard_run_button.grid(row=2, column=0, sticky="w", pady=(8, 0))

    def _wizard_build_step_review(self) -> None:
        frame = self._wizard_add_step_frame()
        ttk.Label(frame, text="[STEP 5] REVIEW + PROMOTE/REJECT", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 6))
        ttk.Label(frame, text="Review notes").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.wizard_review_notes_var).grid(row=2, column=0, sticky="ew", pady=(0, 6))
        owner_row = ttk.Frame(frame)
        owner_row.grid(row=3, column=0, sticky="ew", pady=(0, 6))
        owner_row.columnconfigure(1, weight=1)
        ttk.Label(owner_row, text="Note owner").grid(row=0, column=0, sticky="w")
        self.wizard_review_owner_var = tk.StringVar(value="")
        ttk.Entry(owner_row, textvariable=self.wizard_review_owner_var).grid(row=0, column=1, sticky="ew", padx=(8, 0))
        ttk.Button(owner_row, text="Append note", command=self._wizard_append_review_note).grid(row=0, column=2, padx=(8, 0))
        self.wizard_review_log_text = tk.Text(frame, height=5)
        self.wizard_review_log_text.grid(row=4, column=0, sticky="ew", pady=(0, 6))
        self.wizard_review_log_text.configure(state="disabled")
        decision_row = ttk.Frame(frame)
        decision_row.grid(row=5, column=0, sticky="w")
        ttk.Radiobutton(decision_row, text="Promote", value="promote", variable=self.wizard_promotion_decision_var).pack(side="left")
        ttk.Radiobutton(decision_row, text="Reject", value="reject", variable=self.wizard_promotion_decision_var).pack(side="left", padx=(8, 0))

    def _wizard_append_review_note(self) -> None:
        note = self.wizard_review_notes_var.get().strip()
        if not note:
            return
        owner = self.wizard_review_owner_var.get().strip() or self.wizard_idea_owner_var.get().strip() or "research_lab_ui"
        entry = {
            "owner": owner,
            "note": note,
            "timestamp": datetime.now().isoformat(timespec="seconds"),
        }
        self._wizard_comments.append(entry)
        self._wizard_append_history_event("review_note", {"owner": owner, "note": note})
        self.wizard_review_notes_var.set("")
        self._wizard_render_comments_log()
        self._wizard_persist_state()

    def _wizard_render_comments_log(self) -> None:
        if not hasattr(self, "wizard_review_log_text"):
            return
        self.wizard_review_log_text.configure(state="normal")
        self.wizard_review_log_text.delete("1.0", tk.END)
        for row in self._wizard_comments:
            self.wizard_review_log_text.insert(
                tk.END,
                f"[{row.get('timestamp', '')}] {row.get('owner', 'research_lab_ui')}: {row.get('note', '')}\n",
            )
        self.wizard_review_log_text.configure(state="disabled")

    def _wizard_append_history_event(self, event: str, details: dict[str, Any] | None = None) -> None:
        row: dict[str, str] = {
            "event": event,
            "timestamp": datetime.now().isoformat(timespec="seconds"),
        }
        if details:
            for key, value in details.items():
                row[str(key)] = str(value)
        self._wizard_history.append(row)

    def _wizard_attach_state_traces(self) -> None:
        tracked_vars = (
            self.wizard_idea_name_var,
            self.wizard_idea_thesis_var,
            self.wizard_idea_owner_var,
            self.wizard_data_universe_var,
            self.wizard_period_start_var,
            self.wizard_period_end_var,
            self.wizard_sector_include_var,
            self.wizard_sector_exclude_var,
            self.wizard_adv_threshold_var,
            self.wizard_liquidity_threshold_var,
            self.wizard_price_min_var,
            self.wizard_price_max_var,
            self.wizard_market_cap_min_var,
            self.wizard_market_cap_max_var,
            self.wizard_min_option_oi_var,
            self.wizard_min_option_volume_var,
            self.wizard_min_option_dte_var,
            self.wizard_require_weeklies_var,
            self.wizard_test_plan_var,
            self.wizard_acceptance_var,
            self.wizard_run_validation_var,
            self.wizard_run_optimization_var,
            self.wizard_run_stress_var,
            self.wizard_review_notes_var,
            self.wizard_promotion_decision_var,
        )
        for var in tracked_vars:
            var.trace_add("write", self._wizard_on_state_change)

    def _wizard_on_state_change(self, *_: object) -> None:
        self._wizard_persist_state()
        self._wizard_refresh_nav_state()

    def _wizard_go_back(self) -> None:
        if self._wizard_step_index <= 0:
            return
        self._wizard_show_step(self._wizard_step_index - 1)

    def _wizard_go_next(self) -> None:
        is_valid, _ = self._wizard_validate_step(self._wizard_step_index)
        if not is_valid:
            return
        if self._wizard_step_index < len(self._wizard_steps) - 1:
            self._wizard_show_step(self._wizard_step_index + 1)
            return
        self._wizard_finalize_review()

    def _wizard_show_step(self, step_index: int) -> None:
        self._wizard_step_index = max(0, min(step_index, len(self._wizard_steps) - 1))
        for index, frame in enumerate(self._wizard_step_frames):
            if index == self._wizard_step_index:
                frame.grid()
            else:
                frame.grid_remove()
        self._wizard_persist_state()
        self._wizard_refresh_nav_state()

    def _wizard_validate_step(self, step_index: int) -> tuple[bool, str]:
        if step_index == 0:
            if not self.wizard_idea_name_var.get().strip() or not self.wizard_idea_thesis_var.get().strip():
                return False, "Idea capture requires both idea name and core thesis."
            return True, ""
        if step_index == 1:
            if not [item.strip() for item in self.wizard_data_universe_var.get().split(",") if item.strip()]:
                return False, "Data universe must include at least one ticker."
            start = parse_date(self.wizard_period_start_var.get().strip())
            end = parse_date(self.wizard_period_end_var.get().strip())
            if start is None or end is None:
                return False, "Period requires valid start and end dates."
            if start >= end:
                return False, "Period start must be before period end."
            return True, ""
        if step_index == 2:
            if not self.wizard_test_plan_var.get().strip() or not self.wizard_acceptance_var.get().strip():
                return False, "Test plan requires a primary test and acceptance criteria."
            return True, ""
        if step_index == 3:
            if not any(
                (
                    bool(self.wizard_run_validation_var.get()),
                    bool(self.wizard_run_optimization_var.get()),
                    bool(self.wizard_run_stress_var.get()),
                )
            ):
                return False, "Select at least one workflow before running."
            return True, ""
        if step_index == 4:
            if self.wizard_promotion_decision_var.get().strip().lower() not in {"promote", "reject"}:
                return False, "Review step requires either Promote or Reject."
            if not self.wizard_review_notes_var.get().strip():
                return False, "Review step requires notes to justify the decision."
            return True, ""
        return True, ""

    def _has_running_task(self) -> bool:
        return any(task.state == "running" for task in self._task_queue)

    def _wizard_refresh_nav_state(self) -> None:
        valid, error = self._wizard_validate_step(self._wizard_step_index)
        step_name = self._wizard_steps[self._wizard_step_index].upper()
        self._wizard_section_label.configure(text=f"Section {self._wizard_step_index + 1}/{len(self._wizard_steps)} — {step_name}")
        self._wizard_validation_label.configure(text=error)
        self._wizard_back_button.configure(state="normal" if self._wizard_step_index > 0 else "disabled")
        next_label = "Finish" if self._wizard_step_index == len(self._wizard_steps) - 1 else "Next"
        self._wizard_next_button.configure(text=next_label, state="normal" if valid else "disabled")
        if hasattr(self, "_wizard_run_button"):
            run_valid, _ = self._wizard_validate_step(3)
            run_enabled = run_valid and not self._has_running_task()
            self._wizard_run_button.configure(state="normal" if run_enabled else "disabled")

    def _wizard_run_selected_workflows(self) -> None:
        valid, error = self._wizard_validate_step(3)
        if not valid:
            messagebox.showinfo("Invalid run configuration", error)
            return
        context = self._build_common_context()
        if context is None:
            return
        context["universe_filters"] = self._build_universe_filters()
        wizard_tickers = [item.strip() for item in self.wizard_data_universe_var.get().split(",") if item.strip()]
        parsed_start = parse_date(self.wizard_period_start_var.get().strip())
        parsed_end = parse_date(self.wizard_period_end_var.get().strip())
        if wizard_tickers:
            context["tickers"] = wizard_tickers
        if parsed_start is not None and parsed_end is not None and parsed_start < parsed_end:
            context["start_date"] = parsed_start
            context["end_date"] = parsed_end

        config = self._build_workflow_config()
        if config is None:
            return

        selected: list[tuple[str, Callable[[dict[str, Any], ResearchWorkflowConfig], str]]] = []
        if self.wizard_run_validation_var.get():
            selected.append(("Walk-forward validation", self._run_walk_forward_workflow))
        if self.wizard_run_optimization_var.get():
            selected.append(("Parameter optimization", self._run_optimization_workflow))
        if self.wizard_run_stress_var.get():
            selected.append(("Stress/scenario tests", self._run_stress_workflow))

        self._append_output("Queueing wizard workflows...")
        for label, runner in selected:
            self._enqueue_task(label=label, target=runner, context=context, config=config)

    def _wizard_finalize_review(self) -> None:
        decision = self.wizard_promotion_decision_var.get().strip().lower()
        note = self.wizard_review_notes_var.get().strip()
        self._append_output(
            "Wizard review complete | "
            f"idea={self.wizard_idea_name_var.get().strip() or 'n/a'} | "
            f"decision={decision} | notes={note}"
        )
        if note:
            self._wizard_append_review_note()
        self._wizard_persist_state()

    def _wizard_state_payload(self) -> dict[str, Any]:
        return {
            "current_step": self._wizard_step_index,
            "idea_name": self.wizard_idea_name_var.get().strip(),
            "idea_thesis": self.wizard_idea_thesis_var.get().strip(),
            "idea_owner": self.wizard_idea_owner_var.get().strip(),
            "data_universe": self.wizard_data_universe_var.get().strip(),
            "period_start": self.wizard_period_start_var.get().strip(),
            "period_end": self.wizard_period_end_var.get().strip(),
            "sector_include": self.wizard_sector_include_var.get().strip(),
            "sector_exclude": self.wizard_sector_exclude_var.get().strip(),
            "adv_threshold": self.wizard_adv_threshold_var.get().strip(),
            "liquidity_threshold": self.wizard_liquidity_threshold_var.get().strip(),
            "price_min": self.wizard_price_min_var.get().strip(),
            "price_max": self.wizard_price_max_var.get().strip(),
            "market_cap_min": self.wizard_market_cap_min_var.get().strip(),
            "market_cap_max": self.wizard_market_cap_max_var.get().strip(),
            "min_option_oi": self.wizard_min_option_oi_var.get().strip(),
            "min_option_volume": self.wizard_min_option_volume_var.get().strip(),
            "min_option_dte": self.wizard_min_option_dte_var.get().strip(),
            "require_weeklies": bool(self.wizard_require_weeklies_var.get()),
            "test_plan": self.wizard_test_plan_var.get().strip(),
            "acceptance_criteria": self.wizard_acceptance_var.get().strip(),
            "run_validation": bool(self.wizard_run_validation_var.get()),
            "run_optimization": bool(self.wizard_run_optimization_var.get()),
            "run_stress": bool(self.wizard_run_stress_var.get()),
            "review_notes": self.wizard_review_notes_var.get().strip(),
            "promotion_decision": self.wizard_promotion_decision_var.get().strip(),
            "wizard_comments": list(self._wizard_comments),
            "wizard_history": list(self._wizard_history),
        }

    def _wizard_state_document(self, payload: dict[str, Any]) -> dict[str, Any]:
        doc: dict[str, Any] = {
            "schema_version": WIZARD_STATE_SCHEMA_VERSION,
            "payload": payload,
        }
        digest = hashlib.sha256(json.dumps(doc, sort_keys=True, separators=(",", ":")).encode("utf-8")).hexdigest()
        doc["checksum"] = f"sha256:{digest}"
        return doc

    def _wizard_extract_payload(self, raw: Any) -> tuple[dict[str, Any] | None, bool]:
        if not isinstance(raw, dict):
            return None, False
        if "payload" not in raw:
            return dict(raw), True
        payload = raw.get("payload")
        if not isinstance(payload, dict):
            return None, False
        schema_version = int(raw.get("schema_version", 0) or 0)
        expected = raw.get("checksum")
        check_doc = {
            "schema_version": schema_version,
            "payload": payload,
        }
        digest = hashlib.sha256(json.dumps(check_doc, sort_keys=True, separators=(",", ":")).encode("utf-8")).hexdigest()
        if isinstance(expected, str) and expected.startswith("sha256:") and expected != f"sha256:{digest}":
            return None, False
        return dict(payload), schema_version < WIZARD_STATE_SCHEMA_VERSION

    def _merge_wizard_activity_rows(self, existing: list[dict[str, Any]], incoming: list[dict[str, Any]]) -> list[dict[str, Any]]:
        merged = list(existing) + list(incoming)
        deduped: list[dict[str, Any]] = []
        seen: set[str] = set()
        for row in merged:
            if not isinstance(row, dict):
                continue
            fingerprint = json.dumps(row, sort_keys=True)
            if fingerprint in seen:
                continue
            seen.add(fingerprint)
            deduped.append(dict(row))
        deduped.sort(key=lambda row: str(row.get("timestamp") or row.get("recorded_at") or row.get("commented_at") or ""))
        return deduped

    def _wizard_apply_payload(self, payload: dict[str, Any], *, merge_activity: bool) -> None:
        current_comments = list(self._wizard_comments)
        current_history = list(self._wizard_history)

        self.wizard_idea_name_var.set(str(payload.get("idea_name", self.wizard_idea_name_var.get())))
        self.wizard_idea_thesis_var.set(str(payload.get("idea_thesis", self.wizard_idea_thesis_var.get())))
        self.wizard_idea_owner_var.set(str(payload.get("idea_owner", self.wizard_idea_owner_var.get())))
        self.wizard_data_universe_var.set(str(payload.get("data_universe", self.wizard_data_universe_var.get())))
        self.wizard_period_start_var.set(str(payload.get("period_start", self.wizard_period_start_var.get())))
        self.wizard_period_end_var.set(str(payload.get("period_end", self.wizard_period_end_var.get())))
        self.wizard_sector_include_var.set(str(payload.get("sector_include", self.wizard_sector_include_var.get())))
        self.wizard_sector_exclude_var.set(str(payload.get("sector_exclude", self.wizard_sector_exclude_var.get())))
        self.wizard_adv_threshold_var.set(str(payload.get("adv_threshold", self.wizard_adv_threshold_var.get())))
        self.wizard_liquidity_threshold_var.set(str(payload.get("liquidity_threshold", self.wizard_liquidity_threshold_var.get())))
        self.wizard_price_min_var.set(str(payload.get("price_min", self.wizard_price_min_var.get())))
        self.wizard_price_max_var.set(str(payload.get("price_max", self.wizard_price_max_var.get())))
        self.wizard_market_cap_min_var.set(str(payload.get("market_cap_min", self.wizard_market_cap_min_var.get())))
        self.wizard_market_cap_max_var.set(str(payload.get("market_cap_max", self.wizard_market_cap_max_var.get())))
        self.wizard_min_option_oi_var.set(str(payload.get("min_option_oi", self.wizard_min_option_oi_var.get())))
        self.wizard_min_option_volume_var.set(str(payload.get("min_option_volume", self.wizard_min_option_volume_var.get())))
        self.wizard_min_option_dte_var.set(str(payload.get("min_option_dte", self.wizard_min_option_dte_var.get())))
        self.wizard_require_weeklies_var.set(bool(payload.get("require_weeklies", self.wizard_require_weeklies_var.get())))
        self.wizard_test_plan_var.set(str(payload.get("test_plan", self.wizard_test_plan_var.get())))
        self.wizard_acceptance_var.set(str(payload.get("acceptance_criteria", self.wizard_acceptance_var.get())))
        self.wizard_run_validation_var.set(bool(payload.get("run_validation", self.wizard_run_validation_var.get())))
        self.wizard_run_optimization_var.set(bool(payload.get("run_optimization", self.wizard_run_optimization_var.get())))
        self.wizard_run_stress_var.set(bool(payload.get("run_stress", self.wizard_run_stress_var.get())))
        self.wizard_review_notes_var.set(str(payload.get("review_notes", self.wizard_review_notes_var.get())))
        self.wizard_promotion_decision_var.set(str(payload.get("promotion_decision", self.wizard_promotion_decision_var.get())))

        saved_comments = payload.get("wizard_comments", [])
        if isinstance(saved_comments, list):
            incoming_comments = [dict(row) for row in saved_comments if isinstance(row, dict)]
            self._wizard_comments = self._merge_wizard_activity_rows(current_comments, incoming_comments) if merge_activity else incoming_comments

        saved_history = payload.get("wizard_history", [])
        if isinstance(saved_history, list):
            incoming_history = [dict(row) for row in saved_history if isinstance(row, dict)]
            self._wizard_history = self._merge_wizard_activity_rows(current_history, incoming_history) if merge_activity else incoming_history

        self._wizard_render_comments_log()
        loaded_step = int(payload.get("current_step", 0))
        self._wizard_step_index = max(0, min(loaded_step, len(self._wizard_steps) - 1))

    def _wizard_persist_state(self) -> None:
        self._research_lab_dir.mkdir(parents=True, exist_ok=True)
        payload = self._wizard_state_payload()
        self._wizard_state_path.write_text(json.dumps(self._wizard_state_document(payload), indent=2, sort_keys=True), encoding="utf-8")

    def _wizard_load_state(self) -> None:
        if not self._wizard_state_path.exists():
            return
        try:
            raw = json.loads(self._wizard_state_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return
        payload, upgraded = self._wizard_extract_payload(raw)
        if payload is None:
            self._append_output("Skipped wizard state load due to checksum mismatch or invalid schema.")
            return
        self._wizard_apply_payload(payload, merge_activity=False)
        if upgraded:
            self._wizard_persist_state()

    def _wizard_session_storage_dir(self) -> Path:
        return self._research_lab_dir / "sessions"

    def _wizard_export_session(self) -> None:
        payload = self._wizard_state_payload()
        label_raw = self.wizard_session_label_var.get().strip() or payload.get("idea_name", "") or "session"
        safe_label = re.sub(r"[^a-zA-Z0-9_-]+", "_", str(label_raw)).strip("_") or "session"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_dir = self._wizard_session_storage_dir()
        out_dir.mkdir(parents=True, exist_ok=True)
        output_path = out_dir / f"{timestamp}_{safe_label}.json"
        output_path.write_text(json.dumps(self._wizard_state_document(payload), indent=2, sort_keys=True), encoding="utf-8")
        self._wizard_append_history_event("session_export", {"label": label_raw, "path": str(output_path)})
        self._wizard_persist_state()
        self._append_output(f"Wizard session exported to {output_path}.")
        messagebox.showinfo("Session exported", f"Saved wizard session to:\n{output_path}")

    def _wizard_import_session(self) -> None:
        source = filedialog.askopenfilename(
            title="Import wizard session",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialdir=str(self._wizard_session_storage_dir()),
        )
        if not source:
            return
        try:
            raw = json.loads(Path(source).read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            messagebox.showinfo("Invalid session", "Could not read wizard session JSON.")
            return
        payload, _ = self._wizard_extract_payload(raw)
        if payload is None:
            messagebox.showinfo("Invalid session", "Session checksum validation failed.")
            return
        self._wizard_apply_payload(payload, merge_activity=True)
        self._wizard_append_history_event("session_import", {"path": source})
        self._wizard_persist_state()
        self._wizard_refresh_nav_state()
        self._append_output(f"Wizard session imported from {source}.")

    def _start_worker(self, target: Callable[[dict[str, Any], ResearchWorkflowConfig], str], label: str) -> None:
        context = self._build_common_context()
        if context is None:
            return
        context["universe_filters"] = self._build_universe_filters()
        config = self._build_workflow_config()
        if config is None:
            return

        self._enqueue_task(label=label, target=target, context=context, config=config)

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

    def _apply_easy_mode_defaults(self) -> None:
        self.optimization_sampler_var.set("tpe")
        self.optimization_enable_pruning_var.set(True)
        self.optimization_prune_constraint_var.set(True)
        self.optimization_prune_lcb_var.set(True)
        self.optimization_min_completed_var.set("5")
        self.optimization_staged_budgets_var.set(
            '[{"label":"coarse","n_trials":12,"sampler":"random","partial_period_fractions":[0.33,0.66]},'
            '{"label":"fine","n_trials":20,"sampler":"tpe","partial_period_fractions":[0.5,1.0]}]'
        )
        self.wf_split_policy_var.set("calendar-based")
        self.stress_enable_historical_replay_var.set(True)
        self.stress_historical_window_fraction_var.set("0.20")
        self.stress_historical_replay_window_bars_var.set("20")
        self.stress_synthetic_jump_magnitude_var.set("0.02")
        self.stress_synthetic_jump_interval_var.set("7")
        self.stress_synthetic_vol_cluster_multiplier_var.set("1.6")
        self.stress_overlay_spread_multiplier_var.set("2.5")
        self.stress_overlay_liquidity_multiplier_var.set("0.4")

    def _on_easy_mode_toggle(self) -> None:
        easy_mode = bool(self.easy_mode_var.get())
        if easy_mode:
            self._apply_easy_mode_defaults()
        for widget in self._advanced_workflow_widgets:
            try:
                widget.configure(state="disabled" if easy_mode else "normal")
            except tk.TclError:
                continue
        self._refresh_workflow_validation_hints()

    def _selected_listbox_values(self, listbox: tk.Listbox) -> list[str]:
        return [str(listbox.get(index)) for index in listbox.curselection()]

    def _set_listbox_selection(
        self,
        listbox: tk.Listbox,
        values: list[str],
        *,
        valid_options: tuple[str, ...],
    ) -> None:
        filtered = [item for item in values if item in valid_options]
        options = [str(listbox.get(index)) for index in range(listbox.size())]
        listbox.selection_clear(0, "end")
        for item in filtered:
            if item in options:
                listbox.selection_set(options.index(item))

    def _sync_signal_selection_vars(self) -> None:
        self.entry_signals_var.set(", ".join(self._selected_listbox_values(self.entry_signals_listbox)))
        self.exit_signals_var.set(", ".join(self._selected_listbox_values(self.exit_signals_listbox)))
        self.benchmark_selection_var.set(", ".join(self._selected_listbox_values(self.benchmark_selection_listbox)))

    def _on_structured_selection_change(self, _field: str) -> None:
        self._sync_signal_selection_vars()
        self._refresh_workflow_validation_hints()

    def _structured_or_text_selection(
        self,
        *,
        listbox: tk.Listbox,
        text_var: tk.StringVar,
        valid_options: tuple[str, ...],
        field_name: str,
        show_popup: bool = True,
    ) -> list[str] | None:
        structured = self._selected_listbox_values(listbox)
        if structured:
            return structured
        return self._parse_signal_csv(
            text_var.get().strip(),
            valid_options=valid_options,
            field_name=field_name,
            show_popup=show_popup,
        )

    def _recommended_values_for_field(self, field: str) -> list[str]:
        preset = self._resolve_workflow_preset(self.workflow_preset_var.get().strip())
        if preset is None:
            return []
        key_by_field = {
            "entry": "entry_signals",
            "exit": "exit_signals",
            "benchmark": "benchmark_selection",
        }
        key = key_by_field.get(field)
        if key is None:
            return []
        values = preset.get(key, [])
        if not isinstance(values, list):
            return []
        return [str(value) for value in values]

    def _select_recommended_set(self, field: str) -> None:
        recommended = self._recommended_values_for_field(field)
        if field == "entry":
            self._set_listbox_selection(self.entry_signals_listbox, recommended, valid_options=self._signal_options)
        elif field == "exit":
            self._set_listbox_selection(self.exit_signals_listbox, recommended, valid_options=self._exit_signal_options)
        elif field == "benchmark":
            self._set_listbox_selection(self.benchmark_selection_listbox, recommended, valid_options=self._benchmark_options)
        self._sync_signal_selection_vars()
        self._refresh_workflow_validation_hints()

    def _attach_workflow_validation_traces(self) -> None:
        watched_vars = [
            self.entry_signals_var,
            self.exit_signals_var,
            self.benchmark_selection_var,
            self.optimization_trials_var,
            self.optimization_sampler_var,
            self.optimization_min_completed_var,
            self.optimization_staged_budgets_var,
            self.wf_train_fraction_var,
            self.wf_validation_fraction_var,
            self.wf_test_fraction_var,
            self.wf_step_fraction_var,
            self.wf_split_policy_var,
        ]
        for variable in watched_vars:
            variable.trace_add("write", lambda *_: self._refresh_workflow_validation_hints())

    def _refresh_workflow_validation_hints(self) -> None:
        issues = self._collect_workflow_validation_issues()
        if not issues:
            self._workflow_validation_var.set("All workflow parameters look valid.")
            return
        self._workflow_validation_var.set("Validation hints before submit:\n- " + "\n- ".join(issues))

    def _collect_workflow_validation_issues(self) -> list[str]:
        issues: list[str] = []
        if self._structured_or_text_selection(
            listbox=self.entry_signals_listbox,
            text_var=self.entry_signals_var,
            valid_options=self._signal_options,
            field_name="Entry signals",
            show_popup=False,
        ) is None:
            issues.append("Entry signals must include at least one supported signal.")
        if self._structured_or_text_selection(
            listbox=self.exit_signals_listbox,
            text_var=self.exit_signals_var,
            valid_options=self._exit_signal_options,
            field_name="Exit signals",
            show_popup=False,
        ) is None:
            issues.append("Exit signals must include at least one supported signal.")

        if self._structured_or_text_selection(
            listbox=self.benchmark_selection_listbox,
            text_var=self.benchmark_selection_var,
            valid_options=self._benchmark_options,
            field_name="Benchmarks",
            show_popup=False,
        ) is None:
            issues.append("Benchmarks must include at least one supported selection.")

        n_trials = int(parse_float(self.optimization_trials_var.get()) or 20)
        if n_trials <= 0:
            issues.append("Optimization trials must be greater than zero.")

        sampler = self.optimization_sampler_var.get().strip().lower() or "tpe"
        if sampler not in self._sampler_options:
            issues.append("Optimization sampler must be one of: tpe, cma-es, random, grid.")

        min_completed_for_pruning = int(parse_float(self.optimization_min_completed_var.get()) or 5)
        if min_completed_for_pruning < 1:
            issues.append("Min completed for pruning must be >= 1.")

        staged_budgets_raw = self.optimization_staged_budgets_var.get().strip()
        if staged_budgets_raw:
            try:
                staged_payload = json.loads(staged_budgets_raw)
                if not isinstance(staged_payload, list):
                    issues.append("Staged budgets must be a JSON list.")
            except json.JSONDecodeError:
                issues.append("Staged budgets must be valid JSON list.")

        train_fraction = float(parse_float(self.wf_train_fraction_var.get()) or 0.70)
        validation_fraction = float(parse_float(self.wf_validation_fraction_var.get()) or 0.15)
        test_fraction = float(parse_float(self.wf_test_fraction_var.get()) or 0.15)
        step_fraction = float(parse_float(self.wf_step_fraction_var.get()) or 0.15)
        if any(frac <= 0.0 for frac in (train_fraction, validation_fraction, test_fraction, step_fraction)):
            issues.append("Walk-forward fractions must all be positive.")
        if abs((train_fraction + validation_fraction + test_fraction) - 1.0) > 1e-6:
            issues.append("Walk-forward train + validation + test fractions must sum to 1.0.")

        split_policy = self.wf_split_policy_var.get().strip().lower()
        if split_policy not in {"calendar-based", "volatility-regime-stratified", "event-exclusion windows"}:
            issues.append("Walk-forward split policy selection is invalid.")
        return issues

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
            "universe_filters": self._build_universe_filters(),
            "rubric_profile": self.hypothesis_rubric_profile_var.get().strip() or "intraday_alpha",
            "hypothesis_novelty": float(parse_float(self.hypothesis_novelty_var.get()) or 3.0),
            "hypothesis_plausibility": float(parse_float(self.hypothesis_plausibility_var.get()) or 3.0),
            "hypothesis_implementation_complexity": float(
                parse_float(self.hypothesis_implementation_complexity_var.get()) or 3.0
            ),
            "hypothesis_expected_capacity": float(parse_float(self.hypothesis_expected_capacity_var.get()) or 3.0),
            "hypothesis_robustness": float(parse_float(self.hypothesis_robustness_var.get()) or 3.0),
        }

    def _build_universe_filters(self) -> dict[str, Any]:
        def _split_csv(raw: str) -> list[str]:
            return [item.strip() for item in raw.split(",") if item.strip()]

        return {
            "sector": {
                "include": _split_csv(self.wizard_sector_include_var.get().strip()),
                "exclude": _split_csv(self.wizard_sector_exclude_var.get().strip()),
            },
            "liquidity": {
                "min_adv": float(parse_float(self.wizard_adv_threshold_var.get()) or 0.0),
                "min_liquidity": float(parse_float(self.wizard_liquidity_threshold_var.get()) or 0.0),
            },
            "price_market_cap": {
                "min_price": float(parse_float(self.wizard_price_min_var.get()) or 0.0),
                "max_price": float(parse_float(self.wizard_price_max_var.get()) or 0.0),
                "min_market_cap": float(parse_float(self.wizard_market_cap_min_var.get()) or 0.0),
                "max_market_cap": float(parse_float(self.wizard_market_cap_max_var.get()) or 0.0),
            },
            "options_eligibility": {
                "min_open_interest": int(parse_float(self.wizard_min_option_oi_var.get()) or 0),
                "min_option_volume": int(parse_float(self.wizard_min_option_volume_var.get()) or 0),
                "min_days_to_expiration": int(parse_float(self.wizard_min_option_dte_var.get()) or 0),
                "require_weeklies": bool(self.wizard_require_weeklies_var.get()),
            },
        }

    def _load_rubric_templates(self) -> dict[str, Any]:
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        if not self._hypothesis_rubric_templates_path.exists():
            self._hypothesis_rubric_templates_path.write_text(
                json.dumps(DEFAULT_HYPOTHESIS_RUBRIC_TEMPLATES, indent=2, sort_keys=True),
                encoding="utf-8",
            )
            return dict(DEFAULT_HYPOTHESIS_RUBRIC_TEMPLATES)

        try:
            payload = json.loads(self._hypothesis_rubric_templates_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return dict(DEFAULT_HYPOTHESIS_RUBRIC_TEMPLATES)

        if not isinstance(payload, dict) or "profiles" not in payload:
            return dict(DEFAULT_HYPOTHESIS_RUBRIC_TEMPLATES)
        return payload

    def _resolve_profile_defaults(self, profile_name: str) -> dict[str, float]:
        profile = self._resolve_rubric_profile(profile_name)
        defaults = profile.get("defaults", {})
        return {
            "novelty": float(defaults.get("novelty", 3.0)),
            "plausibility": float(defaults.get("plausibility", 3.0)),
            "implementation_complexity": float(defaults.get("implementation_complexity", 3.0)),
            "expected_capacity": float(defaults.get("expected_capacity", 3.0)),
            "robustness": float(defaults.get("robustness", 3.0)),
        }

    def _resolve_rubric_profile(self, profile_name: str) -> dict[str, Any]:
        profiles = self._rubric_templates.get("profiles", {})
        if profile_name in profiles and isinstance(profiles[profile_name], dict):
            return profiles[profile_name]
        fallback = str(self._rubric_templates.get("default_profile", "intraday_alpha"))
        if fallback in profiles and isinstance(profiles[fallback], dict):
            return profiles[fallback]
        return DEFAULT_HYPOTHESIS_RUBRIC_TEMPLATES["profiles"]["intraday_alpha"]

    def _load_workflow_presets(self) -> dict[str, Any]:
        self._workflow_preset_warnings: list[str] = []
        if not RESEARCH_LAB_PRESETS_PATH.exists():
            self._workflow_preset_warnings.append(
                f"Preset file not found at {RESEARCH_LAB_PRESETS_PATH}; using built-in defaults."
            )
            return dict(DEFAULT_RESEARCH_WORKFLOW_PRESETS)
        try:
            payload = json.loads(RESEARCH_LAB_PRESETS_PATH.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            self._workflow_preset_warnings.append(
                f"Failed to parse preset file {RESEARCH_LAB_PRESETS_PATH}; using built-in defaults."
            )
            return dict(DEFAULT_RESEARCH_WORKFLOW_PRESETS)

        result = validate_workflow_preset_payload(
            payload,
            fallback_payload=DEFAULT_RESEARCH_WORKFLOW_PRESETS,
        )
        self._workflow_preset_warnings.extend(result.warnings)
        return result.payload

    def _format_preset_warning_text(self) -> str:
        warnings = getattr(self, "_workflow_preset_warnings", [])
        if not warnings:
            return ""
        preview = warnings[:3]
        lines = [f"⚠ Preset validation: {item}" for item in preview]
        if len(warnings) > len(preview):
            lines.append(f"⚠ ... and {len(warnings) - len(preview)} more warning(s).")
        return "\n".join(lines)

    def _resolve_workflow_preset(self, preset_name: str) -> dict[str, Any] | None:
        presets = self._workflow_presets.get("presets", {})
        preset = presets.get(preset_name)
        return preset if isinstance(preset, dict) else None

    def _on_workflow_preset_change(self, *_: object) -> None:
        # Selection is explicit; values are applied when button is clicked.
        return

    def _apply_selected_workflow_preset(self) -> None:
        preset_name = self.workflow_preset_var.get().strip()
        preset = self._resolve_workflow_preset(preset_name)
        if preset is None:
            return

        entry_signals = preset.get("entry_signals", [])
        exit_signals = preset.get("exit_signals", [])
        optimization = preset.get("optimization", {})
        walk_forward = preset.get("walk_forward", {})
        stress_controls = preset.get("stress_controls", {})
        benchmark_selection = preset.get("benchmark_selection", list(self._benchmark_options))

        if isinstance(entry_signals, list) and entry_signals:
            self._set_listbox_selection(
                self.entry_signals_listbox,
                [str(signal) for signal in entry_signals],
                valid_options=self._signal_options,
            )
        if isinstance(exit_signals, list) and exit_signals:
            self._set_listbox_selection(
                self.exit_signals_listbox,
                [str(signal) for signal in exit_signals],
                valid_options=self._exit_signal_options,
            )
        if isinstance(benchmark_selection, list) and benchmark_selection:
            self._set_listbox_selection(
                self.benchmark_selection_listbox,
                [str(item) for item in benchmark_selection],
                valid_options=self._benchmark_options,
            )
        self._sync_signal_selection_vars()

        self.optimization_trials_var.set(str(optimization.get("n_trials", self.optimization_trials_var.get())))
        self.optimization_sampler_var.set(str(optimization.get("sampler", self.optimization_sampler_var.get())))
        self.optimization_enable_pruning_var.set(bool(optimization.get("enable_pruning", self.optimization_enable_pruning_var.get())))
        self.optimization_prune_constraint_var.set(bool(optimization.get("prune_on_constraint", self.optimization_prune_constraint_var.get())))
        self.optimization_prune_lcb_var.set(bool(optimization.get("prune_on_lcb", self.optimization_prune_lcb_var.get())))
        self.optimization_min_completed_var.set(str(optimization.get("min_completed_for_pruning", self.optimization_min_completed_var.get())))
        self.optimization_staged_budgets_var.set(json.dumps(optimization.get("staged_budgets", json.loads(self.optimization_staged_budgets_var.get())), separators=(",", ":")))

        self.wf_train_fraction_var.set(f"{float(walk_forward.get('train_fraction', self.wf_train_fraction_var.get())):.2f}")
        self.wf_validation_fraction_var.set(f"{float(walk_forward.get('validation_fraction', self.wf_validation_fraction_var.get())):.2f}")
        self.wf_test_fraction_var.set(f"{float(walk_forward.get('test_fraction', self.wf_test_fraction_var.get())):.2f}")
        self.wf_step_fraction_var.set(f"{float(walk_forward.get('step_fraction', self.wf_step_fraction_var.get())):.2f}")
        self.wf_split_policy_var.set(str(walk_forward.get("split_policy", self.wf_split_policy_var.get())))

        self.stress_enable_historical_replay_var.set(bool(stress_controls.get("enable_historical_replay_regimes", self.stress_enable_historical_replay_var.get())))
        self.stress_historical_window_fraction_var.set(f"{float(stress_controls.get('historical_window_fraction', self.stress_historical_window_fraction_var.get())):.2f}")
        self.stress_historical_replay_window_bars_var.set(str(stress_controls.get("historical_replay_window_bars", self.stress_historical_replay_window_bars_var.get())))
        self.stress_synthetic_jump_magnitude_var.set(f"{float(stress_controls.get('synthetic_jump_magnitude', self.stress_synthetic_jump_magnitude_var.get())):.2f}")
        self.stress_synthetic_jump_interval_var.set(str(stress_controls.get("synthetic_jump_interval", self.stress_synthetic_jump_interval_var.get())))
        self.stress_synthetic_vol_cluster_multiplier_var.set(f"{float(stress_controls.get('synthetic_vol_cluster_multiplier', self.stress_synthetic_vol_cluster_multiplier_var.get())):.2f}")
        self.stress_overlay_spread_multiplier_var.set(f"{float(stress_controls.get('overlay_spread_multiplier', self.stress_overlay_spread_multiplier_var.get())):.2f}")
        self.stress_overlay_liquidity_multiplier_var.set(f"{float(stress_controls.get('overlay_liquidity_multiplier', self.stress_overlay_liquidity_multiplier_var.get())):.2f}")
        self._refresh_workflow_validation_hints()

    def run_walk_forward(self) -> None:
        self._start_worker(self._run_walk_forward_workflow, "Walk-forward validation")

    def run_optimization(self) -> None:
        self._start_worker(self._run_optimization_workflow, "Parameter optimization")

    def run_stress_tests(self) -> None:
        self._start_worker(self._run_stress_workflow, "Stress/scenario tests")

    def run_hypothesis_pipeline(self) -> None:
        self._start_worker(self._run_hypothesis_pipeline_workflow, "Hypothesis intake pipeline")

    def open_governance_workspace(self) -> None:
        self.controller.show_frame("BacktestingPage")

    def _parse_signal_csv(
        self,
        raw_text: str,
        *,
        valid_options: tuple[str, ...],
        field_name: str,
        show_popup: bool = True,
    ) -> list[str] | None:
        parsed = [item.strip() for item in raw_text.split(",") if item.strip()]
        if not parsed:
            if show_popup:
                messagebox.showinfo("Invalid input", f"{field_name} must include at least one signal.")
            return None
        invalid = [item for item in parsed if item not in valid_options]
        if invalid:
            if show_popup:
                messagebox.showinfo("Invalid input", f"Unsupported {field_name.lower()}: {', '.join(invalid)}")
            return None
        return parsed

    def _build_workflow_config(self) -> ResearchWorkflowConfig | None:
        entry_signals = self._structured_or_text_selection(
            listbox=self.entry_signals_listbox,
            text_var=self.entry_signals_var,
            valid_options=self._signal_options,
            field_name="Entry signals",
        )
        if entry_signals is None:
            return None

        exit_signals = self._structured_or_text_selection(
            listbox=self.exit_signals_listbox,
            text_var=self.exit_signals_var,
            valid_options=self._exit_signal_options,
            field_name="Exit signals",
        )
        if exit_signals is None:
            return None

        benchmark_selection = self._structured_or_text_selection(
            listbox=self.benchmark_selection_listbox,
            text_var=self.benchmark_selection_var,
            valid_options=self._benchmark_options,
            field_name="Benchmarks",
        )
        if benchmark_selection is None:
            return None

        n_trials = int(parse_float(self.optimization_trials_var.get()) or 20)
        if n_trials <= 0:
            messagebox.showinfo("Invalid input", "Optimization trials must be greater than zero.")
            return None

        sampler = self.optimization_sampler_var.get().strip().lower() or "tpe"
        if sampler not in self._sampler_options:
            messagebox.showinfo("Invalid input", "Optimization sampler must be one of: tpe, cma-es, random, grid.")
            return None

        min_completed_for_pruning = int(parse_float(self.optimization_min_completed_var.get()) or 5)
        if min_completed_for_pruning < 1:
            messagebox.showinfo("Invalid input", "Min completed for pruning must be >= 1.")
            return None

        staged_budgets: list[dict[str, object]] = []
        staged_budgets_raw = self.optimization_staged_budgets_var.get().strip()
        if staged_budgets_raw:
            try:
                staged_payload = json.loads(staged_budgets_raw)
            except json.JSONDecodeError:
                messagebox.showinfo("Invalid input", "Staged budgets must be valid JSON list.")
                return None
            if not isinstance(staged_payload, list):
                messagebox.showinfo("Invalid input", "Staged budgets must be a JSON list.")
                return None
            staged_budgets = [dict(item) for item in staged_payload if isinstance(item, dict)]

        train_fraction = float(parse_float(self.wf_train_fraction_var.get()) or 0.70)
        validation_fraction = float(parse_float(self.wf_validation_fraction_var.get()) or 0.15)
        test_fraction = float(parse_float(self.wf_test_fraction_var.get()) or 0.15)
        step_fraction = float(parse_float(self.wf_step_fraction_var.get()) or 0.15)

        if any(frac <= 0.0 for frac in (train_fraction, validation_fraction, test_fraction, step_fraction)):
            messagebox.showinfo("Invalid input", "Walk-forward fractions must all be positive.")
            return None
        if abs((train_fraction + validation_fraction + test_fraction) - 1.0) > 1e-6:
            messagebox.showinfo("Invalid input", "Walk-forward train + validation + test fractions must sum to 1.0.")
            return None

        split_policy = self.wf_split_policy_var.get().strip().lower()
        if split_policy not in {"calendar-based", "volatility-regime-stratified", "event-exclusion windows"}:
            messagebox.showinfo("Invalid input", "Walk-forward split policy selection is invalid.")
            return None

        stress_controls = {
            "enable_historical_replay_regimes": bool(self.stress_enable_historical_replay_var.get()),
            "historical_window_fraction": float(parse_float(self.stress_historical_window_fraction_var.get()) or 0.20),
            "historical_replay_window_bars": int(parse_float(self.stress_historical_replay_window_bars_var.get()) or 20),
            "synthetic_jump_magnitude": float(parse_float(self.stress_synthetic_jump_magnitude_var.get()) or 0.02),
            "synthetic_jump_interval": int(parse_float(self.stress_synthetic_jump_interval_var.get()) or 7),
            "synthetic_vol_cluster_multiplier": float(parse_float(self.stress_synthetic_vol_cluster_multiplier_var.get()) or 1.6),
            "overlay_spread_multiplier": float(parse_float(self.stress_overlay_spread_multiplier_var.get()) or 2.5),
            "overlay_liquidity_multiplier": float(parse_float(self.stress_overlay_liquidity_multiplier_var.get()) or 0.4),
        }

        return ResearchWorkflowConfig(
            preset_name=self.workflow_preset_var.get().strip(),
            entry_signals=entry_signals,
            exit_signals=exit_signals,
            optimization_n_trials=n_trials,
            optimization_sampler=sampler,
            optimization_enable_pruning=bool(self.optimization_enable_pruning_var.get()),
            optimization_prune_on_constraint=bool(self.optimization_prune_constraint_var.get()),
            optimization_prune_on_lcb=bool(self.optimization_prune_lcb_var.get()),
            optimization_min_completed_for_pruning=min_completed_for_pruning,
            optimization_staged_budgets=staged_budgets,
            train_fraction=train_fraction,
            validation_fraction=validation_fraction,
            test_fraction=test_fraction,
            step_fraction=step_fraction,
            walk_forward_split_policy=split_policy,
            stress_controls=stress_controls,
            benchmark_selection=benchmark_selection,
        )

    def _build_signal_grids(
        self,
        context: dict[str, Any],
        config: ResearchWorkflowConfig,
    ) -> tuple[dict[str, list[dict[str, object]]], dict[str, list[dict[str, object]]], dict[str, list[object]]]:
        entry_grid = {signal: [{}] for signal in config.entry_signals}
        exit_grid = {signal: [{}] for signal in config.exit_signals}
        core_grid = {
            "lookback_days": [int(context["lookback"])],
            "skip_days": [int(context["skip"])],
            "costs_bps": [float(context["costs_bps"])],
            "universe_filters": [dict(context.get("universe_filters", {}))],
        }
        return entry_grid, exit_grid, core_grid

    def _run_walk_forward_workflow(self, context: dict[str, Any], config: ResearchWorkflowConfig, cancellation_token: CancellationToken) -> str:
        entry_grid, exit_grid, core_grid = self._build_signal_grids(context, config)
        return run_walk_forward_backtest(
            tickers=list(context["tickers"]),
            start_date=context["start_date"],
            end_date=context["end_date"],
            cache_root=context["cache_root"],
            entry_grid=entry_grid,
            exit_grid=exit_grid,
            core_grid=core_grid,
            train_fraction=config.train_fraction,
            validation_fraction=config.validation_fraction,
            test_fraction=config.test_fraction,
            step_fraction=config.step_fraction,
            split_policy=config.walk_forward_split_policy,
            governance_metadata={
                "promotion_state": "research",
                "approval_status": "pending",
                "workflow_preset": config.preset_name,
                "universe_filters": dict(context.get("universe_filters", {})),
            },
            stress_controls=dict(config.stress_controls),
            benchmarks=list(config.benchmark_selection),
            cancellation_token=cancellation_token,
            run_namespace=str(context.get("run_namespace", "")).strip() or None,
        )

    def _run_optimization_workflow(self, context: dict[str, Any], config: ResearchWorkflowConfig, cancellation_token: CancellationToken) -> str:
        entry_grid, exit_grid, core_grid = self._build_signal_grids(context, config)
        return run_strategy_optimization(
            tickers=list(context["tickers"]),
            start_date=context["start_date"],
            end_date=context["end_date"],
            cache_root=context["cache_root"],
            entry_grid=entry_grid,
            exit_grid=exit_grid,
            core_grid=core_grid,
            seed=42,
            n_trials=config.optimization_n_trials,
            sampler_name=config.optimization_sampler,
            partial_period_fractions=[0.5, 1.0],
            enable_pruning=config.optimization_enable_pruning,
            prune_on_constraint_violation=config.optimization_prune_on_constraint,
            prune_on_lcb=config.optimization_prune_on_lcb,
            min_completed_for_pruning=config.optimization_min_completed_for_pruning,
            staged_budgets=[dict(stage) for stage in config.optimization_staged_budgets],
            governance_metadata={
                "promotion_state": "research",
                "approval_status": "pending",
                "workflow_preset": config.preset_name,
                "universe_filters": dict(context.get("universe_filters", {})),
            },
            stress_controls=dict(config.stress_controls),
            benchmarks=list(config.benchmark_selection),
            cancellation_token=cancellation_token,
            run_namespace=str(context.get("run_namespace", "")).strip() or None,
        )

    def _run_stress_workflow(self, context: dict[str, Any], config: ResearchWorkflowConfig, cancellation_token: CancellationToken) -> str:
        return run_multi_signal_backtest(
            tickers=list(context["tickers"]),
            start_date=context["start_date"],
            end_date=context["end_date"],
            cache_root=context["cache_root"],
            lookback_days=int(context["lookback"]),
            skip_days=int(context["skip"]),
            costs_bps=float(context["costs_bps"]),
            timeframe=str(DEFAULT_BACKTEST_SETTINGS.get("timeframe", "1m")),
            entry_signals=list(config.entry_signals),
            exit_signals=list(config.exit_signals),
            governance_metadata={
                "promotion_state": "research",
                "approval_status": "pending",
                "workflow_preset": config.preset_name,
                "universe_filters": dict(context.get("universe_filters", {})),
            },
            stress_controls=dict(config.stress_controls),
            benchmarks=list(config.benchmark_selection),
            cancellation_token=cancellation_token,
            run_namespace=str(context.get("run_namespace", "")).strip() or None,
        )

    def _run_hypothesis_pipeline_workflow(self, context: dict[str, Any], config: ResearchWorkflowConfig, cancellation_token: CancellationToken) -> str:
        self._research_lab_dir.mkdir(parents=True, exist_ok=True)
        hypothesis_id = f"hyp_{uuid.uuid4().hex}"
        idea_record = self._build_idea_record(hypothesis_id=hypothesis_id, context=context)
        idea_record["lineage"] = self._build_pipeline_lineage(hypothesis_id=hypothesis_id)
        scored = self._score_hypothesis(idea_record, context)
        promoted = scored["decision"] == "accept"
        experiment_path = self._write_experiment_skeleton(scored, context) if promoted else None
        self._append_funnel_event(scored, context)
        funnel = self._compute_funnel_metrics()
        self._refresh_funnel_kpi_dashboard()

        lines = [
            f"Idea intake: {scored['idea']['title']}",
            f"Hypothesis ID: {scored['hypothesis_id']}",
            f"Decision: {scored['decision'].upper()} ({scored['decision_reason']})",
            f"Rubric total score: {scored['rubric']['total_score']:.2f}/5.00",
            f"Funnel conversion to accepted: {funnel['acceptance_rate_pct']:.1f}%",
            (
                f"Generated experiment skeleton: {experiment_path}"
                if experiment_path
                else "Experiment skeleton skipped (hypothesis rejected)."
            ),
        ]
        return "\n".join(lines)

    def _build_pipeline_lineage(self, *, hypothesis_id: str) -> dict[str, Any]:
        optimization_run_id = f"opt_{uuid.uuid4().hex}"
        walk_forward_run_id = f"wf_{uuid.uuid4().hex}"
        stress_run_id = f"stress_{uuid.uuid4().hex}"
        return {
            "hypothesis_id": hypothesis_id,
            "optimization_run_id": optimization_run_id,
            "walk_forward_run_id": walk_forward_run_id,
            "stress_run_id": stress_run_id,
            "links": {
                "hypothesis_to_optimization": {
                    "parent_id": hypothesis_id,
                    "child_id": optimization_run_id,
                },
                "optimization_to_walk_forward": {
                    "parent_id": optimization_run_id,
                    "child_id": walk_forward_run_id,
                },
                "walk_forward_to_stress": {
                    "parent_id": walk_forward_run_id,
                    "child_id": stress_run_id,
                },
            },
        }

    def _build_idea_record(self, *, hypothesis_id: str, context: dict[str, Any]) -> dict[str, Any]:
        tickers = [str(t).upper() for t in context["tickers"]]
        return {
            "hypothesis_id": hypothesis_id,
            "idea": {
                "title": f"Momentum continuation with breakout confirmation on {', '.join(tickers[:3])}",
                "description": "Combine ts_momentum with breakout confirmation to reduce false positives.",
                "submitter": "research_lab_ui",
                "submitted_at": date.today().isoformat(),
            },
            "economic_rationale": {
                "market_inefficiency": "Underreaction to persistent trend shocks.",
                "causal_mechanism": "Delayed repricing by slower participants under changing volatility regimes.",
                "why_now": "Elevated regime shifts amplify trend persistence and breakout follow-through.",
                "failure_modes": ["Mean-reversion regime", "Execution slippage during volatility spikes"],
            },
            "data_requirements": {
                "assets": tickers,
                "start_date": str(context["start_date"]),
                "end_date": str(context["end_date"]),
                "universe_filters": dict(context.get("universe_filters", {})),
                "fields": ["open", "high", "low", "close", "volume"],
                "quality_checks": ["missing_bar_ratio < 1%", "corporate_action_adjusted", "timezone_normalized"],
            },
            "test_design": {
                "primary_test": "walk_forward",
                "secondary_tests": ["stress_scenarios", "parameter_sensitivity"],
                "acceptance_criteria": {
                    "min_sharpe": 0.8,
                    "max_drawdown": -0.25,
                    "min_stability_score": 0.55,
                },
                "reproducibility": {"seed": 42, "train_fraction": 0.7, "validation_fraction": 0.15, "test_fraction": 0.15},
            },
            "results_review": {
                "reviewer": "research_lab_ui",
                "status": "pending",
                "notes": "Awaiting execution of generated experiment skeleton.",
            },
            "promotion_or_rejection": {
                "state": "pending",
                "approval_status": "pending",
                "reason": "Pending rubric scoring and test execution.",
            },
        }

    def _score_hypothesis(self, record: dict[str, Any], context: dict[str, Any]) -> dict[str, Any]:
        profile_name = str(context.get("rubric_profile", "intraday_alpha"))
        profile = self._resolve_rubric_profile(profile_name)
        user_scores = {
            "novelty": float(context.get("hypothesis_novelty", profile["defaults"]["novelty"])),
            "plausibility": float(context.get("hypothesis_plausibility", profile["defaults"]["plausibility"])),
            "implementation_complexity": float(
                context.get("hypothesis_implementation_complexity", profile["defaults"]["implementation_complexity"])
            ),
            "expected_capacity": float(context.get("hypothesis_expected_capacity", profile["defaults"]["expected_capacity"])),
            "robustness": float(context.get("hypothesis_robustness", profile["defaults"]["robustness"])),
        }

        complexity_adjusted = max(0.0, 6.0 - user_scores["implementation_complexity"])
        weights = profile["weights"]
        total = (
            user_scores["novelty"] * float(weights["novelty"])
            + user_scores["plausibility"] * float(weights["plausibility"])
            + complexity_adjusted * float(weights["complexity_adjusted"])
            + user_scores["expected_capacity"] * float(weights["expected_capacity"])
            + user_scores["robustness"] * float(weights["robustness"])
        )

        thresholds = profile["thresholds"]
        accepted = (
            total >= float(thresholds["min_total"])
            and user_scores["plausibility"] >= float(thresholds["min_plausibility"])
            and user_scores["robustness"] >= float(thresholds["min_robustness"])
        )
        decision = "accept" if accepted else "reject"
        reason = "Clears minimum weighted rubric thresholds" if accepted else "Fails weighted rubric threshold"
        record["rubric"] = {
            **user_scores,
            "profile": profile_name,
            "weights": dict(weights),
            "thresholds": dict(thresholds),
            "complexity_adjusted": complexity_adjusted,
            "total_score": total,
        }
        record["decision"] = decision
        record["decision_reason"] = reason
        record["promotion_or_rejection"] = {
            "state": "promoted_to_experiment" if accepted else "rejected",
            "approval_status": "pending" if accepted else "rejected",
            "reason": reason,
        }
        return record

    def _write_experiment_skeleton(self, record: dict[str, Any], context: dict[str, Any]) -> Path:
        skeleton_dir = self._research_lab_dir / "experiment_skeletons" / str(record["hypothesis_id"])
        skeleton_dir.mkdir(parents=True, exist_ok=True)
        output_path = skeleton_dir / "research_project.json"
        generated_at = date.today().isoformat()
        lineage = dict(record.get("lineage", {}))
        optimization_run_id = str(lineage.get("optimization_run_id", "")).strip()
        walk_forward_run_id = str(lineage.get("walk_forward_run_id", "")).strip()
        stress_run_id = str(lineage.get("stress_run_id", "")).strip()
        run_artifacts_dir = skeleton_dir / "run_artifacts"

        def _artifact_path_for(run_id: str, suffix: str) -> str | None:
            if not run_id:
                return None
            return str(run_artifacts_dir / run_id / suffix)

        manifest = {
            "schema_version": "1.0",
            "hypothesis": {
                "hypothesis_id": record["hypothesis_id"],
                "title": str(record.get("idea", {}).get("title", "")),
                "description": str(record.get("idea", {}).get("description", "")),
                "submitter": str(record.get("idea", {}).get("submitter", "research_lab_ui")),
                "submitted_at": str(record.get("idea", {}).get("submitted_at", generated_at)),
                "economic_rationale": dict(record.get("economic_rationale", {})),
                "data_requirements": dict(record.get("data_requirements", {})),
                "test_design": dict(record.get("test_design", {})),
            },
            "lineage": lineage,
            "context": {
                "tickers": list(context["tickers"]),
                "start_date": str(context["start_date"]),
                "end_date": str(context["end_date"]),
                "lookback": int(context["lookback"]),
                "skip": int(context["skip"]),
                "costs_bps": float(context["costs_bps"]),
                "universe_filters": dict(context.get("universe_filters", {})),
            },
            "run_references": [
                {
                    "step": "parameter_optimization",
                    "call": "run_strategy_optimization",
                    "status": "todo",
                    "run_id": optimization_run_id or None,
                    "parent_id": lineage.get("hypothesis_id"),
                    "artifact_path": _artifact_path_for(optimization_run_id, "manifest.json"),
                    "logs_path": _artifact_path_for(optimization_run_id, "logs.txt"),
                },
                {
                    "step": "walk_forward_validation",
                    "call": "run_walk_forward_backtest",
                    "status": "todo",
                    "run_id": walk_forward_run_id or None,
                    "parent_id": optimization_run_id or None,
                    "artifact_path": _artifact_path_for(walk_forward_run_id, "manifest.json"),
                    "logs_path": _artifact_path_for(walk_forward_run_id, "logs.txt"),
                },
                {
                    "step": "stress_testing",
                    "call": "run_multi_signal_backtest",
                    "status": "todo",
                    "run_id": stress_run_id or None,
                    "parent_id": walk_forward_run_id or None,
                    "artifact_path": _artifact_path_for(stress_run_id, "manifest.json"),
                    "logs_path": _artifact_path_for(stress_run_id, "logs.txt"),
                },
            ],
            "gate_outcomes": [
                {
                    "gate": "rubric_score",
                    "status": str(record.get("decision", "pending")),
                    "reason": str(record.get("decision_reason", "")),
                    "score": float(record.get("rubric", {}).get("total_score", 0.0)),
                    "evaluated_at": generated_at,
                }
            ],
            "reviewer_comments": [
                {
                    "reviewer": str(record.get("results_review", {}).get("reviewer", "research_lab_ui")),
                    "status": str(record.get("results_review", {}).get("status", "pending")),
                    "comment": str(record.get("results_review", {}).get("notes", "")),
                    "commented_at": generated_at,
                }
            ],
            "comments": [
                {
                    "owner": str(item.get("owner", "research_lab_ui")),
                    "note": str(item.get("note", "")),
                    "timestamp": str(item.get("timestamp", generated_at)),
                }
                for item in getattr(self, "_wizard_comments", [])
                if str(item.get("note", "")).strip()
            ],
            "hypothesis_history": [
                {
                    "event": str(item.get("event", "update")),
                    "timestamp": str(item.get("timestamp", generated_at)),
                    "details": {
                        str(key): str(value)
                        for key, value in item.items()
                        if str(key) not in {"event", "timestamp"}
                    },
                }
                for item in getattr(self, "_wizard_history", [])
            ],
            "review_actions": [
                {
                    "owner": str(record.get("results_review", {}).get("reviewer", "research_lab_ui")),
                    "action": "review_completed",
                    "status": str(record.get("results_review", {}).get("status", "pending")),
                    "timestamp": generated_at,
                }
            ],
            "decision_log": [
                {
                    "owner": str(record.get("results_review", {}).get("reviewer", "research_lab_ui")),
                    "decision": str(record.get("decision", "pending")),
                    "reason": str(record.get("decision_reason", "")),
                    "timestamp": generated_at,
                }
            ],
            "promotion_history": [
                {
                    "state": str(record.get("promotion_or_rejection", {}).get("state", "pending")),
                    "approval_status": str(record.get("promotion_or_rejection", {}).get("approval_status", "pending")),
                    "reason": str(record.get("promotion_or_rejection", {}).get("reason", "")),
                    "recorded_at": generated_at,
                }
            ],
            "rubric": dict(record.get("rubric", {})),
            "execution_chain": build_default_research_execution_chain(),
            "generated_at": generated_at,
            "last_updated": generated_at,
        }
        manifest["pipeline_graph"] = self._build_pipeline_graph_payload(manifest, manifest_path=output_path)

        if output_path.exists():
            try:
                existing = json.loads(output_path.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                existing = {}
            if isinstance(existing, dict):
                for key in (
                    "run_references",
                    "gate_outcomes",
                    "reviewer_comments",
                    "promotion_history",
                    "comments",
                    "hypothesis_history",
                    "review_actions",
                    "decision_log",
                ):
                    merged = list(existing.get(key, []))
                    merged.extend(manifest[key])
                    deduped: list[dict[str, Any]] = []
                    seen: set[str] = set()
                    for row in merged:
                        fingerprint = json.dumps(row, sort_keys=True)
                        if fingerprint in seen:
                            continue
                        seen.add(fingerprint)
                        deduped.append(row)
                    manifest[key] = deduped

                for key in ("hypothesis", "lineage", "context", "rubric", "generated_at"):
                    if key in existing:
                        manifest[key] = existing[key]
                manifest["last_updated"] = generated_at

        manifest["pipeline_graph"] = self._build_pipeline_graph_payload(manifest, manifest_path=output_path)

        output_path.write_text(json.dumps(manifest, indent=2, sort_keys=True), encoding="utf-8")
        return output_path

    def _append_funnel_event(self, record: dict[str, Any], context: dict[str, Any] | None = None) -> None:
        event_path = self._research_lab_dir / "idea_funnel_events.jsonl"
        submitted_at = str(record.get("idea", {}).get("submitted_at", date.today().isoformat()))
        decision_at = date.today().isoformat()
        strategy_family = "unspecified"
        if context is not None:
            context_strategy = str(context.get("strategy_family", "")).strip()
            if context_strategy:
                strategy_family = context_strategy
        if strategy_family == "unspecified":
            strategy_family = str(record.get("test_design", {}).get("primary_test", "unspecified")).strip() or "unspecified"

        event = {
            "hypothesis_id": record["hypothesis_id"],
            "lineage": dict(record.get("lineage", {})),
            "date": date.today().isoformat(),
            "submitted_at": submitted_at,
            "decision_at": decision_at,
            "strategy_family": strategy_family,
            "decision": record["decision"],
            "promotion_state": str(record.get("promotion_or_rejection", {}).get("state", "pending")),
            "rubric_total": float(record["rubric"]["total_score"]),
            "stages": {
                "idea_intake": True,
                "economic_rationale": True,
                "data_requirements": True,
                "test_design": True,
                "results_review": True,
                "promotion_or_rejection": True,
            },
        }
        with event_path.open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(event, sort_keys=True) + "\n")

    def _compute_funnel_metrics(self) -> dict[str, Any]:
        event_path = self._research_lab_dir / "idea_funnel_events.jsonl"
        if not event_path.exists():
            return {
                "total_ideas": 0.0,
                "accepted_ideas": 0.0,
                "acceptance_rate_pct": 0.0,
                "median_time_to_decision_days": 0.0,
                "false_positive_rate_pct": 0.0,
                "pass_rates_by_strategy_family": [],
                "promotion_conversion_by_month": [],
            }

        total = 0
        accepted = 0
        decision_latencies: list[float] = []
        strategy_counts: dict[str, dict[str, int]] = {}
        month_counts: dict[str, dict[str, int]] = {}
        decisions_by_hypothesis: dict[str, list[tuple[date, str]]] = {}

        for line in event_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line:
                continue
            try:
                parsed = json.loads(line)
            except json.JSONDecodeError:
                continue
            total += 1
            decision = str(parsed.get("decision", "")).lower()
            if decision == "accept":
                accepted += 1

            submitted_at = self._parse_iso_date(parsed.get("submitted_at") or parsed.get("date"))
            decision_at = self._parse_iso_date(parsed.get("decision_at") or parsed.get("date"))
            if submitted_at is not None and decision_at is not None:
                decision_latencies.append(float((decision_at - submitted_at).days))

            strategy_family = str(parsed.get("strategy_family", "unspecified")).strip() or "unspecified"
            strategy_bucket = strategy_counts.setdefault(strategy_family, {"total": 0, "accepted": 0})
            strategy_bucket["total"] += 1
            if decision == "accept":
                strategy_bucket["accepted"] += 1

            decision_month = decision_at.strftime("%Y-%m") if decision_at is not None else "unknown"
            month_bucket = month_counts.setdefault(decision_month, {"total": 0, "promoted": 0})
            month_bucket["total"] += 1
            promotion_state = str(parsed.get("promotion_state", "")).lower().strip()
            if promotion_state.startswith("promoted"):
                month_bucket["promoted"] += 1

            hypothesis_id = str(parsed.get("hypothesis_id", "")).strip()
            if hypothesis_id:
                decision_points = decisions_by_hypothesis.setdefault(hypothesis_id, [])
                decision_date = decision_at or submitted_at
                if decision_date is not None:
                    decision_points.append((decision_date, decision))

        acceptance_rate = (accepted / total * 100.0) if total else 0.0
        median_latency = statistics.median(decision_latencies) if decision_latencies else 0.0

        accepted_hypotheses = 0
        accepted_then_rejected = 0
        for timeline in decisions_by_hypothesis.values():
            ordered = [decision for _dt, decision in sorted(timeline, key=lambda row: row[0])]
            if "accept" not in ordered:
                continue
            accepted_hypotheses += 1
            first_accept_index = ordered.index("accept")
            if any(decision == "reject" for decision in ordered[first_accept_index + 1 :]):
                accepted_then_rejected += 1
        false_positive_rate = (accepted_then_rejected / accepted_hypotheses * 100.0) if accepted_hypotheses else 0.0

        strategy_rates = []
        for family in sorted(strategy_counts):
            row = strategy_counts[family]
            family_total = row["total"]
            family_accepted = row["accepted"]
            strategy_rates.append(
                {
                    "strategy_family": family,
                    "total": family_total,
                    "accepted": family_accepted,
                    "acceptance_rate_pct": (family_accepted / family_total * 100.0) if family_total else 0.0,
                }
            )

        monthly_conversion = []
        for month in sorted(month_counts):
            row = month_counts[month]
            monthly_conversion.append(
                {
                    "month": month,
                    "total": row["total"],
                    "promoted": row["promoted"],
                    "promotion_conversion_pct": (row["promoted"] / row["total"] * 100.0) if row["total"] else 0.0,
                }
            )

        return {
            "total_ideas": float(total),
            "accepted_ideas": float(accepted),
            "acceptance_rate_pct": acceptance_rate,
            "median_time_to_decision_days": float(median_latency),
            "false_positive_rate_pct": false_positive_rate,
            "pass_rates_by_strategy_family": strategy_rates,
            "promotion_conversion_by_month": monthly_conversion,
        }

    def _parse_iso_date(self, value: Any) -> date | None:
        text = str(value).strip()
        if not text:
            return None
        try:
            return date.fromisoformat(text)
        except ValueError:
            return None

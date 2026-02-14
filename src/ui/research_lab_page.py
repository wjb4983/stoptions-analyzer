from __future__ import annotations

import threading
import json
import uuid
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Any, Callable

from backtesting.cache_runner import (
    run_multi_signal_backtest,
    run_strategy_optimization,
    run_walk_forward_backtest,
)
from config import (
    BACKTEST_OUTPUT_DIR,
    CONFIG_DIR,
    DEFAULT_BACKTEST_SETTINGS,
    DEFAULT_HYPOTHESIS_RUBRIC_TEMPLATES,
    HYPOTHESIS_RUBRIC_TEMPLATES_PATH,
)
from utils.parsing import normalize_cache_root, parse_date, parse_float


@dataclass
class ResearchWorkflowConfig:
    entry_signals: list[str]
    exit_signals: list[str]
    optimization_n_trials: int
    optimization_sampler: str
    train_fraction: float
    validation_fraction: float
    test_fraction: float
    step_fraction: float
    stress_controls: dict[str, object]


@dataclass
class ResearchTask:
    task_id: str
    label: str
    target: Callable[[dict[str, Any], ResearchWorkflowConfig], str]
    context: dict[str, Any]
    config: ResearchWorkflowConfig
    state: str = "queued"
    logs: list[str] | None = None
    cancel_requested: bool = False

    def __post_init__(self) -> None:
        if self.logs is None:
            self.logs = []


class ResearchLabPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller
        self._task_queue: list[ResearchTask] = []
        self._active_task_id: str | None = None
        self._research_lab_dir = BACKTEST_OUTPUT_DIR / "research_lab"
        self._sampler_options = ("tpe", "bayesian", "random")
        self._signal_options = ("ts_momentum", "ma_trend", "breakout")
        self._exit_signal_options = ("none", "momentum_flip", "trailing_stop", "max_hold")
        self._wizard_state_path = self._research_lab_dir / "wizard_state.json"
        self._hypothesis_rubric_templates_path = HYPOTHESIS_RUBRIC_TEMPLATES_PATH
        self._rubric_templates = self._load_rubric_templates()


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
        self.output_text = tk.Text(output_frame, height=8, wrap="word")
        self.output_text.pack(fill="both", expand=True, padx=8, pady=8)

        task_logs_frame = ttk.LabelFrame(self, text="Selected Task Logs")
        task_logs_frame.pack(fill="both", expand=True, padx=40, pady=(0, 20))
        self.task_logs_text = tk.Text(task_logs_frame, height=8, wrap="word")
        self.task_logs_text.pack(fill="both", expand=True, padx=8, pady=8)

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
        rows = [f"[{task.state}] {task.label}" for task in self._task_queue]
        self._task_list_var.set(rows)
        task = self._selected_task()
        self._cancel_task_button.configure(state="normal" if task and task.state in {"queued", "running"} else "disabled")
        self._retry_task_button.configure(state="normal" if task and task.state in {"failed", "canceled"} else "disabled")
        self._refresh_selected_task_logs()
        if hasattr(self, "_wizard_next_button"):
            self._wizard_refresh_nav_state()

    def _refresh_selected_task_logs(self) -> None:
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
        target: Callable[[dict[str, Any], ResearchWorkflowConfig], str],
        context: dict[str, Any],
        config: ResearchWorkflowConfig,
    ) -> None:
        task = ResearchTask(
            task_id=uuid.uuid4().hex,
            label=label,
            target=target,
            context=dict(context),
            config=config,
        )
        self._task_queue.append(task)
        self._refresh_task_queue_ui()
        self._run_next_task()

    def _run_next_task(self) -> None:
        if self._active_task_id is not None:
            return
        next_task = next((task for task in self._task_queue if task.state == "queued"), None)
        if next_task is None:
            return
        self._active_task_id = next_task.task_id
        next_task.state = "running"
        self._task_log(next_task, "Task started.")

        def worker(task: ResearchTask) -> None:
            if task.cancel_requested:
                self._schedule_ui_update(lambda: self._finish_task(task.task_id, "", canceled=True))
                return
            try:
                output = task.target(task.context, task.config)
                self._schedule_ui_update(lambda: self._finish_task(task.task_id, output))
            except Exception as exc:
                self._schedule_ui_update(lambda: self._finish_task(task.task_id, f"Research workflow failed: {exc}", failed=True))

        threading.Thread(target=worker, args=(next_task,), daemon=True).start()

    def _finish_task(self, task_id: str, output: str, *, failed: bool = False, canceled: bool = False) -> None:
        task = self._get_task_by_id(task_id)
        if task is None:
            return
        if canceled or task.cancel_requested:
            task.state = "canceled"
            task.logs.append("Task canceled.")
            self._append_output(f"[{task.label}] Task canceled.")
        elif failed:
            task.state = "failed"
            task.logs.append(output)
            self._append_output(f"[{task.label}] Failed: {output}")
        else:
            task.state = "succeeded"
            task.logs.append(output)
            self._append_output(f"[{task.label}] Succeeded.")
        self._active_task_id = None
        self._refresh_task_queue_ui()
        self._run_next_task()

    def _cancel_selected_task(self) -> None:
        task = self._selected_task()
        if task is None:
            return
        task.cancel_requested = True
        if task.state == "queued":
            task.state = "canceled"
            task.logs.append("Task canceled before execution.")
        else:
            task.logs.append("Cancellation requested. Task will be marked canceled when current run returns.")
        self._refresh_task_queue_ui()

    def _retry_selected_task(self) -> None:
        task = self._selected_task()
        if task is None or task.state not in {"failed", "canceled"}:
            return
        task.state = "queued"
        task.cancel_requested = False
        task.logs.append("Task re-queued for retry.")
        self._refresh_task_queue_ui()
        self._run_next_task()

    def _build_workflow_controls(self) -> None:
        controls = ttk.LabelFrame(self, text="Workflow Controls")
        controls.pack(fill="x", padx=40, pady=(6, 8))
        controls.columnconfigure(1, weight=1)

        self.entry_signals_var = tk.StringVar(value="ts_momentum, breakout")
        self.exit_signals_var = tk.StringVar(value="none, momentum_flip")
        self.optimization_trials_var = tk.StringVar(value="20")
        self.optimization_sampler_var = tk.StringVar(value="tpe")
        self.wf_train_fraction_var = tk.StringVar(value="0.70")
        self.wf_validation_fraction_var = tk.StringVar(value="0.15")
        self.wf_test_fraction_var = tk.StringVar(value="0.15")
        self.wf_step_fraction_var = tk.StringVar(value="0.15")

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

        row = 0
        ttk.Label(controls, text="Entry signals (comma-separated)").grid(row=row, column=0, sticky="w", padx=8, pady=5)
        ttk.Entry(controls, textvariable=self.entry_signals_var).grid(row=row, column=1, sticky="ew", padx=8, pady=5)

        row += 1
        ttk.Label(controls, text="Exit signals (comma-separated)").grid(row=row, column=0, sticky="w", padx=8, pady=5)
        ttk.Entry(controls, textvariable=self.exit_signals_var).grid(row=row, column=1, sticky="ew", padx=8, pady=5)

        row += 1
        ttk.Label(controls, text="Optimization trials / sampler").grid(row=row, column=0, sticky="w", padx=8, pady=5)
        optimization_row = ttk.Frame(controls)
        optimization_row.grid(row=row, column=1, sticky="w", padx=8, pady=5)
        ttk.Entry(optimization_row, textvariable=self.optimization_trials_var, width=8).pack(side="left")
        ttk.Label(optimization_row, text=" / ").pack(side="left")
        ttk.Combobox(
            optimization_row,
            textvariable=self.optimization_sampler_var,
            values=self._sampler_options,
            state="readonly",
            width=12,
        ).pack(side="left")

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

        row += 1
        ttk.Label(controls, text="Stress: Replay/Jump/Overlay").grid(row=row, column=0, sticky="w", padx=8, pady=5)
        stress_row = ttk.Frame(controls)
        stress_row.grid(row=row, column=1, sticky="w", padx=8, pady=5)
        ttk.Checkbutton(
            stress_row,
            text="Historical replay regimes",
            variable=self.stress_enable_historical_replay_var,
        ).pack(side="left")

        row += 1
        ttk.Label(controls, text="Stress window frac / replay bars").grid(row=row, column=0, sticky="w", padx=8, pady=5)
        stress_row2 = ttk.Frame(controls)
        stress_row2.grid(row=row, column=1, sticky="w", padx=8, pady=5)
        ttk.Entry(stress_row2, textvariable=self.stress_historical_window_fraction_var, width=8).pack(side="left")
        ttk.Label(stress_row2, text=" / ").pack(side="left")
        ttk.Entry(stress_row2, textvariable=self.stress_historical_replay_window_bars_var, width=8).pack(side="left")

        row += 1
        ttk.Label(controls, text="Stress jump mag / interval / vol cluster").grid(row=row, column=0, sticky="w", padx=8, pady=5)
        stress_row3 = ttk.Frame(controls)
        stress_row3.grid(row=row, column=1, sticky="w", padx=8, pady=5)
        ttk.Entry(stress_row3, textvariable=self.stress_synthetic_jump_magnitude_var, width=8).pack(side="left")
        ttk.Label(stress_row3, text=" / ").pack(side="left")
        ttk.Entry(stress_row3, textvariable=self.stress_synthetic_jump_interval_var, width=8).pack(side="left")
        ttk.Label(stress_row3, text=" / ").pack(side="left")
        ttk.Entry(stress_row3, textvariable=self.stress_synthetic_vol_cluster_multiplier_var, width=8).pack(side="left")

        row += 1
        ttk.Label(controls, text="Stress overlay spread / liquidity").grid(row=row, column=0, sticky="w", padx=8, pady=5)
        stress_row4 = ttk.Frame(controls)
        stress_row4.grid(row=row, column=1, sticky="w", padx=8, pady=5)
        ttk.Entry(stress_row4, textvariable=self.stress_overlay_spread_multiplier_var, width=8).pack(side="left")
        ttk.Label(stress_row4, text=" / ").pack(side="left")
        ttk.Entry(stress_row4, textvariable=self.stress_overlay_liquidity_multiplier_var, width=8).pack(side="left")

        row += 1
        ttk.Label(controls, text="Hypothesis rubric profile").grid(row=row, column=0, sticky="w", padx=8, pady=5)
        ttk.Combobox(
            controls,
            textvariable=self.hypothesis_rubric_profile_var,
            values=profile_options,
            state="readonly",
            width=24,
        ).grid(row=row, column=1, sticky="w", padx=8, pady=5)

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
        self.wizard_test_plan_var = tk.StringVar(value="walk_forward")
        self.wizard_acceptance_var = tk.StringVar(value="Sharpe >= 0.8 and drawdown >= -0.25")
        self.wizard_run_validation_var = tk.BooleanVar(value=True)
        self.wizard_run_optimization_var = tk.BooleanVar(value=True)
        self.wizard_run_stress_var = tk.BooleanVar(value=True)
        self.wizard_review_notes_var = tk.StringVar(value="")
        self.wizard_promotion_decision_var = tk.StringVar(value="pending")

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
        self._wizard_back_button = ttk.Button(nav, text="Back", command=self._wizard_go_back)
        self._wizard_back_button.pack(side="left")
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
        decision_row = ttk.Frame(frame)
        decision_row.grid(row=3, column=0, sticky="w")
        ttk.Radiobutton(decision_row, text="Promote", value="promote", variable=self.wizard_promotion_decision_var).pack(side="left")
        ttk.Radiobutton(decision_row, text="Reject", value="reject", variable=self.wizard_promotion_decision_var).pack(side="left", padx=(8, 0))

    def _wizard_attach_state_traces(self) -> None:
        tracked_vars = (
            self.wizard_idea_name_var,
            self.wizard_idea_thesis_var,
            self.wizard_idea_owner_var,
            self.wizard_data_universe_var,
            self.wizard_period_start_var,
            self.wizard_period_end_var,
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
            "test_plan": self.wizard_test_plan_var.get().strip(),
            "acceptance_criteria": self.wizard_acceptance_var.get().strip(),
            "run_validation": bool(self.wizard_run_validation_var.get()),
            "run_optimization": bool(self.wizard_run_optimization_var.get()),
            "run_stress": bool(self.wizard_run_stress_var.get()),
            "review_notes": self.wizard_review_notes_var.get().strip(),
            "promotion_decision": self.wizard_promotion_decision_var.get().strip(),
        }

    def _wizard_persist_state(self) -> None:
        self._research_lab_dir.mkdir(parents=True, exist_ok=True)
        self._wizard_state_path.write_text(json.dumps(self._wizard_state_payload(), indent=2, sort_keys=True), encoding="utf-8")

    def _wizard_load_state(self) -> None:
        if not self._wizard_state_path.exists():
            return
        try:
            payload = json.loads(self._wizard_state_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return

        self.wizard_idea_name_var.set(str(payload.get("idea_name", self.wizard_idea_name_var.get())))
        self.wizard_idea_thesis_var.set(str(payload.get("idea_thesis", self.wizard_idea_thesis_var.get())))
        self.wizard_idea_owner_var.set(str(payload.get("idea_owner", self.wizard_idea_owner_var.get())))
        self.wizard_data_universe_var.set(str(payload.get("data_universe", self.wizard_data_universe_var.get())))
        self.wizard_period_start_var.set(str(payload.get("period_start", self.wizard_period_start_var.get())))
        self.wizard_period_end_var.set(str(payload.get("period_end", self.wizard_period_end_var.get())))
        self.wizard_test_plan_var.set(str(payload.get("test_plan", self.wizard_test_plan_var.get())))
        self.wizard_acceptance_var.set(str(payload.get("acceptance_criteria", self.wizard_acceptance_var.get())))
        self.wizard_run_validation_var.set(bool(payload.get("run_validation", self.wizard_run_validation_var.get())))
        self.wizard_run_optimization_var.set(bool(payload.get("run_optimization", self.wizard_run_optimization_var.get())))
        self.wizard_run_stress_var.set(bool(payload.get("run_stress", self.wizard_run_stress_var.get())))
        self.wizard_review_notes_var.set(str(payload.get("review_notes", self.wizard_review_notes_var.get())))
        self.wizard_promotion_decision_var.set(str(payload.get("promotion_decision", self.wizard_promotion_decision_var.get())))
        loaded_step = int(payload.get("current_step", 0))
        self._wizard_step_index = max(0, min(loaded_step, len(self._wizard_steps) - 1))

    def _start_worker(self, target: Callable[[dict[str, Any], ResearchWorkflowConfig], str], label: str) -> None:
        context = self._build_common_context()
        if context is None:
            return
        config = self._build_workflow_config()
        if config is None:
            return

        self._enqueue_task(label=label, target=target, context=context, config=config)

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
            "rubric_profile": self.hypothesis_rubric_profile_var.get().strip() or "intraday_alpha",
            "hypothesis_novelty": float(parse_float(self.hypothesis_novelty_var.get()) or 3.0),
            "hypothesis_plausibility": float(parse_float(self.hypothesis_plausibility_var.get()) or 3.0),
            "hypothesis_implementation_complexity": float(
                parse_float(self.hypothesis_implementation_complexity_var.get()) or 3.0
            ),
            "hypothesis_expected_capacity": float(parse_float(self.hypothesis_expected_capacity_var.get()) or 3.0),
            "hypothesis_robustness": float(parse_float(self.hypothesis_robustness_var.get()) or 3.0),
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

    def _parse_signal_csv(self, raw_text: str, *, valid_options: tuple[str, ...], field_name: str) -> list[str] | None:
        parsed = [item.strip() for item in raw_text.split(",") if item.strip()]
        if not parsed:
            messagebox.showinfo("Invalid input", f"{field_name} must include at least one signal.")
            return None
        invalid = [item for item in parsed if item not in valid_options]
        if invalid:
            messagebox.showinfo("Invalid input", f"Unsupported {field_name.lower()}: {', '.join(invalid)}")
            return None
        return parsed

    def _build_workflow_config(self) -> ResearchWorkflowConfig | None:
        entry_signals = self._parse_signal_csv(
            self.entry_signals_var.get().strip(),
            valid_options=self._signal_options,
            field_name="Entry signals",
        )
        if entry_signals is None:
            return None

        exit_signals = self._parse_signal_csv(
            self.exit_signals_var.get().strip(),
            valid_options=self._exit_signal_options,
            field_name="Exit signals",
        )
        if exit_signals is None:
            return None

        n_trials = int(parse_float(self.optimization_trials_var.get()) or 20)
        if n_trials <= 0:
            messagebox.showinfo("Invalid input", "Optimization trials must be greater than zero.")
            return None

        sampler = self.optimization_sampler_var.get().strip().lower() or "tpe"
        if sampler not in self._sampler_options:
            messagebox.showinfo("Invalid input", "Optimization sampler must be one of: tpe, bayesian, random.")
            return None

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
            entry_signals=entry_signals,
            exit_signals=exit_signals,
            optimization_n_trials=n_trials,
            optimization_sampler=sampler,
            train_fraction=train_fraction,
            validation_fraction=validation_fraction,
            test_fraction=test_fraction,
            step_fraction=step_fraction,
            stress_controls=stress_controls,
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
        }
        return entry_grid, exit_grid, core_grid

    def _run_walk_forward_workflow(self, context: dict[str, Any], config: ResearchWorkflowConfig) -> str:
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
            governance_metadata={"promotion_state": "research", "approval_status": "pending"},
            stress_controls=dict(config.stress_controls),
        )

    def _run_optimization_workflow(self, context: dict[str, Any], config: ResearchWorkflowConfig) -> str:
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
            governance_metadata={"promotion_state": "research", "approval_status": "pending"},
            stress_controls=dict(config.stress_controls),
        )

    def _run_stress_workflow(self, context: dict[str, Any], config: ResearchWorkflowConfig) -> str:
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
            governance_metadata={"promotion_state": "research", "approval_status": "pending"},
            stress_controls=dict(config.stress_controls),
        )

    def _run_hypothesis_pipeline_workflow(self, context: dict[str, Any], config: ResearchWorkflowConfig) -> str:
        self._research_lab_dir.mkdir(parents=True, exist_ok=True)
        hypothesis_id = f"hyp_{uuid.uuid4().hex}"
        idea_record = self._build_idea_record(hypothesis_id=hypothesis_id, context=context)
        idea_record["lineage"] = self._build_pipeline_lineage(hypothesis_id=hypothesis_id)
        scored = self._score_hypothesis(idea_record, context)
        promoted = scored["decision"] == "accept"
        experiment_path = self._write_experiment_skeleton(scored, context) if promoted else None
        self._append_funnel_event(scored)
        funnel = self._compute_funnel_metrics()

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
            },
            "run_references": [
                {
                    "step": "parameter_optimization",
                    "call": "run_strategy_optimization",
                    "status": "todo",
                    "run_id": lineage.get("optimization_run_id"),
                    "parent_id": lineage.get("hypothesis_id"),
                },
                {
                    "step": "walk_forward_validation",
                    "call": "run_walk_forward_backtest",
                    "status": "todo",
                    "run_id": lineage.get("walk_forward_run_id"),
                    "parent_id": lineage.get("optimization_run_id"),
                },
                {
                    "step": "stress_testing",
                    "call": "run_multi_signal_backtest",
                    "status": "todo",
                    "run_id": lineage.get("stress_run_id"),
                    "parent_id": lineage.get("walk_forward_run_id"),
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
            "promotion_history": [
                {
                    "state": str(record.get("promotion_or_rejection", {}).get("state", "pending")),
                    "approval_status": str(record.get("promotion_or_rejection", {}).get("approval_status", "pending")),
                    "reason": str(record.get("promotion_or_rejection", {}).get("reason", "")),
                    "recorded_at": generated_at,
                }
            ],
            "rubric": dict(record.get("rubric", {})),
            "generated_at": generated_at,
            "last_updated": generated_at,
        }

        if output_path.exists():
            try:
                existing = json.loads(output_path.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                existing = {}
            if isinstance(existing, dict):
                for key in ("run_references", "gate_outcomes", "reviewer_comments", "promotion_history"):
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

        output_path.write_text(json.dumps(manifest, indent=2, sort_keys=True), encoding="utf-8")
        return output_path

    def _append_funnel_event(self, record: dict[str, Any]) -> None:
        event_path = self._research_lab_dir / "idea_funnel_events.jsonl"
        event = {
            "hypothesis_id": record["hypothesis_id"],
            "lineage": dict(record.get("lineage", {})),
            "date": date.today().isoformat(),
            "decision": record["decision"],
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

    def _compute_funnel_metrics(self) -> dict[str, float]:
        event_path = self._research_lab_dir / "idea_funnel_events.jsonl"
        if not event_path.exists():
            return {"total_ideas": 0.0, "accepted_ideas": 0.0, "acceptance_rate_pct": 0.0}

        total = 0
        accepted = 0
        for line in event_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line:
                continue
            try:
                parsed = json.loads(line)
            except json.JSONDecodeError:
                continue
            total += 1
            if str(parsed.get("decision", "")).lower() == "accept":
                accepted += 1

        acceptance_rate = (accepted / total * 100.0) if total else 0.0
        return {
            "total_ideas": float(total),
            "accepted_ideas": float(accepted),
            "acceptance_rate_pct": acceptance_rate,
        }

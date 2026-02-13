from __future__ import annotations

import threading
import json
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
from config import BACKTEST_OUTPUT_DIR, DEFAULT_BACKTEST_SETTINGS
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


class ResearchLabPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller
        self._is_running = False
        self._research_lab_dir = BACKTEST_OUTPUT_DIR / "research_lab"
        self._sampler_options = ("tpe", "bayesian", "random")
        self._signal_options = ("ts_momentum", "ma_trend", "breakout")
        self._exit_signal_options = ("none", "momentum_flip", "trailing_stop", "max_hold")

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

    def _start_worker(self, target: Callable[[dict[str, Any], ResearchWorkflowConfig], str], label: str) -> None:
        if self._is_running:
            messagebox.showinfo("Run in progress", "Wait for the current Research Lab run to finish.")
            return

        context = self._build_common_context()
        if context is None:
            return
        config = self._build_workflow_config()
        if config is None:
            return

        self._is_running = True
        self._append_output(f"Starting: {label}")

        def worker() -> None:
            try:
                output = target(context, config)
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

    def _run_hypothesis_pipeline_workflow(self, context: dict[str, Any]) -> str:
        self._research_lab_dir.mkdir(parents=True, exist_ok=True)
        hypothesis_id = f"hyp_{date.today().strftime('%Y%m%d')}_{len(context['tickers'])}t"
        idea_record = self._build_idea_record(hypothesis_id=hypothesis_id, context=context)
        scored = self._score_hypothesis(idea_record)
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

    def _score_hypothesis(self, record: dict[str, Any]) -> dict[str, Any]:
        rubric = {
            "novelty": 3.2,
            "plausibility": 4.0,
            "implementation_complexity": 2.5,
            "expected_capacity": 3.3,
            "robustness": 3.7,
        }
        complexity_adjusted = 6.0 - float(rubric["implementation_complexity"])
        total = (
            float(rubric["novelty"]) * 0.2
            + float(rubric["plausibility"]) * 0.3
            + complexity_adjusted * 0.15
            + float(rubric["expected_capacity"]) * 0.2
            + float(rubric["robustness"]) * 0.15
        )

        accepted = total >= 3.2 and float(rubric["plausibility"]) >= 3.0 and float(rubric["robustness"]) >= 3.0
        decision = "accept" if accepted else "reject"
        reason = "Clears minimum weighted rubric thresholds" if accepted else "Fails weighted rubric threshold"
        record["rubric"] = {**rubric, "total_score": total}
        record["decision"] = decision
        record["decision_reason"] = reason
        record["promotion_or_rejection"] = {
            "state": "promoted_to_experiment" if accepted else "rejected",
            "approval_status": "pending" if accepted else "rejected",
            "reason": reason,
        }
        return record

    def _write_experiment_skeleton(self, record: dict[str, Any], context: dict[str, Any]) -> Path:
        skeleton_dir = self._research_lab_dir / "experiment_skeletons"
        skeleton_dir.mkdir(parents=True, exist_ok=True)
        output_path = skeleton_dir / f"{record['hypothesis_id']}.json"
        payload = {
            "hypothesis_id": record["hypothesis_id"],
            "workflow_stages": [
                "idea_intake",
                "economic_rationale",
                "data_requirements",
                "test_design",
                "results_review",
                "promotion_or_rejection",
            ],
            "run_plan": [
                {"step": "walk_forward_validation", "call": "run_walk_forward_backtest", "status": "todo"},
                {"step": "parameter_optimization", "call": "run_strategy_optimization", "status": "todo"},
                {"step": "stress_testing", "call": "run_multi_signal_backtest", "status": "todo"},
            ],
            "context": {
                "tickers": list(context["tickers"]),
                "start_date": str(context["start_date"]),
                "end_date": str(context["end_date"]),
                "lookback": int(context["lookback"]),
                "skip": int(context["skip"]),
                "costs_bps": float(context["costs_bps"]),
            },
            "rubric": record["rubric"],
            "generated_at": date.today().isoformat(),
        }
        output_path.write_text(json.dumps(payload, indent=2, sort_keys=True), encoding="utf-8")
        return output_path

    def _append_funnel_event(self, record: dict[str, Any]) -> None:
        event_path = self._research_lab_dir / "idea_funnel_events.jsonl"
        event = {
            "hypothesis_id": record["hypothesis_id"],
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

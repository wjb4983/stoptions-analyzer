from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk
from typing import Any, Callable


def default_calibration_spec() -> dict[str, Any]:
    return {
        "schema_version": 1,
        "method": "sabr",
        "objective": "rmse",
        "max_iterations": 250,
        "tolerance": 0.0001,
    }


def normalize_calibration_spec(payload: dict[str, Any] | None) -> dict[str, Any]:
    base = default_calibration_spec()
    if not isinstance(payload, dict):
        return dict(base)
    normalized = dict(base)
    method = str(payload.get("method", base["method"])).strip()
    objective = str(payload.get("objective", base["objective"])).strip()
    normalized["method"] = method or base["method"]
    normalized["objective"] = objective or base["objective"]
    try:
        normalized["max_iterations"] = max(1, int(payload.get("max_iterations", base["max_iterations"])))
    except (TypeError, ValueError):
        normalized["max_iterations"] = base["max_iterations"]
    try:
        normalized["tolerance"] = max(1e-8, float(payload.get("tolerance", base["tolerance"])))
    except (TypeError, ValueError):
        normalized["tolerance"] = base["tolerance"]
    return normalized


class CalibrationSpecDesignerPage(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Misc,
        *,
        initial_spec: dict[str, Any] | None,
        on_save: Callable[[dict[str, Any]], None],
    ) -> None:
        super().__init__(parent)
        self.title("Calibration Spec Designer")
        self.geometry("520x260")
        self.transient(parent)

        self._on_save = on_save
        spec = normalize_calibration_spec(initial_spec)

        self.method_var = tk.StringVar(value=str(spec["method"]))
        self.objective_var = tk.StringVar(value=str(spec["objective"]))
        self.max_iterations_var = tk.StringVar(value=str(spec["max_iterations"]))
        self.tolerance_var = tk.StringVar(value=str(spec["tolerance"]))

        frame = ttk.Frame(self, padding=10)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Calibration method").grid(row=0, column=0, sticky="w")
        ttk.Combobox(
            frame,
            textvariable=self.method_var,
            state="readonly",
            values=["sabr", "heston", "local_vol", "term_structure"],
            width=24,
        ).grid(row=0, column=1, sticky="w", padx=6)

        ttk.Label(frame, text="Objective").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Combobox(
            frame,
            textvariable=self.objective_var,
            state="readonly",
            values=["rmse", "mae", "weighted_rmse"],
            width=24,
        ).grid(row=1, column=1, sticky="w", padx=6, pady=(8, 0))

        ttk.Label(frame, text="Max iterations").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(frame, textvariable=self.max_iterations_var, width=26).grid(row=2, column=1, sticky="w", padx=6, pady=(8, 0))

        ttk.Label(frame, text="Tolerance").grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(frame, textvariable=self.tolerance_var, width=26).grid(row=3, column=1, sticky="w", padx=6, pady=(8, 0))

        actions = ttk.Frame(frame)
        actions.grid(row=4, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(actions, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(actions, text="Save calibration spec", command=self._save).pack(side="right", padx=(0, 6))

    def _save(self) -> None:
        try:
            payload = normalize_calibration_spec(
                {
                    "method": self.method_var.get(),
                    "objective": self.objective_var.get(),
                    "max_iterations": int(self.max_iterations_var.get()),
                    "tolerance": float(self.tolerance_var.get()),
                }
            )
        except ValueError as exc:
            messagebox.showerror("Invalid calibration spec", str(exc))
            return

        self._on_save(payload)
        self.destroy()

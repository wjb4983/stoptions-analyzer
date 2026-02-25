from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk
from typing import Any, Callable


def default_event_process_spec() -> dict[str, Any]:
    return {
        "schema_version": 1,
        "process_family": "hawkes",
        "baseline_intensity": 0.05,
        "excitation_alpha": 0.35,
        "decay_half_life_minutes": 45,
    }


def normalize_event_process_spec(payload: dict[str, Any] | None) -> dict[str, Any]:
    base = default_event_process_spec()
    if not isinstance(payload, dict):
        return dict(base)
    normalized = dict(base)
    family = str(payload.get("process_family", base["process_family"])).strip()
    normalized["process_family"] = family or base["process_family"]
    for key, floor in (
        ("baseline_intensity", 0.0),
        ("excitation_alpha", 0.0),
        ("decay_half_life_minutes", 1.0),
    ):
        try:
            normalized[key] = max(floor, float(payload.get(key, base[key])))
        except (TypeError, ValueError):
            normalized[key] = base[key]
    return normalized


class EventProcessDesignerPage(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Misc,
        *,
        initial_spec: dict[str, Any] | None,
        on_save: Callable[[dict[str, Any]], None],
    ) -> None:
        super().__init__(parent)
        self.title("Event Process Spec Designer")
        self.geometry("560x280")
        self.transient(parent)

        self._on_save = on_save
        spec = normalize_event_process_spec(initial_spec)

        self.process_family_var = tk.StringVar(value=str(spec["process_family"]))
        self.baseline_intensity_var = tk.StringVar(value=str(spec["baseline_intensity"]))
        self.excitation_alpha_var = tk.StringVar(value=str(spec["excitation_alpha"]))
        self.decay_half_life_var = tk.StringVar(value=str(spec["decay_half_life_minutes"]))

        frame = ttk.Frame(self, padding=10)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Process family").grid(row=0, column=0, sticky="w")
        ttk.Combobox(
            frame,
            textvariable=self.process_family_var,
            state="readonly",
            values=["hawkes", "poisson", "self_correcting"],
            width=24,
        ).grid(row=0, column=1, sticky="w", padx=6)

        ttk.Label(frame, text="Baseline intensity").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(frame, textvariable=self.baseline_intensity_var, width=26).grid(row=1, column=1, sticky="w", padx=6, pady=(8, 0))

        ttk.Label(frame, text="Excitation alpha").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(frame, textvariable=self.excitation_alpha_var, width=26).grid(row=2, column=1, sticky="w", padx=6, pady=(8, 0))

        ttk.Label(frame, text="Decay half-life (minutes)").grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(frame, textvariable=self.decay_half_life_var, width=26).grid(row=3, column=1, sticky="w", padx=6, pady=(8, 0))

        actions = ttk.Frame(frame)
        actions.grid(row=4, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(actions, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(actions, text="Save event process spec", command=self._save).pack(side="right", padx=(0, 6))

    def _save(self) -> None:
        try:
            payload = normalize_event_process_spec(
                {
                    "process_family": self.process_family_var.get(),
                    "baseline_intensity": float(self.baseline_intensity_var.get()),
                    "excitation_alpha": float(self.excitation_alpha_var.get()),
                    "decay_half_life_minutes": float(self.decay_half_life_var.get()),
                }
            )
        except ValueError as exc:
            messagebox.showerror("Invalid event process spec", str(exc))
            return

        self._on_save(payload)
        self.destroy()

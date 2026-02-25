from __future__ import annotations

from tkinter import ttk
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from main import StoptionsApp


class CreateRegimePage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller

        # Ensure regime state keys are initialized and reflected in the UI.
        if not self.controller.state.regime_definitions:
            self.controller.state.regime_definitions = {
                "baseline": {
                    "label": "Baseline",
                    "description": "Default regime profile",
                }
            }
        if self.controller.state.active_regime_id is None:
            self.controller.state.active_regime_id = next(
                iter(self.controller.state.regime_definitions), None
            )

        regime_count = len(self.controller.state.regime_definitions)
        run_count = len(self.controller.state.regime_training_runs)

        ttk.Label(self, text="Create Regime", font=("Arial", 18, "bold")).pack(pady=12)
        ttk.Label(
            self,
            text=(
                "Build and manage regime definitions from this workspace. "
                "This page is ready for future Create Regime workflows."
            ),
            wraplength=700,
            justify="center",
        ).pack(pady=(0, 16), padx=24)

        ttk.Label(
            self,
            text=f"Definitions: {regime_count} • Training runs: {run_count}",
        ).pack(pady=(0, 12))

        ttk.Button(
            self,
            text="Back to Main Menu",
            command=lambda: controller.show_frame("MainMenu"),
            width=24,
        ).pack(pady=8)

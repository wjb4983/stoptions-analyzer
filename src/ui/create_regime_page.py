from __future__ import annotations

from tkinter import ttk


class CreateRegimePage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller

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

        ttk.Button(
            self,
            text="Back to Main Menu",
            command=lambda: controller.show_frame("MainMenu"),
            width=24,
        ).pack(pady=8)

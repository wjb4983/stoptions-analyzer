from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from config import API_KEY_PATH
from ui.helpers import save_api_key


class MainMenu(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller

        title = ttk.Label(self, text="Stoptions Analyzer", font=("Arial", 24, "bold"))
        title.pack(pady=20)

        description = ttk.Label(
            self,
            text="Manage tickers, select a stock, and explore option strategy analysis.",
            wraplength=600,
            justify="center",
        )
        description.pack(pady=10)

        api_frame = ttk.LabelFrame(self, text="Massive API Key")
        api_frame.pack(pady=15, padx=40, fill="x")
        api_frame.columnconfigure(1, weight=1)

        ttk.Label(api_frame, text="API Key").grid(row=0, column=0, padx=10, pady=8, sticky="w")
        self.api_key_var = tk.StringVar(value=self.controller.api_key)
        self.api_key_entry = ttk.Entry(api_frame, textvariable=self.api_key_var, show="*")
        self.api_key_entry.grid(row=0, column=1, padx=10, pady=8, sticky="ew")
        ttk.Button(api_frame, text="Save Key", command=self.save_api_key).grid(
            row=0, column=2, padx=10, pady=8
        )

        button_frame = ttk.Frame(self)
        button_frame.pack(pady=40)

        ttk.Button(
            button_frame,
            text="Enter Stock Tickers",
            command=lambda: controller.show_frame("TickerEntryPage"),
            width=30,
        ).grid(row=0, column=0, pady=10)

        ttk.Button(
            button_frame,
            text="Select Stock",
            command=lambda: controller.show_frame("TickerSelectPage"),
            width=30,
        ).grid(row=1, column=0, pady=10)

        ttk.Button(
            button_frame,
            text="Analysis",
            command=lambda: controller.show_frame("AnalysisPage"),
            width=30,
        ).grid(row=2, column=0, pady=10)

        ttk.Button(
            button_frame,
            text="General Analysis",
            command=lambda: controller.show_frame("GeneralAnalysisPage"),
            width=30,
        ).grid(row=3, column=0, pady=10)

        ttk.Button(
            button_frame,
            text="Backtesting",
            command=lambda: controller.show_frame("BacktestingPage"),
            width=30,
        ).grid(row=4, column=0, pady=10)

        ttk.Button(
            button_frame,
            text="Research Lab",
            command=lambda: controller.show_frame("ResearchLabPage"),
            width=30,
        ).grid(row=5, column=0, pady=10)

    def refresh(self) -> None:
        self.api_key_var.set(self.controller.api_key)

    def save_api_key(self) -> None:
        key = self.api_key_var.get().strip()
        if not key:
            messagebox.showinfo("Missing key", "Enter a Massive API key first.")
            return
        save_api_key(key)
        self.controller.api_key = key
        messagebox.showinfo(
            "Saved", f"API key saved to {API_KEY_PATH} (not tracked in git)."
        )

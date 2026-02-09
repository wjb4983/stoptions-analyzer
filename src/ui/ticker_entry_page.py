from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk


class TickerEntryPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller

        ttk.Label(self, text="Enter Stock Tickers", font=("Arial", 18, "bold")).pack(pady=10)

        instructions = ttk.Label(
            self,
            text="Enter one ticker per line. Click Save to store them locally.",
        )
        instructions.pack(pady=5)

        self.text_box = tk.Text(self, height=18, width=40)
        self.text_box.pack(pady=10)

        button_row = ttk.Frame(self)
        button_row.pack(pady=10)

        ttk.Button(button_row, text="Save", command=self.save_tickers).grid(row=0, column=0, padx=10)
        ttk.Button(
            button_row,
            text="Back to Main Menu",
            command=lambda: controller.show_frame("MainMenu"),
        ).grid(row=0, column=1, padx=10)

    def refresh(self) -> None:
        self.text_box.delete("1.0", tk.END)
        self.text_box.insert("1.0", "\n".join(self.controller.state.tickers))

    def save_tickers(self) -> None:
        raw = self.text_box.get("1.0", tk.END)
        tickers = [line.strip().upper() for line in raw.splitlines() if line.strip()]
        if not tickers:
            messagebox.showinfo("No tickers", "Please enter at least one ticker.")
            return
        self.controller.state.tickers = tickers
        if self.controller.state.selected_ticker not in tickers:
            self.controller.state.selected_ticker = tickers[0]
        self.controller.persist_state()
        messagebox.showinfo("Saved", "Tickers saved successfully.")

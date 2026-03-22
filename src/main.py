from pathlib import Path
import sys

# Ensure both repository root and src directory are importable when running this file directly.
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
_SRC_ROOT = _PROJECT_ROOT / "src"
for _path in (str(_PROJECT_ROOT), str(_SRC_ROOT)):
    if _path not in sys.path:
        sys.path.insert(0, _path)

import os
import tkinter as tk
from tkinter import ttk

from execution import build_execution_backend
from state import AppState
from ui import (
    AnalysisPage,
    BacktestingPage,
    CallPutAnalysisPage,
    CreateRegimePage,
    GeneralAnalysisPage,
    IntradayReplayPage,
    MainMenu,
    ResearchLabPage,
    SpreadAnalysisPage,
    TickerEntryPage,
    TickerSelectPage,
)
from ui.helpers import load_api_key

class StoptionsApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Stoptions Analyzer")
        self.geometry("1200x800")
        self._maximize_window()
        self.state = AppState.load()
        self.api_key = load_api_key()
        backend_mode = os.getenv("STOPTIONS_EXECUTION_BACKEND", "local")
        self.execution_backend = build_execution_backend(mode=backend_mode)

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        self.frames: dict[str, ttk.Frame] = {}
        for frame_cls in (
            MainMenu,
            ResearchLabPage,
            BacktestingPage,
            TickerEntryPage,
            TickerSelectPage,
            AnalysisPage,
            GeneralAnalysisPage,
            IntradayReplayPage,
            CallPutAnalysisPage,
            SpreadAnalysisPage,
            CreateRegimePage,
        ):
            frame = frame_cls(container, self)
            self.frames[frame_cls.__name__] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("MainMenu")

    def show_frame(self, name: str) -> None:
        frame = self.frames[name]
        if hasattr(frame, "refresh"):
            frame.refresh()
        frame.tkraise()

    def persist_state(self) -> None:
        self.state.save()

    def _maximize_window(self) -> None:
        self.update_idletasks()
        try:
            self.state("zoomed")
        except tk.TclError:
            self.attributes("-fullscreen", True)


if __name__ == "__main__":
    app = StoptionsApp()
    app.mainloop()

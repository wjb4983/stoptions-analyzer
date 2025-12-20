import json
from dataclasses import dataclass, field
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

STATE_PATH = Path(__file__).resolve().parent / "app_state.txt"


@dataclass
class AppState:
    tickers: list[str] = field(default_factory=list)
    selected_ticker: str | None = None
    analysis_type: str = "Stock Analysis"
    knob_delta: int = 50
    knob_risk: int = 50
    knob_prob: int = 50

    def save(self) -> None:
        payload = {
            "tickers": self.tickers,
            "selected_ticker": self.selected_ticker,
            "analysis_type": self.analysis_type,
            "knob_delta": self.knob_delta,
            "knob_risk": self.knob_risk,
            "knob_prob": self.knob_prob,
        }
        STATE_PATH.write_text(json.dumps(payload, indent=2))

    @classmethod
    def load(cls) -> "AppState":
        if not STATE_PATH.exists():
            return cls()
        try:
            payload = json.loads(STATE_PATH.read_text())
        except json.JSONDecodeError:
            return cls()
        return cls(
            tickers=payload.get("tickers", []),
            selected_ticker=payload.get("selected_ticker"),
            analysis_type=payload.get("analysis_type", "Stock Analysis"),
            knob_delta=payload.get("knob_delta", 50),
            knob_risk=payload.get("knob_risk", 50),
            knob_prob=payload.get("knob_prob", 50),
        )


class StoptionsApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Stoptions Analyzer")
        self.geometry("900x600")
        self.state = AppState.load()

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        self.frames: dict[str, ttk.Frame] = {}
        for frame_cls in (MainMenu, TickerEntryPage, TickerSelectPage, AnalysisPage):
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


class TickerSelectPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller

        ttk.Label(self, text="Select a Stock", font=("Arial", 18, "bold")).pack(pady=10)

        list_frame = ttk.Frame(self)
        list_frame.pack(pady=10, fill="both", expand=True)

        self.ticker_list = tk.Listbox(list_frame, height=18)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.ticker_list.yview)
        self.ticker_list.configure(yscrollcommand=scrollbar.set)

        self.ticker_list.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        button_row = ttk.Frame(self)
        button_row.pack(pady=10)

        ttk.Button(button_row, text="Use Selected", command=self.use_selected).grid(
            row=0, column=0, padx=10
        )
        ttk.Button(
            button_row,
            text="Back to Main Menu",
            command=lambda: controller.show_frame("MainMenu"),
        ).grid(row=0, column=1, padx=10)

    def refresh(self) -> None:
        self.ticker_list.delete(0, tk.END)
        for ticker in self.controller.state.tickers:
            self.ticker_list.insert(tk.END, ticker)
        if self.controller.state.selected_ticker in self.controller.state.tickers:
            index = self.controller.state.tickers.index(self.controller.state.selected_ticker)
            self.ticker_list.selection_set(index)
            self.ticker_list.see(index)

    def use_selected(self) -> None:
        selection = self.ticker_list.curselection()
        if not selection:
            messagebox.showinfo("Select a ticker", "Please select a ticker from the list.")
            return
        ticker = self.ticker_list.get(selection[0])
        self.controller.state.selected_ticker = ticker
        self.controller.persist_state()
        self.controller.show_frame("AnalysisPage")


class AnalysisPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller

        ttk.Label(self, text="Analysis", font=("Arial", 18, "bold")).pack(pady=10)

        self.selected_label = ttk.Label(self, text="Selected Ticker: None")
        self.selected_label.pack(pady=5)

        selector_frame = ttk.Frame(self)
        selector_frame.pack(pady=10)

        ttk.Label(selector_frame, text="Analysis Type:").grid(row=0, column=0, padx=5)
        self.analysis_var = tk.StringVar(value=self.controller.state.analysis_type)
        self.analysis_dropdown = ttk.Combobox(
            selector_frame,
            textvariable=self.analysis_var,
            values=[
                "Stock Analysis",
                "Naked Call",
                "Naked Put",
                "Vertical Spread",
                "Calendar Spread",
            ],
            state="readonly",
            width=25,
        )
        self.analysis_dropdown.grid(row=0, column=1, padx=5)
        self.analysis_dropdown.bind("<<ComboboxSelected>>", self.on_analysis_change)

        knobs_frame = ttk.LabelFrame(self, text="Strategy Knobs")
        knobs_frame.pack(pady=10, fill="x", padx=40)

        self.delta_var = tk.IntVar(value=self.controller.state.knob_delta)
        self.risk_var = tk.IntVar(value=self.controller.state.knob_risk)
        self.prob_var = tk.IntVar(value=self.controller.state.knob_prob)

        self._build_slider(knobs_frame, "Option Delta", self.delta_var, 0)
        self._build_slider(knobs_frame, "Risk", self.risk_var, 1)
        self._build_slider(knobs_frame, "Probability of Profit", self.prob_var, 2)

        metrics_frame = ttk.LabelFrame(self, text="Metrics")
        metrics_frame.pack(pady=10, fill="x", padx=40)

        self.metrics_text = tk.Text(metrics_frame, height=6)
        self.metrics_text.pack(fill="x", padx=10, pady=5)
        self.metrics_text.insert(
            "1.0",
            "Option Delta: --\nRisk: --\nProbability of Profit: --\n",
        )
        self.metrics_text.configure(state="disabled")

        chart_frame = ttk.LabelFrame(self, text="Chart Placeholder")
        chart_frame.pack(pady=10, fill="both", expand=True, padx=40)

        self.chart_canvas = tk.Canvas(chart_frame, height=180, bg="#f0f0f0")
        self.chart_canvas.pack(fill="both", expand=True, padx=10, pady=10)
        self.chart_canvas.create_text(
            200,
            90,
            text="Charts will render here in a future update.",
            fill="#666",
        )

        button_row = ttk.Frame(self)
        button_row.pack(pady=10)

        ttk.Button(button_row, text="Save Analysis", command=self.save_analysis).grid(
            row=0, column=0, padx=10
        )
        ttk.Button(
            button_row,
            text="Back to Main Menu",
            command=lambda: controller.show_frame("MainMenu"),
        ).grid(row=0, column=1, padx=10)

    def _build_slider(self, parent: ttk.Frame, label: str, var: tk.IntVar, row: int) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, padx=10, pady=5, sticky="w")
        slider = ttk.Scale(parent, from_=0, to=100, orient="horizontal", variable=var)
        slider.grid(row=row, column=1, padx=10, pady=5, sticky="ew")
        value_label = ttk.Label(parent, textvariable=var)
        value_label.grid(row=row, column=2, padx=10, pady=5)
        parent.columnconfigure(1, weight=1)

    def refresh(self) -> None:
        ticker = self.controller.state.selected_ticker or "None"
        self.selected_label.config(text=f"Selected Ticker: {ticker}")
        self.analysis_var.set(self.controller.state.analysis_type)
        self.delta_var.set(self.controller.state.knob_delta)
        self.risk_var.set(self.controller.state.knob_risk)
        self.prob_var.set(self.controller.state.knob_prob)

    def on_analysis_change(self, _event: object) -> None:
        self.controller.state.analysis_type = self.analysis_var.get()
        self.controller.persist_state()

    def save_analysis(self) -> None:
        self.controller.state.analysis_type = self.analysis_var.get()
        self.controller.state.knob_delta = int(self.delta_var.get())
        self.controller.state.knob_risk = int(self.risk_var.get())
        self.controller.state.knob_prob = int(self.prob_var.get())
        self.controller.persist_state()
        messagebox.showinfo("Saved", "Analysis settings saved locally.")


if __name__ == "__main__":
    app = StoptionsApp()
    app.mainloop()

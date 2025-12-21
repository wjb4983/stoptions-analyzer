import json
import os
from dataclasses import dataclass, field
from datetime import date, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import urlopen

STATE_PATH = Path(__file__).resolve().parent / "app_state.txt"
CONFIG_DIR = Path.home() / ".stoptions_analyzer"
API_KEY_PATH = CONFIG_DIR / "api_key.txt"
API_BASE_URL = os.getenv("MASSIVE_BASE_URL", "https://api.polygon.io")
HORIZON_CONFIGS = [
    ("Day", 1, 10, "10m"),
    ("3 Day", 3, 30, "30m"),
    ("Week", 7, 60, "1h"),
    ("Month", 30, 120, "2h"),
    ("3M", 90, 360, "6h"),
    ("6M", 180, 720, "12h"),
    ("12M", 365, 1440, "1d"),
    ("3Y", 1095, 4320, "3d"),
    ("5Y", 1825, 7200, "5d"),
    ("10Y", 3650, 10080, "7d"),
]


def load_api_key() -> str:
    env_key = os.getenv("MASSIVE_API_KEY", "").strip()
    if env_key:
        return env_key
    if API_KEY_PATH.exists():
        return API_KEY_PATH.read_text().strip()
    return ""


def save_api_key(key: str) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    API_KEY_PATH.write_text(key.strip())
    try:
        API_KEY_PATH.chmod(0o600)
    except OSError:
        pass


class MassiveApiClient:
    def __init__(self, api_key: str, base_url: str = API_BASE_URL) -> None:
        self.api_key = api_key
        self.base_url = base_url.rstrip("/")

    def _request(self, path: str, params: dict[str, str]) -> dict:
        params = {**params, "apiKey": self.api_key}
        url = f"{self.base_url}{path}?{urlencode(params)}"
        with urlopen(url, timeout=10) as response:
            payload = response.read().decode("utf-8")
        return json.loads(payload)

    def fetch_previous_close(self, ticker: str) -> dict:
        data = self._request(f"/v2/aggs/ticker/{ticker}/prev", {"adjusted": "true"})
        result = (data.get("results") or [{}])[0]
        return {
            "close": result.get("c"),
            "open": result.get("o"),
            "high": result.get("h"),
            "low": result.get("l"),
            "volume": result.get("v"),
        }

    def fetch_option_contracts(self, ticker: str, limit: int = 5) -> list[dict]:
        data = self._request(
            "/v3/reference/options/contracts",
            {"underlying_ticker": ticker, "limit": str(limit)},
        )
        return data.get("results", [])

    def fetch_aggregates(self, ticker: str, days_back: int, minutes_per_bar: int) -> list[dict]:
        end_date = date.today()
        start_date = end_date - timedelta(days=days_back)
        data = self._request(
            f"/v2/aggs/ticker/{ticker}/range/{minutes_per_bar}/minute/{start_date}/{end_date}",
            {"adjusted": "true", "sort": "asc", "limit": "5000"},
        )
        return data.get("results", [])


@dataclass
class AppState:
    tickers: list[str] = field(default_factory=list)
    selected_ticker: str | None = None
    analysis_mode: str = "Stock Analysis"
    option_strategy: str = "Naked Call"
    knob_delta: int = 50
    knob_risk: int = 50
    knob_prob: int = 50

    def save(self) -> None:
        payload = {
            "tickers": self.tickers,
            "selected_ticker": self.selected_ticker,
            "analysis_mode": self.analysis_mode,
            "option_strategy": self.option_strategy,
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
            analysis_mode=payload.get("analysis_mode", payload.get("analysis_type", "Stock Analysis")),
            option_strategy=payload.get("option_strategy", "Naked Call"),
            knob_delta=payload.get("knob_delta", 50),
            knob_risk=payload.get("knob_risk", 50),
            knob_prob=payload.get("knob_prob", 50),
        )


class StoptionsApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Stoptions Analyzer")
        self.geometry("1200x800")
        self._maximize_window()
        self.state = AppState.load()
        self.api_key = load_api_key()

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

    def _maximize_window(self) -> None:
        self.update_idletasks()
        try:
            self.state("zoomed")
        except tk.TclError:
            self.attributes("-fullscreen", True)


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
        self.api_client: MassiveApiClient | None = None
        self.option_contract: dict | None = None
        self.scroll_canvas = tk.Canvas(self, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.scroll_canvas.yview)
        self.scroll_canvas.configure(yscrollcommand=self.scrollbar.set)
        self.scroll_canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.content_frame = ttk.Frame(self.scroll_canvas)
        self.scroll_window = self.scroll_canvas.create_window(
            (0, 0), window=self.content_frame, anchor="nw"
        )
        self.content_frame.bind("<Configure>", self._on_content_configure)
        self.scroll_canvas.bind("<Configure>", self._on_canvas_configure)
        self.scroll_canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        ttk.Label(self.content_frame, text="Analysis", font=("Arial", 20, "bold")).pack(
            pady=10
        )

        self.selected_label = ttk.Label(self.content_frame, text="Selected Ticker: None")
        self.selected_label.pack(pady=5)

        integration_note = ttk.Label(
            self.content_frame,
            text="Massive integration uses the API key saved in the main menu (or env var).",
            foreground="#555",
        )
        integration_note.pack(pady=5)

        ttk.Button(self.content_frame, text="Load Data", command=self.load_market_data).pack(
            pady=10
        )

        stock_frame = ttk.LabelFrame(self.content_frame, text="Stock Analysis")
        stock_frame.pack(pady=10, fill="both", expand=True, padx=40)

        chart_header = ttk.Label(
            stock_frame,
            text="Current (or previous trading day) chart",
            font=("Arial", 12, "bold"),
        )
        chart_header.pack(pady=(10, 5))

        chart_frame = ttk.Frame(stock_frame)
        chart_frame.pack(pady=5, fill="both", expand=True)

        self.chart_canvas = tk.Canvas(chart_frame, height=220, bg="#f0f0f0")
        self.chart_canvas.pack(fill="both", expand=True, padx=10, pady=10)
        self.chart_canvas.create_text(
            220,
            110,
            text="Daily chart preview will render here.",
            fill="#666",
        )

        slider_frame = ttk.Frame(stock_frame)
        slider_frame.pack(fill="x", padx=20, pady=(5, 10))

        ttk.Label(slider_frame, text="Time Horizon").grid(
            row=0, column=0, columnspan=len(HORIZON_CONFIGS), sticky="w"
        )
        self.horizon_var = tk.IntVar(value=0)
        self.horizon_slider = tk.Scale(
            slider_frame,
            from_=0,
            to=len(HORIZON_CONFIGS) - 1,
            orient="horizontal",
            variable=self.horizon_var,
            resolution=1,
            showvalue=False,
            command=self._snap_horizon,
            length=600,
        )
        self.horizon_slider.grid(
            row=1, column=0, columnspan=len(HORIZON_CONFIGS), sticky="ew", pady=5
        )
        for index in range(len(HORIZON_CONFIGS)):
            slider_frame.columnconfigure(index, weight=1)

        labels_frame = ttk.Frame(slider_frame)
        labels_frame.grid(row=2, column=0, columnspan=len(HORIZON_CONFIGS), sticky="ew")
        for index, (label, _days, _minutes, cadence_label) in enumerate(HORIZON_CONFIGS):
            ttk.Label(labels_frame, text=f"{label}\n({cadence_label})").grid(
                row=0, column=index, padx=4
            )
            labels_frame.columnconfigure(index, weight=1)

        self.stock_info_frame = ttk.LabelFrame(stock_frame, text="Stock Snapshot")
        self.stock_info_frame.pack(padx=20, pady=(5, 15), fill="x")
        self.stock_values: dict[str, ttk.Label] = {}
        self._build_info_grid(
            self.stock_info_frame,
            [
                ("Price", "price"),
                ("Previous Close", "prev_close"),
                ("Open", "open"),
                ("High", "high"),
                ("Low", "low"),
                ("Volume", "volume"),
                ("Market Cap", "market_cap"),
                ("52 Week Range", "range_52w"),
            ],
            self.stock_values,
            columns=2,
        )

        self.option_info_frame = ttk.LabelFrame(stock_frame, text="Option Snapshot")
        self.option_info_frame.pack(padx=20, pady=(5, 15), fill="x")
        self.option_values: dict[str, ttk.Label] = {}
        self._build_info_grid(
            self.option_info_frame,
            [
                ("Contract", "contract"),
                ("Expiration", "expiration"),
                ("Type", "type"),
                ("Strike", "strike"),
            ],
            self.option_values,
        )

        self.options_frame = ttk.LabelFrame(stock_frame, text="Option Contracts (Sample)")
        self.options_frame.pack(padx=20, pady=(5, 15), fill="x")

        self.options_text = tk.Text(self.options_frame, height=6)
        self.options_text.pack(fill="x", padx=10, pady=8)
        self.options_text.insert("1.0", "No option data loaded yet.\n")
        self.options_text.configure(state="disabled")

        selector_frame = ttk.Frame(self.content_frame)
        selector_frame.pack(pady=5)
        selector_frame.columnconfigure(1, weight=1)

        ttk.Label(selector_frame, text="Analysis Mode:").grid(row=0, column=0, padx=5, sticky="w")
        self.analysis_mode_var = tk.StringVar(value=self.controller.state.analysis_mode)
        self.analysis_mode_dropdown = ttk.Combobox(
            selector_frame,
            textvariable=self.analysis_mode_var,
            values=["Stock Analysis", "Option Analysis"],
            state="readonly",
            width=20,
        )
        self.analysis_mode_dropdown.grid(row=0, column=1, padx=5, sticky="w")
        self.analysis_mode_dropdown.bind("<<ComboboxSelected>>", self.on_analysis_mode_change)

        self.strategy_frame = ttk.Frame(self.content_frame)
        self.strategy_frame.pack(pady=5)

        ttk.Label(self.strategy_frame, text="Option Strategy:").grid(row=0, column=0, padx=5)
        self.strategy_var = tk.StringVar(value=self.controller.state.option_strategy)
        self.strategy_dropdown = ttk.Combobox(
            self.strategy_frame,
            textvariable=self.strategy_var,
            values=[
                "Naked Call",
                "Naked Put",
                "Vertical Spread",
                "Calendar Spread",
            ],
            state="readonly",
            width=25,
        )
        self.strategy_dropdown.grid(row=0, column=1, padx=5)
        self.strategy_dropdown.bind("<<ComboboxSelected>>", self.on_strategy_change)

        self.knobs_frame = ttk.LabelFrame(self.content_frame, text="Option Strategy Knobs")
        self.knobs_frame.pack(pady=10, fill="x", padx=40)

        self.delta_var = tk.IntVar(value=self.controller.state.knob_delta)
        self.risk_var = tk.IntVar(value=self.controller.state.knob_risk)
        self.prob_var = tk.IntVar(value=self.controller.state.knob_prob)

        self._build_slider(self.knobs_frame, "Option Delta", self.delta_var, 0)
        self._build_slider(self.knobs_frame, "Risk", self.risk_var, 1)
        self._build_slider(self.knobs_frame, "Probability of Profit", self.prob_var, 2)

        button_row = ttk.Frame(self.content_frame)
        button_row.pack(pady=10)

        ttk.Button(button_row, text="Save Analysis", command=self.save_analysis).grid(
            row=0, column=0, padx=10
        )
        ttk.Button(
            button_row,
            text="Select Stock",
            command=lambda: controller.show_frame("TickerSelectPage"),
        ).grid(row=0, column=1, padx=10)
        ttk.Button(
            button_row,
            text="Back to Main Menu",
            command=lambda: controller.show_frame("MainMenu"),
        ).grid(row=0, column=2, padx=10)

    def _build_slider(self, parent: ttk.Frame, label: str, var: tk.IntVar, row: int) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, padx=10, pady=5, sticky="w")
        slider = ttk.Scale(parent, from_=0, to=100, orient="horizontal", variable=var)
        slider.grid(row=row, column=1, padx=10, pady=5, sticky="ew")
        value_label = ttk.Label(parent, textvariable=var)
        value_label.grid(row=row, column=2, padx=10, pady=5)
        parent.columnconfigure(1, weight=1)

    def _snap_horizon(self, value: str) -> None:
        snapped = int(round(float(value)))
        self.horizon_var.set(snapped)
        self.horizon_slider.set(snapped)

    def _build_info_grid(
        self,
        parent: ttk.Frame,
        rows: list[tuple[str, str]],
        target: dict[str, ttk.Label],
        columns: int = 1,
    ) -> None:
        for item_index, (label, key) in enumerate(rows):
            row_index = item_index // columns
            column_index = (item_index % columns) * 2
            ttk.Label(parent, text=label).grid(
                row=row_index, column=column_index, padx=10, pady=4, sticky="w"
            )
            value_label = ttk.Label(parent, text="--", foreground="#b00020")
            value_label.grid(
                row=row_index, column=column_index + 1, padx=10, pady=4, sticky="w"
            )
            target[key] = value_label
        for index in range(columns * 2):
            parent.columnconfigure(index, weight=1)

    def _set_value(self, label: ttk.Label, value: str | int | float | None) -> None:
        if value in (None, "", "--"):
            label.config(text="--", foreground="#b00020")
        else:
            label.config(text=str(value), foreground="#0a7a2f")

    def _render_chart(self, closes: list[float]) -> None:
        self.chart_canvas.delete("all")
        numeric_closes: list[float] = []
        for value in closes:
            try:
                numeric_closes.append(float(value))
            except (TypeError, ValueError):
                continue
        if not numeric_closes:
            self.chart_canvas.create_text(
                220,
                110,
                text="No chart data available for this range.",
                fill="#666",
            )
            return
        if len(numeric_closes) < 2:
            self.chart_canvas.update_idletasks()
            width = max(self.chart_canvas.winfo_width(), 1)
            height = max(self.chart_canvas.winfo_height(), 1)
            padding = 20
            x = width / 2
            y = height / 2
            self.chart_canvas.create_oval(
                x - 4,
                y - 4,
                x + 4,
                y + 4,
                fill="#1f77b4",
                outline="",
            )
            self.chart_canvas.create_text(
                padding,
                padding / 2,
                anchor="w",
                text=f"{numeric_closes[0]:.2f}",
                fill="#1f77b4",
            )
            return
        self.chart_canvas.update_idletasks()
        width = max(self.chart_canvas.winfo_width(), 1)
        height = max(self.chart_canvas.winfo_height(), 1)
        padding = 20
        min_price = min(numeric_closes)
        max_price = max(numeric_closes)
        price_span = max(max_price - min_price, 1e-6)
        x_span = max(len(numeric_closes) - 1, 1)

        points = []
        for idx, price in enumerate(numeric_closes):
            x = padding + (width - 2 * padding) * (idx / x_span)
            y = height - padding - (height - 2 * padding) * ((price - min_price) / price_span)
            points.extend([x, y])

        if len(points) < 4:
            self.chart_canvas.create_text(
                220,
                110,
                text="Not enough chart data to render a line.",
                fill="#666",
            )
            return

        try:
            if len(points) < 4:
                raise tk.TclError("Insufficient points for line rendering.")
            self.chart_canvas.create_line(*points, fill="#1f77b4", width=2, smooth=True)
        except tk.TclError:
            self.chart_canvas.create_text(
                220,
                110,
                text="Unable to render chart line for this data.",
                fill="#666",
            )
            return
        self.chart_canvas.create_text(
            padding,
            padding / 2,
            anchor="w",
            text=f"{numeric_closes[-1]:.2f}",
            fill="#1f77b4",
        )
        self.chart_canvas.create_text(
            width - padding,
            padding / 2,
            anchor="e",
            text=f"{numeric_closes[0]:.2f}",
            fill="#1f77b4",
        )

    def _on_content_configure(self, _event: tk.Event) -> None:
        self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all"))

    def _on_canvas_configure(self, event: tk.Event) -> None:
        self.scroll_canvas.itemconfigure(self.scroll_window, width=event.width)

    def _on_mousewheel(self, event: tk.Event) -> None:
        if self.scroll_canvas.winfo_height() < self.content_frame.winfo_height():
            self.scroll_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all"))

    def _sync_option_snapshot(self) -> None:
        contract = self.option_contract or {}
        self._set_value(self.option_values["contract"], contract.get("ticker"))
        self._set_value(self.option_values["expiration"], contract.get("expiration_date"))
        contract_type = contract.get("contract_type")
        display_type = contract_type.upper() if contract_type else None
        self._set_value(self.option_values["type"], display_type)
        self._set_value(self.option_values["strike"], contract.get("strike_price"))

    def _toggle_info_panels(self) -> None:
        is_stock = self.analysis_mode_var.get() == "Stock Analysis"

        if not self.stock_info_frame.winfo_ismapped():
            self.stock_info_frame.pack(padx=20, pady=(5, 15), fill="x")

        if is_stock:
            self.option_info_frame.pack_forget()
            self.options_frame.pack_forget()
            self.strategy_frame.pack_forget()
            self.knobs_frame.pack_forget()
        else:
            if not self.option_info_frame.winfo_ismapped():
                self.option_info_frame.pack(padx=20, pady=(5, 15), fill="x")
            if not self.options_frame.winfo_ismapped():
                self.options_frame.pack(padx=20, pady=(5, 15), fill="x")
            if not self.strategy_frame.winfo_ismapped():
                self.strategy_frame.pack(pady=5)
            if not self.knobs_frame.winfo_ismapped():
                self.knobs_frame.pack(pady=10, fill="x", padx=40)
        self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all"))

    def refresh(self) -> None:
        ticker = self.controller.state.selected_ticker or "None"
        self.selected_label.config(text=f"Selected Ticker: {ticker}")
        self.analysis_mode_var.set(self.controller.state.analysis_mode)
        self.strategy_var.set(self.controller.state.option_strategy)
        self.delta_var.set(self.controller.state.knob_delta)
        self.risk_var.set(self.controller.state.knob_risk)
        self.prob_var.set(self.controller.state.knob_prob)
        api_key = load_api_key()
        self.api_client = MassiveApiClient(api_key) if api_key else None
        self._toggle_info_panels()
        self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all"))

    def on_analysis_mode_change(self, _event: object) -> None:
        self.controller.state.analysis_mode = self.analysis_mode_var.get()
        self.controller.persist_state()
        self._toggle_info_panels()
        self._sync_option_snapshot()

    def on_strategy_change(self, _event: object) -> None:
        self.controller.state.option_strategy = self.strategy_var.get()
        self.controller.persist_state()

    def load_market_data(self) -> None:
        if not self.api_client:
            messagebox.showinfo(
                "Missing key", "Enter or set a Massive API key to load data."
            )
            return
        ticker = self.controller.state.selected_ticker
        if not ticker:
            messagebox.showinfo("Missing ticker", "Select a ticker first.")
            return
        try:
            stock_data = self.api_client.fetch_previous_close(ticker)
            option_data = self.api_client.fetch_option_contracts(ticker)
            horizon_index = int(round(self.horizon_var.get()))
            horizon_index = min(max(horizon_index, 0), len(HORIZON_CONFIGS) - 1)
            _label, days_back, minutes_per_bar, _cadence_label = HORIZON_CONFIGS[horizon_index]
            aggregates = self.api_client.fetch_aggregates(ticker, days_back, minutes_per_bar)
        except HTTPError as exc:
            messagebox.showerror(
                "API Error",
                f"Massive API returned an error: {exc.code} {exc.reason}",
            )
            return
        except URLError as exc:
            messagebox.showerror(
                "Connection Error",
                f"Could not reach Massive API: {exc.reason}",
            )
            return

        self._set_value(self.stock_values["price"], stock_data.get("close"))
        self._set_value(self.stock_values["prev_close"], stock_data.get("close"))
        self._set_value(self.stock_values["open"], stock_data.get("open"))
        self._set_value(self.stock_values["high"], stock_data.get("high"))
        self._set_value(self.stock_values["low"], stock_data.get("low"))
        self._set_value(self.stock_values["volume"], stock_data.get("volume"))
        self._set_value(self.stock_values["market_cap"], "--")
        self._set_value(self.stock_values["range_52w"], "--")
        self.option_contract = option_data[0] if option_data else None
        self._sync_option_snapshot()

        closes = [item.get("c") for item in aggregates if item.get("c") is not None]
        self._render_chart(closes)

        self.options_text.configure(state="normal")
        self.options_text.delete("1.0", tk.END)
        if not option_data:
            self.options_text.insert("1.0", "No option contracts returned.\n")
        else:
            lines = []
            for contract in option_data[:5]:
                lines.append(
                    "{ticker} {expiration} {type} {strike}".format(
                        ticker=contract.get("ticker", "--"),
                        expiration=contract.get("expiration_date", "--"),
                        type=contract.get("contract_type", "--").upper(),
                        strike=contract.get("strike_price", "--"),
                    )
                )
            self.options_text.insert("1.0", "\n".join(lines))
        self.options_text.configure(state="disabled")

    def save_analysis(self) -> None:
        self.controller.state.analysis_mode = self.analysis_mode_var.get()
        self.controller.state.option_strategy = self.strategy_var.get()
        self.controller.state.knob_delta = int(self.delta_var.get())
        self.controller.state.knob_risk = int(self.risk_var.get())
        self.controller.state.knob_prob = int(self.prob_var.get())
        self.controller.persist_state()
        self._sync_option_snapshot()
        messagebox.showinfo("Saved", "Analysis settings saved locally.")


if __name__ == "__main__":
    app = StoptionsApp()
    app.mainloop()

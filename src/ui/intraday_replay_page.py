from __future__ import annotations

from collections import defaultdict
from datetime import datetime, timedelta
import threading
import tkinter as tk
from tkinter import messagebox, ttk
from zoneinfo import ZoneInfo

from data_access.api_client import MassiveApiClient
from ui.helpers import load_api_key

EASTERN_TZ = ZoneInfo("America/New_York")
REPLAY_WINDOWS_MINUTES = (10, 20, 30, 60)
SESSION_OPEN_MINUTE = 9 * 60 + 30
SESSION_CLOSE_MINUTE = 16 * 60


def _group_bars_by_day(bars: list[dict]) -> list[tuple[str, list[dict]]]:
    grouped: dict[str, list[dict]] = defaultdict(list)
    for bar in bars:
        ts = bar.get("t")
        close = bar.get("c")
        if ts is None or close is None:
            continue
        dt = datetime.fromtimestamp(ts / 1000, tz=EASTERN_TZ)
        day_key = dt.date().isoformat()
        grouped[day_key].append({"t": ts, "c": close})
    ordered: list[tuple[str, list[dict]]] = []
    for day in sorted(grouped):
        ordered.append((day, sorted(grouped[day], key=lambda row: row["t"])))
    return ordered


def _minute_of_day(ts_ms: int) -> int:
    dt = datetime.fromtimestamp(ts_ms / 1000, tz=EASTERN_TZ)
    return dt.hour * 60 + dt.minute


class IntradayReplayPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller
        self.api_client: MassiveApiClient | None = None
        self.day_series: list[tuple[str, list[dict]]] = []
        self.day_index = 0
        self.stage_index = 0
        self._job_id: str | None = None

        ttk.Label(self, text="Intraday Replay", font=("Arial", 18, "bold")).pack(pady=(14, 6))
        ttk.Label(
            self,
            text="Use ←/→ arrow keys to change trading day. Press Play to animate early-session convexity/concavity windows.",
            wraplength=860,
            justify="center",
        ).pack(pady=(0, 10))

        controls = ttk.Frame(self)
        controls.pack(fill="x", padx=18, pady=8)

        default_ticker = str(self.controller.state.selected_ticker or "").strip().upper()
        ttk.Label(controls, text="Ticker").grid(row=0, column=0, sticky="w")
        self.ticker_var = tk.StringVar(value=default_ticker)
        ttk.Entry(controls, textvariable=self.ticker_var, width=10).grid(row=0, column=1, padx=(8, 18), sticky="w")

        ttk.Label(controls, text="Days back").grid(row=0, column=2, sticky="w")
        self.days_back_var = tk.StringVar(value="14")
        ttk.Entry(controls, textvariable=self.days_back_var, width=6).grid(row=0, column=3, padx=(8, 18), sticky="w")

        ttk.Label(controls, text="Speed multiplier").grid(row=0, column=4, sticky="w")
        self.speed_var = tk.StringVar(value="1.0")
        ttk.Entry(controls, textvariable=self.speed_var, width=8).grid(row=0, column=5, padx=(8, 18), sticky="w")

        ttk.Button(controls, text="Load data", command=self.load_data).grid(row=0, column=6, padx=6)
        ttk.Button(controls, text="Play", command=self.start_replay).grid(row=0, column=7, padx=6)
        ttk.Button(controls, text="Stop", command=self.stop_replay).grid(row=0, column=8, padx=6)
        ttk.Button(controls, text="Back to Main Menu", command=lambda: controller.show_frame("MainMenu")).grid(
            row=0, column=9, padx=(24, 0)
        )

        self.status_var = tk.StringVar(value="Load data to begin.")
        ttk.Label(self, textvariable=self.status_var).pack(pady=(4, 8))

        self.chart = tk.Canvas(self, bg="#0f172a", highlightthickness=0, height=520)
        self.chart.pack(fill="both", expand=True, padx=18, pady=(0, 16))

        self.bind_all("<Left>", self._on_left_key)
        self.bind_all("<Right>", self._on_right_key)

    def refresh(self) -> None:
        if hasattr(self, "ticker_var"):
            selected = str(self.controller.state.selected_ticker or "").strip().upper()
            if selected and not self.ticker_var.get().strip():
                self.ticker_var.set(selected)
        self.focus_force()
        if self.day_series:
            self._draw_current_day()

    def load_data(self) -> None:
        ticker = self._read_ticker()
        if not ticker:
            messagebox.showinfo("Missing ticker", "Enter a ticker (e.g., NVDA) or select one on Select Stock.")
            return
        self.controller.state.selected_ticker = ticker
        self.controller.persist_state()
        api_key = load_api_key()
        if not api_key:
            messagebox.showinfo("Missing API key", "Save your Massive API key in the main menu first.")
            return
        self.status_var.set(f"Loading minute bars for {ticker}...")

        def worker() -> None:
            try:
                client = MassiveApiClient(api_key)
                end_date = datetime.now(tz=EASTERN_TZ).date()
                days_back = self._read_days_back()
                start_date = end_date - timedelta(days=days_back)
                bars = client.fetch_aggregates_range(ticker, start_date, end_date, minutes_per_bar=1)
                series = _group_bars_by_day(bars)
            except Exception as exc:  # noqa: BLE001
                self.after(0, lambda: self.status_var.set(f"Load failed: {exc}"))
                return

            def apply() -> None:
                self.api_client = client
                self.day_series = series
                self.day_index = max(0, len(series) - 1)
                self.stage_index = 0
                if not self.day_series:
                    self.status_var.set(f"No intraday data returned for {ticker}.")
                    self._draw_empty("No intraday data available")
                    return
                self.status_var.set(
                    f"Loaded {len(self.day_series)} trading days for {ticker}. Use ←/→ keys to change days."
                )
                self._draw_current_day()

            self.after(0, apply)

        threading.Thread(target=worker, daemon=True).start()

    def shift_day(self, offset: int, autoplay: bool = False) -> None:
        if not self.day_series:
            self.status_var.set("No data loaded. Press Load data first.")
            return
        next_index = self.day_index + offset
        if next_index < 0:
            self.status_var.set("Already at oldest available day.")
            return
        if next_index >= len(self.day_series):
            self.status_var.set("Already at newest available day.")
            return
        self.stop_replay()
        self.day_index = next_index
        self.stage_index = 0
        if autoplay:
            self._schedule_next_stage()
            return
        self._draw_current_day()


    def _on_left_key(self, _event: tk.Event) -> None:
        if not self.winfo_viewable():
            return
        self.shift_day(-1, autoplay=True)

    def _on_right_key(self, _event: tk.Event) -> None:
        if not self.winfo_viewable():
            return
        self.shift_day(1, autoplay=True)

    def _read_ticker(self) -> str:
        return self.ticker_var.get().strip().upper()

    def _read_days_back(self) -> int:
        raw = self.days_back_var.get().strip()
        try:
            value = int(raw)
        except ValueError:
            value = 14
        return min(max(value, 1), 60)

    def start_replay(self) -> None:
        if not self.day_series:
            self.load_data()
            return
        self.stop_replay()
        self.stage_index = 0
        self._schedule_next_stage()

    def stop_replay(self) -> None:
        self._job_id = None

    def _schedule_next_stage(self) -> None:
        if not self.day_series:
            return
        day, bars = self.day_series[self.day_index]
        stage_total = len(REPLAY_WINDOWS_MINUTES) + 1
        stage = min(self.stage_index, stage_total - 1)
        if stage < len(REPLAY_WINDOWS_MINUTES):
            cutoff = SESSION_OPEN_MINUTE + REPLAY_WINDOWS_MINUTES[stage]
            stage_label = f"first {REPLAY_WINDOWS_MINUTES[stage]} minutes"
        else:
            cutoff = SESSION_CLOSE_MINUTE + 1
            stage_label = "full regular session"
        self._draw_current_day(cutoff_minute=cutoff)
        self.status_var.set(f"{day} — showing {stage_label}")

        if stage >= stage_total - 1:
            self._job_id = None
            return

        self.stage_index += 1
        speed = self._read_speed_multiplier()
        delay_ms = max(80, int(1000 / speed))
        token = f"{datetime.now(tz=EASTERN_TZ).timestamp()}"
        self._job_id = token

        def continue_replay() -> None:
            if self._job_id != token:
                return
            self._schedule_next_stage()

        self.after(delay_ms, continue_replay)

    def _read_speed_multiplier(self) -> float:
        try:
            parsed = float(self.speed_var.get().strip())
        except ValueError:
            parsed = 1.0
        return min(max(parsed, 0.1), 10.0)

    def _draw_current_day(self, cutoff_minute: int | None = None) -> None:
        if not self.day_series:
            self._draw_empty("No data loaded")
            return
        day, bars = self.day_series[self.day_index]
        visible = [bar for bar in bars if cutoff_minute is None or _minute_of_day(int(bar["t"])) <= cutoff_minute]
        if not visible:
            self._draw_empty(f"{day} has no visible points for this stage")
            return

        self.chart.delete("all")
        width = max(self.chart.winfo_width(), 800)
        height = max(self.chart.winfo_height(), 420)
        left, right, top, bottom = 76, width - 24, 28, height - 54

        self.chart.create_rectangle(left, top, right, bottom, fill="#111827", outline="#1f2937")

        prices = [float(row["c"]) for row in visible]
        min_p, max_p = min(prices), max(prices)
        pad = (max_p - min_p) * 0.15 if max_p != min_p else max_p * 0.005 or 1.0
        y_min, y_max = min_p - pad, max_p + pad

        x_start = SESSION_OPEN_MINUTE
        x_end = SESSION_CLOSE_MINUTE

        for hour in range(10, 17):
            minute = hour * 60
            x = left + (minute - x_start) / (x_end - x_start) * (right - left)
            self.chart.create_line(x, top, x, bottom, fill="#1f2937")
            label = f"{hour}:00"
            self.chart.create_text(x, bottom + 18, text=label, fill="#94a3b8", font=("Arial", 9))

        for step in range(6):
            y = top + step / 5 * (bottom - top)
            self.chart.create_line(left, y, right, y, fill="#1f2937")
            p = y_max - (step / 5) * (y_max - y_min)
            self.chart.create_text(left - 10, y, text=f"{p:.2f}", fill="#cbd5e1", anchor="e", font=("Arial", 9))

        points: list[float] = []
        for row in visible:
            minute = _minute_of_day(int(row["t"]))
            px = float(row["c"])
            x = left + (minute - x_start) / (x_end - x_start) * (right - left)
            y = bottom - (px - y_min) / (y_max - y_min) * (bottom - top)
            points.extend((x, y))

        if len(points) >= 4:
            self.chart.create_line(*points, fill="#38bdf8", width=2.5, smooth=True)

        self.chart.create_text(
            width // 2,
            14,
            text=f"{self.controller.state.selected_ticker or ''} — {day}",
            fill="#e2e8f0",
            font=("Arial", 12, "bold"),
        )
        self.chart.create_text(width // 2, height - 20, text="Time (ET)", fill="#cbd5e1", font=("Arial", 10))
        self.chart.create_text(22, (top + bottom) // 2, text="Price", fill="#cbd5e1", angle=90, font=("Arial", 10))

    def _draw_empty(self, text: str) -> None:
        self.chart.delete("all")
        width = max(self.chart.winfo_width(), 800)
        height = max(self.chart.winfo_height(), 420)
        self.chart.create_rectangle(0, 0, width, height, fill="#0f172a", outline="#0f172a")
        self.chart.create_text(width // 2, height // 2, text=text, fill="#94a3b8", font=("Arial", 13))

from __future__ import annotations

import json
import math
import os
import random
import socket
import time
import threading
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, time as dt_time, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk
from urllib.error import HTTPError, URLError
from zoneinfo import ZoneInfo

import numpy as np

from analysis.cross_sectional import (
    CrossSectionalSettings,
    MomentumSettings,
    STRATEGY_REGISTRY,
    compute_cross_sectional_momentum,
)
from analysis.reporting import format_cross_sectional_report, format_time_series_report
from analysis.time_series import (
    TIME_SERIES_STRATEGY_REGISTRY,
    TimeSeriesMomentumSettings,
    TimeSeriesSettings,
    compute_time_series_momentum,
)
from config import (
    ANALYSIS_OUTPUT_DIR,
    API_KEY_PATH,
    BACKTEST_CACHE_DIR,
    BACKTEST_OUTPUT_DIR,
    CONFIG_DIR,
    DATA_DIR,
    DEFAULT_BACKTEST_SETTINGS,
    DEFAULT_GENERAL_ANALYSIS_SETTINGS,
    HORIZON_CONFIGS,
)
from data_access.api_client import MassiveApiClient
from data_access.cache import _safe_ticker_name, load_cached_market_data, save_cached_market_data
from data_access.option_loader import load_option_records
from ui.helpers import load_api_key, save_api_key
from utils.parsing import (
    _coerce_number,
    _get_nested_value,
    _has_fundamentals_data,
    _parse_iso_date,
    build_npz_payload,
    chunk_results_by_year,
    combine_greeks,
    effective_market_date,
    extract_greeks,
    format_http_error_detail,
    format_strike,
    normalize_cache_root,
    normalize_contract_type,
    normalize_likelihood_threshold,
    normalize_option_records,
    option_likelihood,
    option_mid_price,
    parse_date,
    parse_float,
    parse_int,
)


class SpreadAnalysisPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller
        self.api_client: MassiveApiClient | None = None
        self.spread_records: list[dict] = []
        self.all_spread_records: list[dict] = []
        self.selected_spread: dict | None = None

        header = ttk.Label(self, text="Spread Analysis", font=("Arial", 18, "bold"))
        header.pack(pady=10)

        self.strategy_label = ttk.Label(self, text="Strategy: --", font=("Arial", 12))
        self.strategy_label.pack(pady=(0, 10))

        self.options_frame = ttk.LabelFrame(self, text="Spread Candidates")
        self.options_frame.pack(padx=30, pady=10, fill="both", expand=True)
        self.options_frame.columnconfigure(0, weight=1)
        self.options_frame.columnconfigure(1, weight=0)
        self.options_frame.rowconfigure(0, weight=1)

        list_frame = ttk.Frame(self.options_frame)
        list_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=8)
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        filter_frame = ttk.Frame(self.options_frame)
        filter_frame.grid(row=0, column=1, sticky="nsew", padx=(0, 10), pady=8)
        filter_frame.columnconfigure(1, weight=1)

        ttk.Label(filter_frame, text="Max Loss / Spread Price").grid(
            row=0, column=0, padx=5, pady=2, sticky="w"
        )
        self.max_loss_var = tk.StringVar()
        self.max_loss_entry = ttk.Entry(filter_frame, textvariable=self.max_loss_var, width=12)
        self.max_loss_entry.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        self.max_loss_entry.bind("<KeyRelease>", self.on_filter_change)

        ttk.Label(filter_frame, text="Min Likelihood (%)").grid(
            row=1, column=0, padx=5, pady=2, sticky="w"
        )
        self.likelihood_var = tk.StringVar()
        self.likelihood_entry = ttk.Entry(
            filter_frame, textvariable=self.likelihood_var, width=12
        )
        self.likelihood_entry.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        self.likelihood_entry.bind("<KeyRelease>", self.on_filter_change)

        ttk.Label(filter_frame, text="Expiration").grid(
            row=2, column=0, padx=5, pady=2, sticky="w"
        )
        self.expiration_var = tk.StringVar(value="All")
        self.expiration_dropdown = ttk.Combobox(
            filter_frame, textvariable=self.expiration_var, state="readonly", width=18
        )
        self.expiration_dropdown.grid(row=2, column=1, padx=5, pady=2, sticky="ew")
        self.expiration_dropdown.bind("<<ComboboxSelected>>", self.on_filter_change)

        ttk.Label(filter_frame, text="Strike").grid(
            row=3, column=0, padx=5, pady=2, sticky="w"
        )
        self.strike_var = tk.StringVar(value="All")
        self.strike_dropdown = ttk.Combobox(
            filter_frame, textvariable=self.strike_var, state="readonly", width=12
        )
        self.strike_dropdown.grid(row=3, column=1, padx=5, pady=2, sticky="ew")
        self.strike_dropdown.bind("<<ComboboxSelected>>", self.on_filter_change)

        ttk.Label(filter_frame, text="Type").grid(row=4, column=0, padx=5, pady=2, sticky="w")
        self.type_var = tk.StringVar(value="All")
        self.type_dropdown = ttk.Combobox(
            filter_frame, textvariable=self.type_var, state="readonly", width=10
        )
        self.type_dropdown.grid(row=4, column=1, padx=5, pady=2, sticky="ew")
        self.type_dropdown.bind("<<ComboboxSelected>>", self.on_filter_change)

        self.options_list = tk.Listbox(list_frame, height=12, width=48)
        options_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.options_list.yview)
        self.options_list.configure(yscrollcommand=options_scroll.set)
        self.options_list.grid(row=0, column=0, sticky="nsew")
        options_scroll.grid(row=0, column=1, sticky="ns")
        self.options_list.bind("<<ListboxSelect>>", self.on_spread_select)

        self.spread_info_frame = ttk.LabelFrame(self, text="Selected Spread")
        self.spread_info_frame.pack(padx=30, pady=(5, 10), fill="x")
        self.spread_values: dict[str, ttk.Label] = {}
        self._build_info_grid(
            self.spread_info_frame,
            [
                ("Strategy", "strategy"),
                ("Expiration(s)", "expiration"),
                ("Type", "type"),
                ("Strikes", "strikes"),
                ("Spread Price", "premium"),
                ("Likelihood", "likelihood"),
            ],
            self.spread_values,
            columns=3,
        )

        self.legs_frame = ttk.LabelFrame(self, text="Leg Details")
        self.legs_frame.pack(padx=30, pady=(5, 10), fill="x")
        self.leg_values: dict[str, ttk.Label] = {}
        self._build_info_grid(
            self.legs_frame,
            [
                ("Long Leg", "long_leg"),
                ("Short Leg", "short_leg"),
            ],
            self.leg_values,
            columns=1,
        )

        self.greeks_frame = ttk.LabelFrame(self, text="Spread Greeks")
        self.greeks_frame.pack(padx=30, pady=(5, 10), fill="x")
        self.greeks_values: dict[str, ttk.Label] = {}
        self._build_info_grid(
            self.greeks_frame,
            [
                ("Delta", "delta"),
                ("Gamma", "gamma"),
                ("Theta", "theta"),
                ("Vega", "vega"),
                ("Rho", "rho"),
                ("IV", "iv"),
            ],
            self.greeks_values,
            columns=3,
        )

        button_row = ttk.Frame(self)
        button_row.pack(pady=10)
        ttk.Button(button_row, text="Refresh", command=self.load_market_data).grid(
            row=0, column=0, padx=10
        )
        ttk.Button(
            button_row,
            text="Select Stock",
            command=lambda: controller.show_frame("TickerSelectPage"),
        ).grid(row=0, column=1, padx=10)
        ttk.Button(
            button_row,
            text="Back to Analysis",
            command=lambda: controller.show_frame("AnalysisPage"),
        ).grid(row=0, column=2, padx=10)
        ttk.Button(
            button_row,
            text="Back to Main Menu",
            command=lambda: controller.show_frame("MainMenu"),
        ).grid(row=0, column=3, padx=10)

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

    def _format_float(self, value: float) -> str:
        decimals = 2 if abs(value) >= 1 else 4
        multiplier = 10**decimals
        truncated = math.trunc(value * multiplier) / multiplier
        return f"{truncated:.{decimals}f}".rstrip("0").rstrip(".")

    def _set_value(self, label: ttk.Label, value: str | int | float | None) -> None:
        if value in (None, "", "--"):
            label.config(text="--", foreground="#b00020")
        else:
            if isinstance(value, float):
                text = self._format_float(value)
            elif isinstance(value, int):
                text = str(value)
            else:
                text = str(value)
            label.config(text=text, foreground="#0a7a2f")

    def refresh(self) -> None:
        self.api_client = MassiveApiClient(load_api_key()) if load_api_key() else None
        self.load_market_data()

    def load_market_data(self) -> None:
        if not self.api_client:
            messagebox.showinfo(
                "Missing key", "Enter or set a Massive API key to load options data."
            )
            return
        ticker = self.controller.state.selected_ticker
        if not ticker:
            messagebox.showinfo("Missing ticker", "Select a ticker first.")
            return
        strategy = self.controller.state.option_strategy
        self.strategy_label.config(text=f"Strategy: {strategy}")
        try:
            option_records = load_option_records(self.api_client, ticker)
        except HTTPError as exc:
            self._show_api_error(exc, "Massive", "Verify your Massive API key.")
            return
        except URLError as exc:
            self._show_error_dialog(
                "Connection Error",
                f"Could not reach Massive API endpoint: {exc.reason}",
            )
            return
        self.all_spread_records = self._build_spread_records(option_records, strategy)
        self._refresh_spread_filters(reset=True)

    def _build_spread_records(self, option_records: list[dict], strategy: str) -> list[dict]:
        spreads: list[dict] = []
        if strategy == "Calendar Spread":
            grouped: dict[tuple[str | None, str | None], list[dict]] = {}
            for record in option_records:
                key = (normalize_contract_type(record.get("contract_type")), record.get("strike_price"))
                grouped.setdefault(key, []).append(record)
            for (contract_type, strike), records in grouped.items():
                records_sorted = sorted(
                    records,
                    key=lambda r: r.get("expiration_date") or "",
                )
                for short_leg, long_leg in zip(records_sorted, records_sorted[1:]):
                    premium_long = option_mid_price(long_leg)
                    premium_short = option_mid_price(short_leg)
                    net_premium = None
                    if isinstance(premium_long, (int, float)) or isinstance(
                        premium_short, (int, float)
                    ):
                        net_premium = (premium_long or 0) - (premium_short or 0)
                        net_premium = abs(net_premium)
                    likelihood_values = [
                        option_likelihood(long_leg),
                        option_likelihood(short_leg),
                    ]
                    likelihoods = [
                        value for value in likelihood_values if isinstance(value, (int, float))
                    ]
                    likelihood = sum(likelihoods) / len(likelihoods) if likelihoods else None
                    expiration_label = "{short}→{long}".format(
                        short=short_leg.get("expiration_date", "--"),
                        long=long_leg.get("expiration_date", "--"),
                    )
                    strike_label = format_strike(strike)
                    spreads.append(
                        {
                            "strategy": strategy,
                            "ticker": short_leg.get("ticker") or long_leg.get("ticker"),
                            "expiration_label": expiration_label,
                            "contract_type": contract_type,
                            "strike_pair": strike_label,
                            "long_leg": long_leg,
                            "short_leg": short_leg,
                            "net_premium": net_premium,
                            "likelihood": likelihood,
                            "greeks": combine_greeks(long_leg, short_leg),
                        }
                    )
        else:
            grouped: dict[tuple[str | None, str | None], list[dict]] = {}
            for record in option_records:
                key = (
                    normalize_contract_type(record.get("contract_type")),
                    record.get("expiration_date"),
                )
                grouped.setdefault(key, []).append(record)
            for (contract_type, expiration), records in grouped.items():
                records_sorted = sorted(
                    records,
                    key=lambda r: float(r.get("strike_price") or 0),
                )
                for lower, higher in zip(records_sorted, records_sorted[1:]):
                    if contract_type == "PUT":
                        long_leg = higher
                        short_leg = lower
                    else:
                        long_leg = lower
                        short_leg = higher
                    premium_long = option_mid_price(long_leg)
                    premium_short = option_mid_price(short_leg)
                    net_premium = None
                    if isinstance(premium_long, (int, float)) or isinstance(
                        premium_short, (int, float)
                    ):
                        net_premium = (premium_long or 0) - (premium_short or 0)
                        net_premium = abs(net_premium)
                    likelihood_values = [
                        option_likelihood(long_leg),
                        option_likelihood(short_leg),
                    ]
                    likelihoods = [
                        value for value in likelihood_values if isinstance(value, (int, float))
                    ]
                    likelihood = sum(likelihoods) / len(likelihoods) if likelihoods else None
                    strike_pair = "{low}/{high}".format(
                        low=format_strike(lower.get("strike_price")),
                        high=format_strike(higher.get("strike_price")),
                    )
                    spreads.append(
                        {
                            "strategy": strategy,
                            "ticker": lower.get("ticker") or higher.get("ticker"),
                            "expiration_label": expiration,
                            "contract_type": contract_type,
                            "strike_pair": strike_pair,
                            "long_leg": long_leg,
                            "short_leg": short_leg,
                            "net_premium": net_premium,
                            "likelihood": likelihood,
                            "greeks": combine_greeks(long_leg, short_leg),
                        }
                    )
        return spreads

    def _get_filter_value(self, var: tk.StringVar) -> str | None:
        value = var.get()
        return None if value == "All" else value

    def _record_matches_filters(self, record: dict, filters: dict[str, str | None]) -> bool:
        expiration = record.get("expiration_label")
        strike = record.get("strike_pair")
        contract_type = normalize_contract_type(record.get("contract_type"))
        if filters.get("expiration") and filters["expiration"] != expiration:
            return False
        if filters.get("strike") and filters["strike"] != strike:
            return False
        if filters.get("type") and filters["type"] != contract_type:
            return False
        return True

    def _record_matches_constraints(self, record: dict) -> bool:
        max_loss = parse_float(self.max_loss_var.get())
        min_likelihood = normalize_likelihood_threshold(self.likelihood_var.get())
        premium = record.get("net_premium")
        likelihood = record.get("likelihood")
        if max_loss is not None:
            if not isinstance(premium, (int, float)) or premium > max_loss:
                return False
        if min_likelihood is not None:
            if not isinstance(likelihood, (int, float)) or likelihood < min_likelihood:
                return False
        return True

    def _compute_filter_options(
        self, records: list[dict], current: dict[str, str | None]
    ) -> dict[str, list[str]]:
        options: dict[str, set[str]] = {"expiration": set(), "strike": set(), "type": set()}
        for record in records:
            expiration = record.get("expiration_label")
            strike = record.get("strike_pair")
            contract_type = normalize_contract_type(record.get("contract_type"))
            if self._record_matches_filters(record, {**current, "expiration": None}):
                if expiration:
                    options["expiration"].add(expiration)
            if self._record_matches_filters(record, {**current, "strike": None}):
                if strike:
                    options["strike"].add(strike)
            if self._record_matches_filters(record, {**current, "type": None}):
                if contract_type:
                    options["type"].add(contract_type)
        return {
            "expiration": sorted(options["expiration"]),
            "strike": sorted(options["strike"]),
            "type": sorted(options["type"]),
        }

    def _refresh_spread_filters(self, reset: bool = False) -> None:
        if reset:
            self.expiration_var.set("All")
            self.strike_var.set("All")
            self.type_var.set("All")
        eligible_records = [
            record for record in self.all_spread_records if self._record_matches_constraints(record)
        ]
        filters = {
            "expiration": self._get_filter_value(self.expiration_var),
            "strike": self._get_filter_value(self.strike_var),
            "type": self._get_filter_value(self.type_var),
        }
        options = self._compute_filter_options(eligible_records, filters)
        for key, dropdown, var in (
            ("expiration", self.expiration_dropdown, self.expiration_var),
            ("strike", self.strike_dropdown, self.strike_var),
            ("type", self.type_dropdown, self.type_var),
        ):
            values = ["All"] + options[key]
            dropdown["values"] = values
            if var.get() not in values:
                var.set("All")
        self._apply_spread_filters(eligible_records)

    def _apply_spread_filters(self, eligible_records: list[dict]) -> None:
        filters = {
            "expiration": self._get_filter_value(self.expiration_var),
            "strike": self._get_filter_value(self.strike_var),
            "type": self._get_filter_value(self.type_var),
        }
        self.spread_records = [
            record
            for record in eligible_records
            if self._record_matches_filters(record, filters)
        ]
        self.options_list.delete(0, tk.END)
        if not self.spread_records:
            self.options_list.insert(tk.END, "No spreads returned.")
            self.selected_spread = None
        else:
            for spread in self.spread_records:
                likelihood = spread.get("likelihood")
                likelihood_label = (
                    f"{likelihood * 100:.0f}%" if isinstance(likelihood, (int, float)) else "--"
                )
                premium = spread.get("net_premium")
                premium_label = f"{premium:.2f}" if isinstance(premium, (int, float)) else "--"
                self.options_list.insert(
                    tk.END,
                    "{ticker} {expiration} {type} {strike} | Loss {loss} | Likely {likelihood}".format(
                        ticker=spread.get("ticker", "--"),
                        expiration=spread.get("expiration_label", "--"),
                        type=str(spread.get("contract_type", "--")).upper(),
                        strike=spread.get("strike_pair", "--"),
                        loss=premium_label,
                        likelihood=likelihood_label,
                    ),
                )
            self.options_list.selection_set(0)
            self.options_list.see(0)
            self.selected_spread = self.spread_records[0]
        self._sync_spread_snapshot()

    def _sync_spread_snapshot(self) -> None:
        spread = self.selected_spread or {}
        self._set_value(self.spread_values["strategy"], spread.get("strategy"))
        self._set_value(self.spread_values["expiration"], spread.get("expiration_label"))
        contract_type = normalize_contract_type(spread.get("contract_type"))
        self._set_value(self.spread_values["type"], contract_type)
        self._set_value(self.spread_values["strikes"], spread.get("strike_pair"))
        premium = spread.get("net_premium")
        self._set_value(self.spread_values["premium"], premium)
        likelihood = spread.get("likelihood")
        likelihood_label = (
            f"{likelihood * 100:.1f}%" if isinstance(likelihood, (int, float)) else None
        )
        self._set_value(self.spread_values["likelihood"], likelihood_label)
        long_leg = spread.get("long_leg") or {}
        short_leg = spread.get("short_leg") or {}
        long_label = "{ticker} {exp} {type} {strike}".format(
            ticker=long_leg.get("ticker", "--"),
            exp=long_leg.get("expiration_date", "--"),
            type=str(long_leg.get("contract_type", "--")).upper(),
            strike=long_leg.get("strike_price", "--"),
        )
        short_label = "{ticker} {exp} {type} {strike}".format(
            ticker=short_leg.get("ticker", "--"),
            exp=short_leg.get("expiration_date", "--"),
            type=str(short_leg.get("contract_type", "--")).upper(),
            strike=short_leg.get("strike_price", "--"),
        )
        self._set_value(self.leg_values["long_leg"], long_label)
        self._set_value(self.leg_values["short_leg"], short_label)
        greeks = spread.get("greeks") or {}
        self._set_value(self.greeks_values["delta"], greeks.get("delta"))
        self._set_value(self.greeks_values["gamma"], greeks.get("gamma"))
        self._set_value(self.greeks_values["theta"], greeks.get("theta"))
        self._set_value(self.greeks_values["vega"], greeks.get("vega"))
        self._set_value(self.greeks_values["rho"], greeks.get("rho"))
        self._set_value(self.greeks_values["iv"], greeks.get("iv"))

    def on_spread_select(self, _event: object) -> None:
        selection = self.options_list.curselection()
        if not selection:
            return
        index = selection[0]
        if index >= len(self.spread_records):
            return
        self.selected_spread = self.spread_records[index]
        self._sync_spread_snapshot()

    def on_filter_change(self, _event: object) -> None:
        self._refresh_spread_filters()

    def _show_api_error(self, exc: HTTPError, service: str, hint: str | None = None) -> None:
        detail = format_http_error_detail(exc)
        detail_msg = f"\nDetails: {detail}" if detail else ""
        hint_msg = f"\n{hint}" if hint else ""
        self._show_error_dialog(
            "API Error",
            f"{service} API returned an error: {exc.code} {exc.reason}.{detail_msg}{hint_msg}",
        )

    def _show_error_dialog(self, title: str, message: str) -> None:
        dialog = tk.Toplevel(self)
        dialog.title(title)
        dialog.geometry("620x320")
        dialog.transient(self)

        dialog.rowconfigure(0, weight=1)
        dialog.columnconfigure(0, weight=1)

        text_frame = ttk.Frame(dialog)
        text_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

        text_widget = tk.Text(text_frame, wrap="word")
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        text_widget.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        text_widget.insert("1.0", message)
        text_widget.focus_set()

        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=1, column=0, pady=(0, 10))

        def copy_to_clipboard() -> None:
            dialog.clipboard_clear()
            dialog.clipboard_append(text_widget.get("1.0", "end-1c"))
            dialog.update_idletasks()

        ttk.Button(button_frame, text="Copy", command=copy_to_clipboard).grid(
            row=0, column=0, padx=5
        )
        ttk.Button(button_frame, text="Close", command=dialog.destroy).grid(
            row=0, column=1, padx=5
        )

if __name__ == "__main__":
    app = StoptionsApp()
    app.mainloop()

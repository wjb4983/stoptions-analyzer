from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path

import pytest

from state import AppState
from ui import analysis_page, backtesting_page, call_put_analysis_page, general_analysis_page, main_menu, research_lab_page, spread_analysis_page, ticker_entry_page, ticker_select_page


class Var:
    def __init__(self, value=None):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class FakeText:
    def __init__(self, value: str = ""):
        self.value = value

    def get(self, *_args):
        return self.value

    def delete(self, *_args):
        self.value = ""

    def insert(self, *_args):
        self.value = _args[-1]


class FakeListbox:
    def __init__(self):
        self.items: list[str] = []
        self._selection: tuple[int, ...] = ()

    def delete(self, *_args):
        self.items = []

    def insert(self, _idx, item):
        self.items.append(item)

    def selection_set(self, idx: int):
        self._selection = (idx,)

    def curselection(self):
        return self._selection

    def see(self, _idx: int):
        return None

    def get(self, idx: int):
        return self.items[idx]


class FakeWidget:
    def __init__(self):
        self.config_calls = []

    def config(self, **kwargs):
        self.config_calls.append(kwargs)


@dataclass
class FakeController:
    state: AppState
    api_key: str = ""

    def __post_init__(self):
        self.persist_count = 0
        self.frames: list[str] = []

    def persist_state(self):
        self.persist_count += 1

    def show_frame(self, name: str):
        self.frames.append(name)


def test_ticker_entry_and_select_missing_and_success(monkeypatch):
    infos = []
    monkeypatch.setattr(ticker_entry_page.messagebox, "showinfo", lambda title, msg: infos.append((title, msg)))
    monkeypatch.setattr(ticker_select_page.messagebox, "showinfo", lambda title, msg: infos.append((title, msg)))

    controller = FakeController(AppState())

    entry = ticker_entry_page.TickerEntryPage.__new__(ticker_entry_page.TickerEntryPage)
    entry.controller = controller
    entry.text_box = FakeText("\n\n")
    entry.save_tickers()
    assert infos[-1][0] == "No tickers"

    entry.text_box = FakeText("aapl\nmsft\n")
    entry.save_tickers()
    assert controller.state.tickers == ["AAPL", "MSFT"]
    assert controller.state.selected_ticker == "AAPL"

    select = ticker_select_page.TickerSelectPage.__new__(ticker_select_page.TickerSelectPage)
    select.controller = controller
    select.ticker_list = FakeListbox()
    select.refresh()
    assert select.ticker_list.items == ["AAPL", "MSFT"]

    select.ticker_list._selection = ()
    select.use_selected()
    assert infos[-1][0] == "Select a ticker"

    select.ticker_list._selection = (1,)
    select.use_selected()
    assert controller.state.selected_ticker == "MSFT"
    assert controller.frames[-1] == "AnalysisPage"


def test_backtesting_invalid_date_no_signal_and_output_artifacts(monkeypatch, tmp_path):
    infos = []
    monkeypatch.setattr(backtesting_page.messagebox, "showinfo", lambda title, msg: infos.append((title, msg)))

    controller = FakeController(AppState(tickers=["AAPL"]))
    page = backtesting_page.BacktestingPage.__new__(backtesting_page.BacktestingPage)
    page.controller = controller
    page._update_validation_hint = lambda: False
    page._validate_common_inputs = lambda: ("momentum", 20, 1, 1.0, 10000.0, 0.1)
    page.start_date_var = Var("2024-01-10")
    page.end_date_var = Var("2024-01-01")
    page.run_button = FakeWidget()

    page.run_backtest()
    assert infos[-1] == ("Invalid dates", "Start date must be before end date.")

    assert page._selected_signal_names({"x": Var(False), "y": Var(True)}) == ["y"]

    run_dir = tmp_path / "run"
    run_dir.mkdir()
    (run_dir / "aggregate_metrics.json").write_text('{"Sharpe": 1.2}')
    page.logs_text = FakeText()
    page._register_run_dirs = lambda _dirs: None
    page._set_tree_data = lambda *_args, **_kwargs: None
    page._refresh_artifacts_view = lambda: None
    page._render_single_run = lambda _run: None
    page.leaderboard_tree = object()

    page._consume_run_outputs(f"Saved outputs to: {run_dir}")
    assert "Saved outputs to:" in page.logs_text.value


def test_general_analysis_success_persists_and_records_artifact(monkeypatch, tmp_path):
    infos = []
    monkeypatch.setattr(general_analysis_page.messagebox, "showinfo", lambda title, msg: infos.append((title, msg)))

    controller = FakeController(AppState(tickers=["AAPL", "MSFT"]))
    controller.api_key = "k"

    page = general_analysis_page.GeneralAnalysisPage.__new__(general_analysis_page.GeneralAnalysisPage)
    page.controller = controller
    page.api_client = object()
    page.analysis_type_var = Var("Cross-Sectional")
    page.lookback_var = Var("30")
    page.skip_var = Var("5")
    page.top_quantile_var = Var("0.2")
    page.bottom_quantile_var = Var("0.2")
    page.output_dir_var = Var(str(tmp_path))
    page.momentum_volatility_var = Var(False)
    page.momentum_residual_var = Var(False)
    page.momentum_multi_horizon_var = Var(False)

    page._selected_strategy_key = lambda: "Momentum"
    page._collect_price_history = lambda *_a: ({"AAPL": []}, [])
    page._collect_fundamentals = lambda *_a: ({"AAPL": {}}, [])
    page._run_cross_sectional_analysis = lambda *_a, **_k: "Sharpe: 1.1\nCAGR: 10%"
    page._run_time_series_analysis = lambda *_a, **_k: "unused"
    records = []
    page._record_run = lambda **kwargs: records.append(kwargs)
    out = tmp_path / "report.txt"
    page._write_report = lambda *_a, **_k: out

    page.run_analysis()

    assert controller.state.general_analysis_settings["lookback_days"] == 30
    assert records and records[-1]["status"] == "success"
    assert records[-1]["artifact_path"] == str(out)
    assert infos and infos[-1][0] == "Analysis complete"


def test_missing_ticker_guards_across_option_pages(monkeypatch):
    infos = []
    monkeypatch.setattr(analysis_page.messagebox, "showinfo", lambda title, msg: infos.append((title, msg)))
    monkeypatch.setattr(call_put_analysis_page.messagebox, "showinfo", lambda title, msg: infos.append((title, msg)))
    monkeypatch.setattr(spread_analysis_page.messagebox, "showinfo", lambda title, msg: infos.append((title, msg)))

    controller = FakeController(AppState(selected_ticker=None))

    for module, cls_name in [
        (analysis_page, "AnalysisPage"),
        (call_put_analysis_page, "CallPutAnalysisPage"),
        (spread_analysis_page, "SpreadAnalysisPage"),
    ]:
        page_cls = getattr(module, cls_name)
        page = page_cls.__new__(page_cls)
        page.controller = controller
        page.api_client = object()
        page.load_market_data()

    assert sum(1 for title, _ in infos if title == "Missing ticker") == 3


def test_main_menu_and_research_navigation(monkeypatch):
    infos = []
    saved = []
    monkeypatch.setattr(main_menu.messagebox, "showinfo", lambda title, msg: infos.append((title, msg)))
    monkeypatch.setattr(main_menu, "save_api_key", lambda key: saved.append(key))

    controller = FakeController(AppState(), api_key="")
    menu = main_menu.MainMenu.__new__(main_menu.MainMenu)
    menu.controller = controller
    menu.api_key_var = Var(" abc ")
    menu.save_api_key()
    assert saved == ["abc"]
    assert controller.api_key == "abc"

    lab = research_lab_page.ResearchLabPage.__new__(research_lab_page.ResearchLabPage)
    lab.controller = controller
    lab.open_governance_workspace()
    assert controller.frames[-1] == "BacktestingPage"


def test_research_lab_validation_no_signal_and_invalid_dates():
    lab = research_lab_page.ResearchLabPage.__new__(research_lab_page.ResearchLabPage)
    lab.wizard_data_universe_var = Var("AAPL")
    lab.wizard_period_start_var = Var("2024-02-01")
    lab.wizard_period_end_var = Var("2024-01-01")
    ok, msg = lab._wizard_validate_step(1)
    assert not ok
    assert "before" in msg

    messages = []
    research_lab_page.messagebox.showinfo = lambda title, msg: messages.append((title, msg))
    assert lab._parse_signal_csv("", valid_options=("ts_momentum",), field_name="Entry signals") is None
    assert messages[-1][0] == "Invalid input"


def test_navigation_smoke_all_pages_non_interactive(monkeypatch):
    monkeypatch.setattr(main_menu.messagebox, "showinfo", lambda *_: None)
    monkeypatch.setattr(ticker_entry_page.messagebox, "showinfo", lambda *_: None)
    monkeypatch.setattr(analysis_page, "load_api_key", lambda: "")
    monkeypatch.setattr(analysis_page.messagebox, "showinfo", lambda *_: None)
    monkeypatch.setattr(call_put_analysis_page.messagebox, "showinfo", lambda *_: None)
    monkeypatch.setattr(spread_analysis_page.messagebox, "showinfo", lambda *_: None)
    controller = FakeController(AppState(tickers=["AAPL"], selected_ticker="AAPL"), api_key="key")

    menu = main_menu.MainMenu.__new__(main_menu.MainMenu)
    menu.controller = controller
    menu.api_key_var = Var("key")
    menu.refresh = lambda: menu.api_key_var.set(menu.controller.api_key)
    menu.refresh()

    entry = ticker_entry_page.TickerEntryPage.__new__(ticker_entry_page.TickerEntryPage)
    entry.controller = controller
    entry.text_box = FakeText("AAPL")
    entry.save_tickers()

    select = ticker_select_page.TickerSelectPage.__new__(ticker_select_page.TickerSelectPage)
    select.controller = controller
    select.ticker_list = FakeListbox()
    select.refresh()

    analysis = analysis_page.AnalysisPage.__new__(analysis_page.AnalysisPage)
    analysis.controller = controller
    analysis.analysis_mode_var = Var("Option Analysis")
    analysis.strategy_var = Var("Naked Call")
    analysis._toggle_info_panels = lambda: None
    analysis.scroll_canvas = type("C", (), {"configure": lambda *a, **k: None, "bbox": lambda *a, **k: (0, 0, 1, 1)})()
    analysis.after = lambda *_a, **_k: None
    analysis.refresh()

    for module, cls_name in [
        (call_put_analysis_page, "CallPutAnalysisPage"),
        (spread_analysis_page, "SpreadAnalysisPage"),
    ]:
        pcls = getattr(module, cls_name)
        page = pcls.__new__(pcls)
        page.controller = controller
        page.api_client = None
        page.load_market_data()


def test_research_lab_funnel_kpi_metrics_expand_beyond_acceptance_rate(tmp_path):
    lab = research_lab_page.ResearchLabPage.__new__(research_lab_page.ResearchLabPage)
    lab._research_lab_dir = tmp_path

    events = [
        {
            "hypothesis_id": "h1",
            "date": "2024-01-03",
            "submitted_at": "2024-01-01",
            "decision_at": "2024-01-03",
            "strategy_family": "trend",
            "decision": "accept",
            "promotion_state": "promoted_to_experiment",
        },
        {
            "hypothesis_id": "h1",
            "date": "2024-02-01",
            "submitted_at": "2024-01-01",
            "decision_at": "2024-02-01",
            "strategy_family": "trend",
            "decision": "reject",
            "promotion_state": "rejected",
        },
        {
            "hypothesis_id": "h2",
            "date": "2024-01-11",
            "submitted_at": "2024-01-10",
            "decision_at": "2024-01-11",
            "strategy_family": "mean_reversion",
            "decision": "accept",
            "promotion_state": "promoted_to_experiment",
        },
        {
            "hypothesis_id": "h3",
            "date": "2024-01-19",
            "submitted_at": "2024-01-15",
            "decision_at": "2024-01-19",
            "strategy_family": "trend",
            "decision": "reject",
            "promotion_state": "rejected",
        },
    ]
    event_path = tmp_path / "idea_funnel_events.jsonl"
    event_path.write_text("\n".join(json.dumps(event) for event in events), encoding="utf-8")

    metrics = lab._compute_funnel_metrics()

    assert metrics["acceptance_rate_pct"] == pytest.approx(50.0)
    assert metrics["median_time_to_decision_days"] == pytest.approx(3.0)
    assert metrics["false_positive_rate_pct"] == pytest.approx(50.0)

    strategy_rates = {row["strategy_family"]: row for row in metrics["pass_rates_by_strategy_family"]}
    assert strategy_rates["trend"]["acceptance_rate_pct"] == pytest.approx(100 / 3)
    assert strategy_rates["mean_reversion"]["acceptance_rate_pct"] == pytest.approx(100.0)

    month_conversion = {row["month"]: row for row in metrics["promotion_conversion_by_month"]}
    assert month_conversion["2024-01"]["promotion_conversion_pct"] == pytest.approx(200 / 3)
    assert month_conversion["2024-02"]["promotion_conversion_pct"] == pytest.approx(0.0)

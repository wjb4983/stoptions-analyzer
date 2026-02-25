from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path

import pytest

from state import AppState
from ui import analysis_page, backtesting_page, call_put_analysis_page, create_regime_page, general_analysis_page, main_menu, research_lab_page, spread_analysis_page, ticker_entry_page, ticker_select_page
from backtesting.regime_training_pipeline import RegimeTrainingResult


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
        self.value += _args[-1]

    def see(self, *_args):
        return None


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

    def configure(self, **kwargs):
        self.config(**kwargs)


class FakeFrameWidget(FakeWidget):
    def __init__(self):
        super().__init__()
        self.grid_calls = 0
        self.grid_remove_calls = 0

    def grid(self):
        self.grid_calls += 1

    def grid_remove(self):
        self.grid_remove_calls += 1


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

    menu.open_create_regime_workspace()
    assert controller.frames[-1] == "CreateRegimePage"

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


def test_research_lab_governance_dashboard_renders_counts_missing_and_statuses():
    lab = research_lab_page.ResearchLabPage.__new__(research_lab_page.ResearchLabPage)
    lab._governance_gate_counts_var = Var()
    lab._governance_missing_checks_var = Var()
    lab._governance_promotion_ready_var = Var()
    lab._governance_approval_status_var = Var()

    lab._load_governance_payload_from_output = lambda _output: {
        "gate_checks": {
            "deflated_sharpe_reality_check": True,
            "parameter_stability_penalty": False,
            "train_validation_test_drift": True,
        },
        "missing_required_checks": ["capacity", "audit_log"],
        "is_promotion_ready": False,
        "approval_status": "in_review",
        "promotion_state": "paper",
    }

    lab._refresh_governance_dashboard_from_output("Saved outputs to: /tmp/fake")

    assert lab._governance_gate_counts_var.get() == "2 passed / 1 failed (3 total)"
    assert lab._governance_missing_checks_var.get() == "capacity, audit_log"
    assert lab._governance_promotion_ready_var.get() == "Not ready"
    assert lab._governance_approval_status_var.get() == "in_review (paper)"

    lab._load_governance_payload_from_output = lambda _output: {
        "gate_checks": {"deflated_sharpe_reality_check": True},
        "missing_required_checks": [],
        "is_promotion_ready": True,
        "approval_status": "approved",
        "promotion_state": "production",
    }
    lab._refresh_governance_dashboard_from_output("Saved outputs to: /tmp/fake")
    assert lab._governance_gate_counts_var.get() == "1 passed / 0 failed (1 total)"
    assert lab._governance_missing_checks_var.get() == "None"
    assert lab._governance_promotion_ready_var.get() == "Ready"
    assert lab._governance_approval_status_var.get() == "approved (production)"


def test_research_wizard_step_validation_transitions_for_universe_dates_and_signals(monkeypatch):
    lab = research_lab_page.ResearchLabPage.__new__(research_lab_page.ResearchLabPage)
    lab.wizard_data_universe_var = Var("")
    lab.wizard_period_start_var = Var("2024-01-10")
    lab.wizard_period_end_var = Var("2024-02-01")

    ok, msg = lab._wizard_validate_step(1)
    assert not ok and "at least one ticker" in msg

    lab.wizard_data_universe_var.set("AAPL, MSFT")
    lab.wizard_period_start_var.set("bad-date")
    ok, msg = lab._wizard_validate_step(1)
    assert not ok and "valid start and end dates" in msg

    lab.wizard_period_start_var.set("2024-02-05")
    lab.wizard_period_end_var.set("2024-02-01")
    ok, msg = lab._wizard_validate_step(1)
    assert not ok and "before" in msg

    lab.wizard_period_start_var.set("2024-01-01")
    lab.wizard_period_end_var.set("2024-02-01")
    ok, msg = lab._wizard_validate_step(1)
    assert ok and msg == ""

    popups = []
    monkeypatch.setattr(research_lab_page.messagebox, "showinfo", lambda title, msg: popups.append((title, msg)))
    assert lab._parse_signal_csv("", valid_options=("ts_momentum",), field_name="Entry signals") is None
    assert "must include at least one signal" in popups[-1][1]
    assert lab._parse_signal_csv("ts_momentum,unknown", valid_options=("ts_momentum",), field_name="Entry signals") is None
    assert "Unsupported entry signals" in popups[-1][1]
    assert lab._parse_signal_csv("ts_momentum", valid_options=("ts_momentum",), field_name="Entry signals") == ["ts_momentum"]


def test_backtesting_mode_toggle_and_preset_selection_apply_expected_settings():
    controller = FakeController(AppState())
    page = backtesting_page.BacktestingPage.__new__(backtesting_page.BacktestingPage)
    page.controller = controller
    page.ui_mode_var = Var("basic")
    page.use_walk_forward_var = Var(True)
    page.show_advanced_controls_var = Var(False)
    page.strategy_var = Var("momentum")
    page._update_validation_hint = lambda: False
    frame = FakeFrameWidget()
    adv_widget = FakeWidget()
    page._advanced_widgets = {"walk_forward_frame": frame, "use_optimizer": adv_widget}

    page._on_mode_changed()
    assert adv_widget.config_calls[-1]["state"] == "disabled"
    assert frame.grid_remove_calls == 1
    assert page.use_walk_forward_var.get() is False

    page.ui_mode_var.set("advanced")
    page._on_mode_changed()
    assert adv_widget.config_calls[-1]["state"] == "normal"
    assert frame.grid_calls == 1

    page.preset_var = Var("Intraday Momentum (intraday_momentum)")
    page._preset_display_to_key = {
        "Custom": "custom",
        "Intraday Momentum (intraday_momentum)": "intraday_momentum",
    }
    page._preset_key_to_display = {"custom": "Custom", "intraday_momentum": "Intraday Momentum (intraday_momentum)"}
    applied = []
    page._apply_settings = lambda settings: applied.append(settings)

    page._on_preset_selected()

    expected = backtesting_page.BACKTEST_STRATEGY_PRESETS["intraday_momentum"]["settings"]
    assert applied[-1] == expected
    assert controller.state.backtest_settings["selected_preset"] == "intraday_momentum"
    assert controller.persist_count == 1


def test_output_and_selected_task_logs_refresh_deterministically(tmp_path):
    run_dir = tmp_path / "run_a"
    run_dir.mkdir()
    (run_dir / "aggregate_metrics.json").write_text('{"Sharpe": 1.23}', encoding="utf-8")

    bt_controller = FakeController(AppState())
    bt_page = backtesting_page.BacktestingPage.__new__(backtesting_page.BacktestingPage)
    bt_page.controller = bt_controller
    bt_page.logs_text = FakeText("stale")
    bt_page._register_run_dirs = lambda _dirs: None
    bt_page._refresh_artifacts_view = lambda: None
    bt_page._render_single_run = lambda _run: None
    bt_page._set_tree_data = lambda *_args, **_kwargs: None
    bt_page.leaderboard_tree = object()

    bt_page._consume_run_outputs(f"run started\nSaved outputs to: {run_dir}")
    assert bt_page.logs_text.get("1.0", "end") == f"run started\nSaved outputs to: {run_dir}"

    bt_page._consume_run_outputs("no artifacts yet")
    assert bt_page.logs_text.get("1.0", "end") == "no artifacts yet"

    lab = research_lab_page.ResearchLabPage.__new__(research_lab_page.ResearchLabPage)
    lab.task_logs_text = FakeText()
    task = research_lab_page.ResearchTask(task_id="t1", label="Task", target=lambda *_a: "", context={}, config={})
    task.logs = ["first", "second"]
    lab._task_queue = [task]
    lab._selected_task = lambda: task

    lab._refresh_selected_task_logs()
    assert lab.task_logs_text.get("1.0", "end") == "first\nsecond\n"

    task.logs.append("third")
    lab._refresh_selected_task_logs()
    assert lab.task_logs_text.get("1.0", "end") == "first\nsecond\nthird\n"


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


def _build_wizard_lab(tmp_path: Path) -> research_lab_page.ResearchLabPage:
    lab = research_lab_page.ResearchLabPage.__new__(research_lab_page.ResearchLabPage)
    lab._research_lab_dir = tmp_path
    lab._wizard_state_path = tmp_path / "wizard_state.json"
    lab._wizard_steps = ["idea", "universe", "plan", "run", "review"]
    lab._wizard_step_index = 0
    lab._wizard_comments = []
    lab._wizard_history = []
    lab._append_output = lambda *_a, **_k: None
    lab._wizard_render_comments_log = lambda: None
    lab._wizard_refresh_nav_state = lambda: None
    lab.wizard_idea_name_var = Var("Idea")
    lab.wizard_idea_thesis_var = Var("Thesis")
    lab.wizard_idea_owner_var = Var("alice")
    lab.wizard_data_universe_var = Var("AAPL")
    lab.wizard_period_start_var = Var("2024-01-01")
    lab.wizard_period_end_var = Var("2024-06-01")
    lab.wizard_sector_include_var = Var("")
    lab.wizard_sector_exclude_var = Var("")
    lab.wizard_adv_threshold_var = Var("")
    lab.wizard_liquidity_threshold_var = Var("")
    lab.wizard_price_min_var = Var("")
    lab.wizard_price_max_var = Var("")
    lab.wizard_market_cap_min_var = Var("")
    lab.wizard_market_cap_max_var = Var("")
    lab.wizard_min_option_oi_var = Var("")
    lab.wizard_min_option_volume_var = Var("")
    lab.wizard_min_option_dte_var = Var("")
    lab.wizard_require_weeklies_var = Var(False)
    lab.wizard_test_plan_var = Var("walk_forward")
    lab.wizard_acceptance_var = Var("Sharpe > 1")
    lab.wizard_run_validation_var = Var(True)
    lab.wizard_run_optimization_var = Var(False)
    lab.wizard_run_stress_var = Var(False)
    lab.wizard_review_notes_var = Var("")
    lab.wizard_promotion_decision_var = Var("pending")
    lab.wizard_session_label_var = Var("team sync")
    lab.show_advanced_controls_var = Var(False)
    lab.easy_mode_var = Var(True)
    lab._advanced_workflow_widgets = []
    lab._refresh_workflow_validation_hints = lambda: None
    lab._on_show_advanced_controls_toggle = lambda: None
    return lab


def test_wizard_state_document_roundtrip_and_upgrade(tmp_path):
    lab = _build_wizard_lab(tmp_path)
    lab._wizard_comments = [{"owner": "alice", "note": "n1", "timestamp": "2024-01-01T00:00:00"}]
    payload = lab._wizard_state_payload()
    doc = lab._wizard_state_document(payload)
    parsed, upgraded = lab._wizard_extract_payload(doc)
    assert parsed is not None
    assert not upgraded
    assert parsed["wizard_comments"][0]["owner"] == "alice"

    legacy_payload, legacy_upgraded = lab._wizard_extract_payload(payload)
    assert legacy_payload is not None
    assert legacy_upgraded


def test_wizard_export_import_session_merges_comments_history(monkeypatch, tmp_path):
    lab = _build_wizard_lab(tmp_path)
    infos = []
    monkeypatch.setattr(research_lab_page.messagebox, "showinfo", lambda title, msg: infos.append((title, msg)))

    lab._wizard_comments = [{"owner": "alice", "note": "start", "timestamp": "2024-01-01T00:00:00"}]
    lab._wizard_history = [{"event": "created", "timestamp": "2024-01-01T00:00:00"}]
    lab._wizard_export_session()

    session_files = sorted((tmp_path / "sessions").glob("*.json"))
    assert session_files

    lab._wizard_comments = [{"owner": "bob", "note": "local", "timestamp": "2024-01-02T00:00:00"}]
    lab._wizard_history = [{"event": "local", "timestamp": "2024-01-02T00:00:00"}]

    monkeypatch.setattr(research_lab_page.filedialog, "askopenfilename", lambda **_k: str(session_files[0]))
    lab._wizard_import_session()

    owners = {row["owner"] for row in lab._wizard_comments}
    events = {row["event"] for row in lab._wizard_history}
    assert owners == {"alice", "bob"}
    assert "created" in events and "local" in events
    assert any(title == "Session exported" for title, _ in infos)


def _build_create_regime_logic_page() -> create_regime_page.CreateRegimePage:
    page = create_regime_page.CreateRegimePage.__new__(create_regime_page.CreateRegimePage)
    page.regime_legs = [
        create_regime_page.CreateRegimePage._build_default_leg(page, "Trend Following")
    ]
    page.selected_leg_index = 0
    page.leg_control_vars = {}
    page.validation_badge_vars = {}
    page.validation_badges = {}
    page._is_training = False
    return page


def _build_create_regime_training_page() -> tuple[create_regime_page.CreateRegimePage, FakeController]:
    controller = FakeController(
        AppState(
            regime_definitions={
                "baseline": {
                    "label": "Baseline",
                    "global_risk_limits": {"max_drawdown_stop": 0.14, "max_position_pct": 0.09},
                    "training_window": {"lookback_days": 190, "retrain_frequency_days": 14},
                    "confidence_thresholds": {"model_confidence_min": 0.71},
                }
            },
            active_regime_id="baseline",
        )
    )
    page = _build_create_regime_logic_page()
    page.controller = controller
    page.train_button = FakeWidget()
    page.export_button = FakeWidget()
    page._can_train_export = lambda: (True, "")
    page._update_validation_and_actions = lambda: None
    page.after = lambda _delay, callback: callback()
    return page, controller


def test_create_regime_leg_type_switching_updates_controls():
    page = _build_create_regime_logic_page()

    leg = page._selected_leg()
    assert leg["model_type"] == "Trend Following"
    assert float(leg["controls"]["lookback_days"]) == 90

    page._apply_leg_type("Mean Reversion")
    leg = page._selected_leg()
    assert leg["model_type"] == "Mean Reversion"
    assert float(leg["controls"]["lookback_days"]) == 30
    assert float(leg["controls"]["entry_zscore"]) == 2.0

    page._apply_leg_type("Regime Change")
    leg = page._selected_leg()
    assert leg["model_type"] == "Regime Change"
    assert float(leg["controls"]["lookback_days"]) == 60
    assert float(leg["controls"]["detection_threshold"]) == 1.6




def test_create_regime_change_selection_initializes_expected_controls_defaults():
    page = _build_create_regime_logic_page()

    page._apply_leg_type("Regime Change")
    leg = page._selected_leg()

    assert leg["model_type"] == "Regime Change"
    assert set(leg["controls"]) == {
        "lookback_days",
        "detection_threshold",
        "max_position_pct",
        "max_drawdown_stop",
        "turnover_limit",
        "slippage_bps",
        "model_confidence_min",
        "regime_stability_min",
    }
    assert float(leg["controls"]["lookback_days"]) == 60
    assert float(leg["controls"]["detection_threshold"]) == 1.6
    assert float(leg["controls"]["max_position_pct"]) == 0.05
    assert float(leg["controls"]["max_drawdown_stop"]) == 0.08

def test_create_regime_unknown_leg_mapping_blocks_training() -> None:
    page = _build_create_regime_logic_page()
    page.regime_legs[0]["model_type"] = "Not A Leg"

    ok, message = page._can_train_export()
    assert not ok
    assert "Unsupported UI leg" in message


def test_create_regime_invalid_knob_combinations_are_blocked():
    page = _build_create_regime_logic_page()
    leg = page._selected_leg()
    leg["controls"]["entry_zscore"] = 0.8
    leg["controls"]["model_confidence_min"] = 0.9

    ok, message = page._can_train_export()
    assert not ok
    assert "Entry z-score" in message


def test_create_regime_change_invalid_knob_combinations_are_blocked():
    page = _build_create_regime_logic_page()
    page._apply_leg_type("Regime Change")
    leg = page._selected_leg()
    leg["controls"]["detection_threshold"] = 0.7

    ok, message = page._can_train_export()
    assert not ok
    assert "Detection threshold is too permissive" in message


@pytest.mark.parametrize(
    "overrides, expected",
    [
        ({"lookback_days": 120, "entry_zscore": 1.4, "turnover_limit": 0.3, "slippage_bps": 8}, True),
        ({"lookback_days": 20, "entry_zscore": 3.2, "turnover_limit": 1.2, "slippage_bps": 45}, False),
    ],
)
def test_create_regime_train_export_buttons_follow_validation(overrides, expected):
    page = _build_create_regime_logic_page()
    page.train_button = FakeWidget()
    page.export_button = FakeWidget()
    page.validation_message_var = Var()
    page.risk_summary_var = Var()

    class _FakeProsCons:
        def configure(self, **_kwargs):
            return None

        def delete(self, *_args):
            return None

        def insert(self, *_args):
            return None

    page.pros_cons_text = _FakeProsCons()

    leg = page._selected_leg()
    for key, value in overrides.items():
        leg["controls"][key] = value

    page._update_validation_and_actions()

    expected_state = "normal" if expected else "disabled"
    assert page.train_button.config_calls[-1]["state"] == expected_state
    assert page.export_button.config_calls[-1]["state"] == expected_state


def test_create_regime_run_train_calls_pipeline_with_translated_legs(monkeypatch):
    infos = []
    monkeypatch.setattr(create_regime_page.messagebox, "showinfo", lambda title, msg: infos.append((title, msg)))

    page, _controller = _build_create_regime_training_page()
    monkeypatch.setattr(create_regime_page.messagebox, "showinfo", lambda *_args, **_kwargs: None)
    page._append_structured_log = lambda **_kwargs: None

    page.regime_legs = [
        {
            "name": "Trend Following leg",
            "model_type": "Trend Following",
            "controls": {
                "lookback_days": 120,
                "entry_zscore": 1.3,
                "model_confidence_min": 0.68,
                "max_position_pct": 0.07,
                "max_drawdown_stop": 0.11,
            },
        },
        {
            "name": "Regime Change leg",
            "model_type": "Regime Change",
            "controls": {
                "lookback_days": 70,
                "detection_threshold": 1.9,
                "max_position_pct": 0.03,
                "max_drawdown_stop": 0.06,
            },
        },
    ]

    calls = []

    def _fake_run(request):
        calls.append(request)
        return RegimeTrainingResult(
            run_id="run-1",
            status="success",
            metrics={"sharpe": 1.0},
            artifact_paths={"manifest": "manifest.json"},
            timestamps={"started_at": "2024-01-01T00:00:00+00:00", "completed_at": "2024-01-01T00:01:00+00:00"},
            summary="ok",
        )

    class _ImmediateThread:
        def __init__(self, *, target, daemon):
            self._target = target

        def start(self):
            self._target()

    monkeypatch.setattr(create_regime_page, "execute_regime_training_pipeline", _fake_run)
    monkeypatch.setattr(create_regime_page.threading, "Thread", _ImmediateThread)

    page._run_train()

    assert len(calls) == 1
    request = calls[0]
    assert [leg.model_type for leg in request.legs] == ["timeseries_momentum", "regime_change_detection"]
    assert request.legs[0].controls["sizing_cap"] == pytest.approx(0.07)
    assert request.legs[0].controls["stop_loss_pct"] == pytest.approx(0.11)
    assert request.legs[1].controls["detection_threshold"] == pytest.approx(1.9)
    assert request.training_data_settings["required_history_years"] == 5
    assert isinstance(request.training_data_settings["scenario_settings"], list)
    assert request.training_data_settings["cache_audit_strict"] is True
    assert request.training_data_settings["allow_synthetic_fallback"] is False
    assert request.training_data_settings["min_symbol_coverage_ratio"] == 1.0
    assert request.training_data_settings["min_bars_per_year"] == 1


def test_create_regime_successful_training_appends_run_and_persists(monkeypatch):
    infos = []
    monkeypatch.setattr(create_regime_page.messagebox, "showinfo", lambda title, msg: infos.append((title, msg)))

    page, controller = _build_create_regime_training_page()
    logs = []
    page._append_structured_log = lambda **kwargs: logs.append(kwargs)

    result = RegimeTrainingResult(
        run_id="run-ok",
        status="success",
        metrics={"sharpe": 1.2},
        artifact_paths={"manifest": "data/manifest.json"},
        timestamps={"started_at": "2024-01-01T00:00:00+00:00", "completed_at": "2024-01-01T00:01:00+00:00"},
        summary="all good",
        metadata={"regime": "baseline"},
        logs=("line-1",),
    )

    page._on_training_succeeded(result)

    assert len(controller.state.regime_training_runs) == 1
    assert controller.state.regime_training_runs[0]["run_id"] == "run-ok"
    assert controller.persist_count == 1
    assert infos[-1] == ("Training completed", "all good")
    assert any(row["event"] == "training_completed" for row in logs)


def test_create_regime_run_train_failure_shows_error_and_no_success_log(monkeypatch):
    errors = []
    monkeypatch.setattr(create_regime_page.messagebox, "showerror", lambda title, msg: errors.append((title, msg)))

    page, controller = _build_create_regime_training_page()
    structured_logs = []
    page._append_structured_log = lambda **kwargs: structured_logs.append(kwargs)

    class _ImmediateThread:
        def __init__(self, *, target, daemon):
            self._target = target

        def start(self):
            self._target()

    monkeypatch.setattr(create_regime_page, "execute_regime_training_pipeline", lambda _request: (_ for _ in ()).throw(RuntimeError("runner exploded")))
    monkeypatch.setattr(create_regime_page.threading, "Thread", _ImmediateThread)

    page._run_train()

    assert errors and errors[-1][0] == "Training failed"
    assert "runner exploded" in errors[-1][1]
    assert controller.state.regime_training_runs == []
    assert not any(entry["event"] == "training_completed" for entry in structured_logs)
    assert any(entry["event"] == "training_failed" for entry in structured_logs)


def test_create_regime_run_export_builds_bundle_from_latest_manifest(monkeypatch):
    infos = []
    monkeypatch.setattr(create_regime_page.messagebox, "showinfo", lambda title, msg: infos.append((title, msg)))

    page, controller = _build_create_regime_training_page()
    logs = []
    page._append_structured_log = lambda **kwargs: logs.append(kwargs)
    controller.state.regime_training_runs = [
        {
            "run_id": "run-1",
            "status": "success",
            "artifact_path": "data/run_1/manifest.json",
            "summary": "trained",
        }
    ]

    captured = []

    class _FakeBundle:
        bundle_dir = "data/regime_exports/regime_export_run-1-abc"
        bundle_manifest_path = "data/regime_exports/regime_export_run-1-abc/bundle_manifest.json"
        deployment_version = "run-1-abc"

    def _fake_export(path):
        captured.append(path)
        return _FakeBundle()

    monkeypatch.setattr(create_regime_page, "export_regime_training_bundle", _fake_export)

    page._run_export()

    assert captured == ["data/run_1/manifest.json"]
    assert infos and infos[-1][0] == "Export completed"
    assert "Bundle path: data/regime_exports/regime_export_run-1-abc" in infos[-1][1]
    assert any(entry.get("event") == "export_completed" for entry in logs)


def test_prepare_regime_training_cache_selective_backfill_symbols(monkeypatch, tmp_path):
    page, controller = _build_create_regime_training_page()
    controller.api_key = "demo"

    controller.state.tickers = ["AAPL", "MSFT"]
    request = page._build_regime_training_request()
    request.training_data_settings["cache_root"] = str(tmp_path)
    request.training_data_settings["required_history_years"] = 3

    captured = {}

    monkeypatch.setattr(
        create_regime_page,
        "audit_universe_history",
        lambda **_kwargs: {
            "failing_symbols": ["MSFT"],
            "failing_details": {"MSFT": {"missing_year_gaps": [2024], "corrupted_years": []}},
            "pass": False,
        },
    )

    def _fake_backfill(*, tickers, start_date, end_date, cache_root, api_key):
        captured["tickers"] = list(tickers)
        captured["start_date"] = start_date
        captured["end_date"] = end_date
        captured["cache_root"] = cache_root
        captured["api_key"] = api_key

    monkeypatch.setattr(create_regime_page, "run_backtest_cache", _fake_backfill)

    report = page._prepare_regime_training_cache(request)

    assert captured["tickers"] == ["MSFT"]
    assert report["failing_symbols"] == ["MSFT"]
    assert Path(report["artifact_path"]).exists()


def test_create_regime_run_train_attaches_cache_audit_metadata(monkeypatch, tmp_path):
    page, _controller = _build_create_regime_training_page()
    monkeypatch.setattr(create_regime_page.messagebox, "showinfo", lambda *_args, **_kwargs: None)
    page._append_structured_log = lambda **_kwargs: None

    class _ImmediateThread:
        def __init__(self, *, target, daemon):
            self._target = target

        def start(self):
            self._target()

    monkeypatch.setattr(create_regime_page.threading, "Thread", _ImmediateThread)
    monkeypatch.setattr(
        create_regime_page.CreateRegimePage,
        "_prepare_regime_training_cache",
        lambda _self, _request: {"note": "ok", "failing_symbols": [], "artifact_path": str(tmp_path / "audit.json")},
    )

    calls = []

    def _fake_run(request):
        calls.append(dict(request.training_data_settings))
        return RegimeTrainingResult(
            run_id="run-1",
            status="success",
            metrics={"sharpe": 1.0},
            artifact_paths={"manifest": "manifest.json"},
            timestamps={"started_at": "2024-01-01T00:00:00+00:00", "completed_at": "2024-01-01T00:01:00+00:00"},
            summary="ok",
            metadata={},
        )

    monkeypatch.setattr(create_regime_page, "execute_regime_training_pipeline", _fake_run)

    page._run_train()

    assert calls
    assert calls[0]["cache_audit_report"].endswith("audit.json")


def test_create_regime_training_failure_shows_hard_error_for_insufficient_real_history(monkeypatch):
    errors = []
    monkeypatch.setattr(create_regime_page.messagebox, "showerror", lambda title, msg: errors.append((title, msg)))

    page, _controller = _build_create_regime_training_page()
    page._append_structured_log = lambda **_kwargs: None

    result = RegimeTrainingResult(
        run_id="run-bad",
        status="failed",
        metrics={},
        artifact_paths={},
        timestamps={"started_at": "2024-01-01T00:00:00+00:00", "completed_at": "2024-01-01T00:01:00+00:00"},
        errors=("INSUFFICIENT_REAL_HISTORY: coverage_diagnostics={}",),
        summary="failed",
    )

    page._on_training_result_failed(result)

    assert errors
    assert errors[-1][0] == "Training blocked: insufficient real history"
    assert "Synthetic fallback is disabled" in errors[-1][1]

class MultiSelectListbox:
    def __init__(self, items: list[str]):
        self.items = list(items)
        self._selection: set[int] = set()
        self.state = "normal"

    def get(self, idx: int):
        return self.items[idx]

    def size(self):
        return len(self.items)

    def curselection(self):
        return tuple(sorted(self._selection))

    def selection_set(self, idx: int):
        self._selection.add(idx)

    def selection_clear(self, *_args):
        self._selection.clear()

    def configure(self, **kwargs):
        if "state" in kwargs:
            self.state = kwargs["state"]


class FakeCombo:
    def __init__(self):
        self.values = []

    def configure(self, **kwargs):
        if "values" in kwargs:
            self.values = list(kwargs["values"])

    def cget(self, key: str):
        if key == "values":
            return tuple(self.values)
        raise KeyError(key)


def test_backtesting_trained_regime_dropdown_population_and_hydration(tmp_path):
    training_manifest = tmp_path / "run" / "manifest.json"
    training_manifest.parent.mkdir(parents=True, exist_ok=True)
    training_manifest.write_text(
        json.dumps(
            {
                "run_id": "run-ui",
                "request": {
                    "regime_name": "UI Risk",
                    "model_choice": "auto_model_search",
                    "training_window": {"lookback_days": 77, "retrain_frequency_days": 4},
                    "risk_limits": {"max_gross_exposure": 1.2, "max_net_exposure": 0.4},
                    "training_data_settings": {"scenario_settings": [{"name": "panic_crash"}]},
                },
            }
        ),
        encoding="utf-8",
    )

    page = backtesting_page.BacktestingPage.__new__(backtesting_page.BacktestingPage)
    page.controller = FakeController(AppState(regime_training_runs=[{"run_id": "run-ui", "artifact_path": str(training_manifest), "summary": "UI Risk"}]))
    page.trained_regime_combo = FakeCombo()
    page.trained_regime_var = Var("")
    page.strategy_var = Var("momentum")
    page.lookback_days_var = Var("90")
    page.skip_days_var = Var("5")
    page.portfolio_max_gross_var = Var("1.0")
    page.portfolio_min_net_var = Var("-1.0")
    page.portfolio_max_net_var = Var("1.0")
    page.portfolio_max_symbol_var = Var("0.25")
    page.portfolio_max_sector_var = Var("0.60")
    page.gov_min_stability_var = Var("0.55")
    page.gov_expected_signal_agreement_var = Var("1.0")
    page.gov_max_signal_agreement_drift_var = Var("0.10")
    page.stress_enable_historical_replay_var = Var(False)
    page.scenario_pack_listbox = MultiSelectListbox(["panic_crash", "bull_low_vol"])
    page._scenario_pack_options = ("panic_crash", "bull_low_vol")
    page.regime_provenance_var = Var("")
    page.regime_diff_var = Var("")
    page.regime_lock_var = Var(True)
    page._regime_locked_fields = True
    page._active_regime_contract = None
    page._regime_loaded_values = {}
    page._regime_loading_defaults = False
    page.strategy_combo = FakeWidget()
    page.lookback_days_entry = FakeWidget()
    page.skip_days_entry = FakeWidget()
    page.portfolio_max_gross_entry = FakeWidget()
    page.portfolio_min_net_entry = FakeWidget()
    page.portfolio_max_net_entry = FakeWidget()
    page.portfolio_max_symbol_entry = FakeWidget()
    page.portfolio_max_sector_entry = FakeWidget()
    page.gov_min_stability_entry = FakeWidget()
    page.gov_expected_signal_agreement_entry = FakeWidget()
    page.gov_max_signal_agreement_drift_entry = FakeWidget()

    page._refresh_trained_regime_choices()
    assert any("training run" in label for label in page.trained_regime_combo.values)

    selected_label = next(label for label in page.trained_regime_combo.values if "training run" in label)
    page.trained_regime_var.set(selected_label)
    page._on_trained_regime_selected()

    assert page.lookback_days_var.get() == "77"
    assert page.skip_days_var.get() == "4"
    assert page.portfolio_max_gross_var.get() == "1.20"
    assert "Loaded from regime UI Risk" in page.regime_provenance_var.get()


def test_backtesting_trained_regime_invalid_bundle_shows_error(monkeypatch, tmp_path):
    bundle_manifest = tmp_path / "exports" / "bad" / "bundle_manifest.json"
    bundle_manifest.parent.mkdir(parents=True, exist_ok=True)
    bundle_manifest.write_text('{"bundle_id":"bad","bundle_version":"9.0","contents":{}}', encoding="utf-8")

    errors = []
    monkeypatch.setattr(backtesting_page.messagebox, "showerror", lambda title, msg: errors.append((title, msg)))

    page = backtesting_page.BacktestingPage.__new__(backtesting_page.BacktestingPage)
    page.trained_regime_var = Var("bad bundle")
    page._regime_option_lookup = {
        "bad bundle": backtesting_page.RegimeBacktestOption(
            option_id="bundle:bad",
            label="bad bundle",
            source="bundle",
            manifest_path=str(bundle_manifest),
        )
    }
    page.regime_provenance_var = Var("")
    page.regime_diff_var = Var("")
    page.regime_lock_var = Var(False)
    page._regime_locked_fields = False
    page._active_regime_contract = None
    page._regime_loaded_values = {}
    page._apply_regime_lock_state = lambda: None

    page._on_trained_regime_selected()

    assert errors
    assert "compatibility" in errors[-1][0].lower()

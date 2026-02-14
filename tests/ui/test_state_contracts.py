from __future__ import annotations

import ast
from pathlib import Path

from state import AppState

UI_FILES = [
    "src/ui/main_menu.py",
    "src/ui/analysis_page.py",
    "src/ui/backtesting_page.py",
    "src/ui/research_lab_page.py",
    "src/ui/general_analysis_page.py",
    "src/ui/spread_analysis_page.py",
    "src/ui/call_put_analysis_page.py",
    "src/ui/ticker_select_page.py",
    "src/ui/ticker_entry_page.py",
]


class StateAccessVisitor(ast.NodeVisitor):
    def __init__(self):
        self.keys: set[str] = set()

    def visit_Attribute(self, node: ast.Attribute):
        # Match self.controller.state.<field>
        if isinstance(node.value, ast.Attribute) and node.value.attr == "state":
            controller_attr = node.value.value
            if isinstance(controller_attr, ast.Attribute) and controller_attr.attr == "controller":
                self.keys.add(node.attr)
        self.generic_visit(node)


def test_ui_state_keys_align_with_app_state_contract():
    known = set(AppState.__dataclass_fields__.keys())
    for rel_path in UI_FILES:
        tree = ast.parse(Path(rel_path).read_text(encoding="utf-8"))
        visitor = StateAccessVisitor()
        visitor.visit(tree)
        assert visitor.keys <= known, f"{rel_path} touches unknown state keys: {sorted(visitor.keys - known)}"


def test_required_state_contracts_present_for_core_pages():
    expected = {
        "src/ui/ticker_entry_page.py": {"tickers", "selected_ticker"},
        "src/ui/ticker_select_page.py": {"tickers", "selected_ticker"},
        "src/ui/analysis_page.py": {"selected_ticker", "analysis_mode", "option_strategy"},
        "src/ui/general_analysis_page.py": {"tickers", "general_analysis_settings"},
        "src/ui/backtesting_page.py": {"tickers", "backtest_settings", "backtest_templates"},
    }
    for rel_path, expected_keys in expected.items():
        tree = ast.parse(Path(rel_path).read_text(encoding="utf-8"))
        visitor = StateAccessVisitor()
        visitor.visit(tree)
        assert expected_keys <= visitor.keys

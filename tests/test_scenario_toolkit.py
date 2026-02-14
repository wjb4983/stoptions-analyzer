from __future__ import annotations

import pytest

from src.backtesting.scenario_toolkit import list_scenario_pack_templates, resolve_scenario_pack_templates


def test_scenario_pack_templates_include_required_named_templates() -> None:
    templates = set(list_scenario_pack_templates())
    assert {"volatility_shock", "liquidity_crunch", "gap_risk_burst", "correlation_spike"}.issubset(templates)


def test_resolve_scenario_pack_templates_expands_transforms() -> None:
    rows = resolve_scenario_pack_templates(["volatility_shock", "correlation_spike"])
    names = {row["name"] for row in rows}
    assert "pack_volatility_shock" in names
    assert "pack_correlation_spike" in names
    assert all(isinstance(row.get("transforms"), dict) for row in rows)


def test_resolve_scenario_pack_templates_rejects_unknown_pack() -> None:
    with pytest.raises(ValueError):
        resolve_scenario_pack_templates(["unknown_pack"])

from ui.backtesting_insights import aggregate_regime_market_stress, aggregate_rows


def test_aggregate_rows_sums_and_means() -> None:
    rows = [
        {"regime": "risk_on", "pnl": 10.0, "cost": 1.0},
        {"regime": "risk_on", "pnl": "5.0", "cost": 2.0},
        {"regime": "risk_off", "pnl": -3.0, "cost": 0.5},
    ]
    out = aggregate_rows(rows, group_field="regime", numeric_fields=["pnl", "cost"])
    lookup = {row["regime"]: row for row in out}
    assert lookup["risk_on"]["count"] == 2
    assert lookup["risk_on"]["pnl_total"] == 15.0
    assert lookup["risk_on"]["pnl_mean"] == 7.5


def test_aggregate_regime_market_stress_shapes() -> None:
    rows = [
        {"regime": "r1", "market_state": "bull", "stress_scenario": "none", "pnl": 2.0, "cost": 0.2},
        {"regime": "r2", "market_state": "bear", "stress_scenario": "shock", "pnl": -1.0, "cost": 0.3},
    ]
    out = aggregate_regime_market_stress(rows)
    assert set(out.keys()) == {"by_regime", "by_market_state", "by_stress_scenario"}
    assert out["by_regime"]
    assert out["by_market_state"]
    assert out["by_stress_scenario"]

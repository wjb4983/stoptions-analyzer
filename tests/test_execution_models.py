from __future__ import annotations

import json

import numpy as np

from src.backtesting.execution import (
    AssetClassCarryCost,
    CarryModel,
    ParticipationImpactSlippage,
    SpreadSlippage,
    VolatilityScaledSlippage,
    calibrate_impact_coefficient_bps,
    load_impact_calibration_buckets,
    load_slippage_calibration_snapshots,
    select_slippage_calibration_snapshot,
)
from src.backtesting.cache_runner import _build_slippage_model
from src.backtesting.vectorized import backtest_vectorized


def test_spread_slippage_is_side_aware() -> None:
    model = SpreadSlippage(spread_bps=10.0)
    buy = model.apply(100.0, 1.0)
    sell = model.apply(100.0, -1.0)
    assert buy > 100.0
    assert sell < 100.0
    assert np.isclose((buy - 100.0) / 100.0, (100.0 - sell) / 100.0)


def test_spread_slippage_increases_with_worse_queue_rank() -> None:
    model = SpreadSlippage(spread_bps=10.0)
    front = model.apply(100.0, 1.0, {"queue_rank_proxy": 0.0})
    back = model.apply(100.0, 1.0, {"queue_rank_proxy": 1.0})
    assert back > front


def test_participation_impact_cost_monotonicity() -> None:
    prices = np.array([100.0, 100.0, 100.0, 100.0])
    low_signals = np.array([0.0, 0.1, 0.1, 0.1])
    high_signals = np.array([0.0, 0.8, 0.8, 0.8])
    model = ParticipationImpactSlippage(base_bps=0.0, impact_coefficient_bps=50.0, participation_exponent=1.0)
    common_kwargs = {
        "slippage_model": model,
        "volumes": np.full_like(prices, 1.0),
        "adv": np.full_like(prices, 1.0),
    }

    low = backtest_vectorized(prices, low_signals, **common_kwargs)
    high = backtest_vectorized(prices, high_signals, **common_kwargs)

    assert high.cost_breakdown["totals"]["slippage"] > low.cost_breakdown["totals"]["slippage"]


def test_participation_impact_uses_realized_participation_when_provided() -> None:
    model = ParticipationImpactSlippage(base_bps=0.0, impact_coefficient_bps=100.0, participation_exponent=1.0)
    low = model.apply(100.0, 10.0, {"realized_participation": 0.1, "queue_rank_proxy": 0.0})
    high = model.apply(100.0, 10.0, {"realized_participation": 0.4, "queue_rank_proxy": 0.0})
    assert high > low


def test_volatility_scaled_slippage_increases_with_volatility() -> None:
    model = VolatilityScaledSlippage(base_bps=5.0, target_volatility=0.01)
    low = model.apply(100.0, 1.0, {"volatility": 0.01})
    high = model.apply(100.0, 1.0, {"volatility": 0.05})
    assert high > low


def test_volatility_scaled_slippage_penalizes_queue_and_participation() -> None:
    model = VolatilityScaledSlippage(base_bps=5.0, target_volatility=0.01)
    relaxed = model.apply(100.0, 1.0, {"volatility": 0.02, "realized_participation": 0.0, "queue_rank_proxy": 0.0})
    stressed = model.apply(100.0, 1.0, {"volatility": 0.02, "realized_participation": 0.5, "queue_rank_proxy": 1.0})
    assert stressed > relaxed


def test_calibration_helpers_roundtrip(tmp_path) -> None:
    buckets = [
        {"participation": 0.1, "slippage_bps": 2.0, "count": 10},
        {"participation": 0.2, "slippage_bps": 4.0, "count": 10},
    ]
    path = tmp_path / "impact_buckets.json"
    path.write_text(json.dumps(buckets))
    loaded = load_impact_calibration_buckets(path)
    coeff = calibrate_impact_coefficient_bps(loaded)
    model = ParticipationImpactSlippage.from_calibration_buckets(loaded)
    assert np.isclose(coeff, 20.0)
    assert np.isclose(model.impact_coefficient_bps, 20.0)


def test_asset_class_carry_respects_class_rates() -> None:
    model = AssetClassCarryCost(
        asset_classes=["equity", "etf"],
        annual_short_borrow_rates={"equity": 0.06, "etf": 0.01},
        annual_long_financing_rates={"equity": 0.02, "etf": 0.00},
        periods_per_year=252.0,
    )
    positions = np.array([[1.0, -1.0]])
    carry = model.calculate(positions)
    assert carry.shape == positions.shape
    assert carry[0, 0] > carry[0, 1]


def test_carry_model_applies_time_varying_borrow_and_availability() -> None:
    positions = np.array([
        [-1.0],
        [-1.0],
        [-1.0],
    ])
    model = CarryModel(
        asset_classes=["equity"],
        annual_short_borrow_rates={"equity": 0.03},
        borrow_rate_series=np.array([0.01, 0.10, 0.01]),
        borrow_available_flags=np.array([True, False, True]),
        hard_to_borrow_spike_multiplier=3.0,
        periods_per_year=1.0,
    )
    carry = model.calculate(positions)
    assert carry.shape == positions.shape
    assert carry[1, 0] > carry[0, 0]
    assert carry[1, 0] > carry[2, 0]


def test_carry_model_applies_asset_class_specific_components() -> None:
    positions = np.array([[1.0, 1.0, -1.0]])
    model = CarryModel(
        asset_classes=["futures", "options", "equity"],
        annual_short_borrow_rates={"equity": 0.02},
        annual_futures_roll_rates={"futures": 0.06},
        annual_options_theta_rates={"options": 0.12},
        periods_per_year=12.0,
    )
    carry = model.calculate(positions)
    assert carry.shape == positions.shape
    # options theta proxy (0.01) > futures roll (0.005) > equity borrow (0.00166..)
    assert carry[0, 1] > carry[0, 0] > carry[0, 2]


def test_backtest_vectorized_exports_carry_attribution_by_asset() -> None:
    prices = np.array(
        [
            [100.0, 50.0],
            [100.0, 50.0],
            [100.0, 50.0],
        ]
    )
    signals = np.array(
        [
            [0.0, 0.0],
            [-1.0, 1.0],
            [-1.0, 1.0],
        ]
    )
    carry = CarryModel(
        asset_classes=["equity", "futures"],
        annual_short_borrow_rates={"equity": 0.12},
        annual_futures_roll_rates={"futures": 0.24},
        periods_per_year=12.0,
    )
    result = backtest_vectorized(
        prices,
        signals,
        borrow_cost_model=carry,
        carry_asset_classes=["equity", "futures"],
    )
    assert "carry_attribution_by_asset" in result.cost_breakdown
    attribution = np.asarray(result.cost_breakdown["carry_attribution_by_asset"])
    assert attribution.shape == prices.shape
    assert np.any(attribution > 0.0)


def test_load_slippage_calibration_snapshots_and_date_selection(tmp_path) -> None:
    payload = {
        "default_params": {"impact_coefficient_bps": 15.0},
        "snapshots": [
            {"effective_date": "2024-01-01", "stable": True, "params": {"impact_coefficient_bps": 18.0}},
            {"effective_date": "2024-03-01", "stable": True, "params": {"impact_coefficient_bps": 24.0}},
        ],
    }
    path = tmp_path / "snapshots.json"
    path.write_text(json.dumps(payload))
    loaded = load_slippage_calibration_snapshots(path)
    selected = select_slippage_calibration_snapshot(loaded, as_of_date="2024-02-15")
    assert selected.source == "snapshot"
    assert selected.effective_date == "2024-01-01"
    assert selected.params["impact_coefficient_bps"] == 18.0


def test_slippage_snapshot_fallback_to_defaults_with_warning(tmp_path) -> None:
    payload = {
        "default_params": {
            "base_bps": 1.0,
            "impact_coefficient_bps": 12.0,
            "participation_exponent": 1.0,
            "max_participation": 1.0,
        },
        "snapshots": [
            {"effective_date": "2025-01-01", "stable": True, "params": {"impact_coefficient_bps": 30.0}},
        ],
    }
    path = tmp_path / "snapshots.json"
    path.write_text(json.dumps(payload))

    model, selection = _build_slippage_model(
        model_name="participation",
        costs_bps=5.0,
        params={"snapshot_path": str(path)},
        as_of_date="2024-06-01",
    )
    assert model.__class__.__name__ == "ParticipationImpactSlippage"
    assert selection.source == "default_params"
    assert "slippage_calibration_snapshot_unavailable_for_date" in selection.warning_flags
    assert model.impact_coefficient_bps == 12.0


def test_slippage_snapshot_uses_latest_stable_snapshot(tmp_path) -> None:
    payload = {
        "default_params": {"impact_coefficient_bps": 10.0},
        "snapshots": [
            {"effective_date": "2024-01-01", "stable": False, "params": {"impact_coefficient_bps": 100.0}},
            {"effective_date": "2024-02-01", "stable": True, "params": {"impact_coefficient_bps": 25.0}},
            {"effective_date": "2024-04-01", "stable": True, "params": {"impact_coefficient_bps": 35.0}},
        ],
    }
    path = tmp_path / "snapshots.json"
    path.write_text(json.dumps(payload))

    model, selection = _build_slippage_model(
        model_name="participation",
        costs_bps=5.0,
        params={"snapshot_path": str(path), "base_bps": 2.0},
        as_of_date="2024-03-15",
    )
    assert model.__class__.__name__ == "ParticipationImpactSlippage"
    assert selection.source == "snapshot"
    assert selection.effective_date == "2024-02-01"
    assert model.impact_coefficient_bps == 25.0

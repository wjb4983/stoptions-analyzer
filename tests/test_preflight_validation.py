from __future__ import annotations

from datetime import datetime, timezone

import numpy as np
import pytest

from src.backtesting import cache_runner
from src.data_access.engine_loader import EngineArrayBundle, EngineArrayMetadata


def _bundle(*, timestamps: np.ndarray, missing_mask: np.ndarray, coverage: float = 1.0, adjustment_violations: int = 0) -> EngineArrayBundle:
    metadata = EngineArrayMetadata(
        symbol_to_column={"AAPL": 0},
        date_index=timestamps,
        missingness_ratio=float(np.mean(missing_mask)),
        missingness_by_symbol={"AAPL": float(np.mean(missing_mask[:, 0]))},
        coverage_by_symbol={"AAPL": float(coverage)},
        tradable_ratio_by_symbol={"AAPL": 1.0 - float(np.mean(missing_mask[:, 0]))},
        excluded_symbols={},
        audit_summary_by_symbol={},
        asset_class_by_symbol={"AAPL": "equity"},
        expiry_by_symbol={"AAPL": None},
        strike_by_symbol={"AAPL": None},
        option_type_by_symbol={"AAPL": None},
        multiplier_by_symbol={"AAPL": 1.0},
        settlement_style_by_symbol={"AAPL": "physical"},
        borrow_availability_tier_by_symbol={"AAPL": "normal"},
        financing_benchmark_by_symbol={"AAPL": "overnight"},
        pit_membership_violations_by_symbol={"AAPL": 0},
        adjustment_violations_by_symbol={"AAPL": int(adjustment_violations)},
        delisted_symbols=[],
        survivorship_bias_flags_by_symbol={"AAPL": False},
        leakage_flags_by_symbol={"AAPL": False},
        data_fingerprint={},
    )
    values = np.array([[100.0], [101.0], [102.0]], dtype=np.float64)
    return EngineArrayBundle(
        date_index=timestamps,
        open_prices=values,
        close_prices=values,
        raw_open_prices=values,
        raw_close_prices=values,
        split_factors=np.ones_like(values),
        dividends=np.zeros_like(values),
        missing_mask=missing_mask,
        metadata=metadata,
    )


def test_preflight_writes_artifact_and_passes(tmp_path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(cache_runner, "BACKTEST_OUTPUT_DIR", tmp_path)
    timestamps = np.array([1_704_206_400_000, 1_704_206_460_000, 1_704_206_520_000], dtype=np.int64)
    bundle = _bundle(timestamps=timestamps, missing_mask=np.array([[False], [False], [False]], dtype=bool), coverage=1.0)

    report = cache_runner._run_preflight_or_raise(
        arrays=bundle,
        requested_tickers=["AAPL"],
        start_dt=datetime(2024, 1, 2, tzinfo=timezone.utc),
        end_dt=datetime(2024, 1, 3, tzinfo=timezone.utc),
        timeframe="1m",
        config=cache_runner.PreflightValidationConfig(),
        workflow_label="unit",
    )

    assert report["status"] == "pass"
    artifacts = list(tmp_path.glob("preflight_report_unit_*.json"))
    assert artifacts


def test_preflight_blocks_on_critical_failure(tmp_path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(cache_runner, "BACKTEST_OUTPUT_DIR", tmp_path)
    timestamps = np.array([1_704_206_400_000, 1_704_206_460_000, 1_704_206_520_000], dtype=np.int64)
    bundle = _bundle(
        timestamps=timestamps,
        missing_mask=np.array([[False], [False], [False]], dtype=bool),
        coverage=0.3,
        adjustment_violations=1,
    )

    with pytest.raises(ValueError, match="Preflight validation blocked workflow"):
        cache_runner._run_preflight_or_raise(
            arrays=bundle,
            requested_tickers=["AAPL", "MSFT"],
            start_dt=datetime(2024, 1, 2, tzinfo=timezone.utc),
            end_dt=datetime(2024, 1, 3, tzinfo=timezone.utc),
            timeframe="1m",
            config=cache_runner.PreflightValidationConfig(min_symbol_coverage_ratio=0.9),
            workflow_label="unit",
        )

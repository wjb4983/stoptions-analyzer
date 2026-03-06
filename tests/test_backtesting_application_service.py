from __future__ import annotations

from datetime import date
from pathlib import Path

import pytest

from backtesting.application_service import (
    BacktestRequestValidationError,
    BacktestingApplicationService,
)
from backtesting.regime_backtest_adapter import RegimeBacktestOption


def _service() -> BacktestingApplicationService:
    return BacktestingApplicationService(output_dir=Path("out"))


def test_build_classic_strategy_request_preserves_mode_and_routing() -> None:
    request = _service().build_classic_strategy_request(
        tickers=["SPY"],
        start_date=date(2023, 1, 1),
        end_date=date(2023, 2, 1),
        run_mode="full_chain",
        selected_backtest_type="classic_strategy",
        strategy="momentum",
        lookback=20,
        skip=1,
        costs_bps=2.0,
        starting_capital=100000.0,
        custom_bet_pct=0.5,
        cache_root=Path("cache"),
        bet_sizing_mode="half_kelly",
        timeframe="1d",
        execution_model="bps",
        execution_model_params={"spread_bps": 2.0},
        portfolio_cfg={"portfolio_method": "equal_weight"},
        governance_payload={"owner": "tester"},
        stress_controls={"selected_profile": "Base"},
        selected_scenario_packs=["base"],
        selected_suite_key="custom",
        suite_composition={"foo": "bar"},
    )

    assert request.run_mode == "full_chain"
    assert request.artifact_routing.cache_root == Path("cache")
    assert request.artifact_routing.output_dir == Path("out")


@pytest.mark.parametrize("run_mode", ["", "invalid"])
def test_build_classic_strategy_request_rejects_invalid_run_mode(run_mode: str) -> None:
    with pytest.raises(BacktestRequestValidationError, match="Unsupported run mode"):
        _service().build_classic_strategy_request(
            tickers=["SPY"],
            start_date=date(2023, 1, 1),
            end_date=date(2023, 2, 1),
            run_mode=run_mode,
            selected_backtest_type="classic_strategy",
            strategy="momentum",
            lookback=20,
            skip=1,
            costs_bps=2.0,
            starting_capital=100000.0,
            custom_bet_pct=0.5,
            cache_root=Path("cache"),
            bet_sizing_mode="half_kelly",
            timeframe="1d",
            execution_model="bps",
            execution_model_params={},
            portfolio_cfg={},
            governance_payload={},
            stress_controls={},
            selected_scenario_packs=[],
            selected_suite_key="custom",
            suite_composition={},
        )


def test_build_trained_regime_request_requires_regime_option() -> None:
    with pytest.raises(BacktestRequestValidationError, match="Select a trained regime manifest"):
        _service().build_trained_regime_replay_request(
            tickers=["SPY"],
            start_date=date(2023, 1, 1),
            end_date=date(2023, 2, 1),
            run_mode="full_chain",
            selected_backtest_type="trained_regime",
            timeframe="1d",
            cache_root=Path("cache"),
            governance_payload={},
            stress_controls={},
            selected_scenario_packs=[],
            selected_suite_key="custom",
            suite_composition={},
            regime_option=None,
        )


def test_build_trained_regime_request_requires_mode_alignment() -> None:
    option = RegimeBacktestOption(
        option_id="oid",
        label="label",
        manifest_path="missing.json",
        source="local",
    )
    with pytest.raises(BacktestRequestValidationError, match="trained_regime workflow"):
        _service().build_trained_regime_replay_request(
            tickers=["SPY"],
            start_date=date(2023, 1, 1),
            end_date=date(2023, 2, 1),
            run_mode="full_chain",
            selected_backtest_type="classic_strategy",
            timeframe="1d",
            cache_root=Path("cache"),
            governance_payload={},
            stress_controls={},
            selected_scenario_packs=[],
            selected_suite_key="custom",
            suite_composition={},
            regime_option=option,
        )

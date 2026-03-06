from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path

from backtesting.regime_backtest_adapter import (
    RegimeBacktestContract,
    RegimeBacktestOption,
    load_regime_backtest_contract,
)


class BacktestRequestValidationError(ValueError):
    """Raised when UI run state cannot be translated into a valid run request."""


@dataclass(frozen=True)
class BacktestArtifactRouting:
    cache_root: Path
    output_dir: Path


@dataclass(frozen=True)
class BaseBacktestRunRequest:
    tickers: list[str]
    start_date: date
    end_date: date
    run_mode: str
    selected_backtest_type: str
    timeframe: str
    artifact_routing: BacktestArtifactRouting
    governance_payload: dict[str, object]
    stress_controls: dict[str, object]
    selected_scenario_packs: list[str]
    selected_suite_key: str
    suite_composition: dict[str, object]


@dataclass(frozen=True)
class ClassicStrategyRunRequest(BaseBacktestRunRequest):
    strategy: str
    lookback: int
    skip: int
    costs_bps: float
    starting_capital: float
    custom_bet_pct: float
    bet_sizing_mode: str
    execution_model: str
    execution_model_params: dict[str, object]
    portfolio_cfg: dict[str, object]


@dataclass(frozen=True)
class TrainedRegimeReplayRunRequest(BaseBacktestRunRequest):
    regime_option: RegimeBacktestOption
    regime_contract: RegimeBacktestContract


class BacktestingApplicationService:
    def __init__(self, *, output_dir: Path) -> None:
        self._output_dir = output_dir

    def build_classic_strategy_request(
        self,
        *,
        tickers: list[str],
        start_date: date,
        end_date: date,
        run_mode: str,
        selected_backtest_type: str,
        strategy: str,
        lookback: int,
        skip: int,
        costs_bps: float,
        starting_capital: float,
        custom_bet_pct: float,
        cache_root: Path,
        bet_sizing_mode: str,
        timeframe: str,
        execution_model: str,
        execution_model_params: dict[str, object],
        portfolio_cfg: dict[str, object],
        governance_payload: dict[str, object],
        stress_controls: dict[str, object],
        selected_scenario_packs: list[str],
        selected_suite_key: str,
        suite_composition: dict[str, object],
    ) -> ClassicStrategyRunRequest:
        self._validate_dates(start_date, end_date)
        self._validate_run_mode(run_mode)
        if selected_backtest_type != "classic_strategy":
            raise BacktestRequestValidationError("Classic request requires classic_strategy workflow mode.")
        if not tickers:
            raise BacktestRequestValidationError("Add tickers before running a backtest.")
        return ClassicStrategyRunRequest(
            tickers=list(tickers),
            start_date=start_date,
            end_date=end_date,
            run_mode=run_mode,
            selected_backtest_type=selected_backtest_type,
            timeframe=timeframe,
            artifact_routing=BacktestArtifactRouting(cache_root=cache_root, output_dir=self._output_dir),
            governance_payload=dict(governance_payload),
            stress_controls=dict(stress_controls),
            selected_scenario_packs=list(selected_scenario_packs),
            selected_suite_key=selected_suite_key,
            suite_composition=dict(suite_composition),
            strategy=strategy,
            lookback=int(lookback),
            skip=int(skip),
            costs_bps=float(costs_bps),
            starting_capital=float(starting_capital),
            custom_bet_pct=float(custom_bet_pct),
            bet_sizing_mode=bet_sizing_mode,
            execution_model=execution_model,
            execution_model_params=dict(execution_model_params),
            portfolio_cfg=dict(portfolio_cfg),
        )

    def build_trained_regime_replay_request(
        self,
        *,
        tickers: list[str],
        start_date: date,
        end_date: date,
        run_mode: str,
        selected_backtest_type: str,
        timeframe: str,
        cache_root: Path,
        governance_payload: dict[str, object],
        stress_controls: dict[str, object],
        selected_scenario_packs: list[str],
        selected_suite_key: str,
        suite_composition: dict[str, object],
        regime_option: RegimeBacktestOption | None,
    ) -> TrainedRegimeReplayRunRequest:
        self._validate_dates(start_date, end_date)
        self._validate_run_mode(run_mode)
        if selected_backtest_type != "trained_regime":
            raise BacktestRequestValidationError("Regime replay request requires trained_regime workflow mode.")
        if not tickers:
            raise BacktestRequestValidationError("Add tickers before running a backtest.")
        if regime_option is None:
            raise BacktestRequestValidationError("Select a trained regime manifest for trained-regime mode.")
        regime_contract = load_regime_backtest_contract(regime_option)
        return TrainedRegimeReplayRunRequest(
            tickers=list(tickers),
            start_date=start_date,
            end_date=end_date,
            run_mode=run_mode,
            selected_backtest_type=selected_backtest_type,
            timeframe=timeframe,
            artifact_routing=BacktestArtifactRouting(cache_root=cache_root, output_dir=self._output_dir),
            governance_payload=dict(governance_payload),
            stress_controls=dict(stress_controls),
            selected_scenario_packs=list(selected_scenario_packs),
            selected_suite_key=selected_suite_key,
            suite_composition=dict(suite_composition),
            regime_option=regime_option,
            regime_contract=regime_contract,
        )

    def _validate_run_mode(self, run_mode: str) -> None:
        if run_mode not in {"full_chain", "stress_only"}:
            raise BacktestRequestValidationError(f"Unsupported run mode: {run_mode}")

    @staticmethod
    def _validate_dates(start_date: date, end_date: date) -> None:
        if start_date >= end_date:
            raise BacktestRequestValidationError("Start date must be before end date.")

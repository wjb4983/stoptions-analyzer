from __future__ import annotations

from datetime import date
from pathlib import Path

from src.backtesting import cache_runner
from src.backtesting.regime_backtest_adapter import RegimeBacktestContract


def test_run_trained_regime_backtest_uses_contract_defaults(monkeypatch, tmp_path: Path) -> None:
    captured: dict[str, object] = {}

    def _fake_run_time_series_momentum_backtest(**kwargs):
        captured.update(kwargs)
        return "ok"

    monkeypatch.setattr(cache_runner, "run_time_series_momentum_backtest", _fake_run_time_series_momentum_backtest)

    contract = RegimeBacktestContract(
        option_id="training:abc",
        regime_name="Risk On",
        source="training_run",
        manifest_path=str(tmp_path / "manifest.json"),
        defaults={
            "strategy": "xsmom",
            "lookback_days": "88",
            "skip_days": "7",
            "portfolio_max_gross_exposure": "1.15",
            "portfolio_min_net_exposure": "-0.35",
            "portfolio_max_net_exposure": "0.35",
            "portfolio_max_symbol_weight": "0.12",
            "portfolio_max_sector_weight": "0.25",
        },
        execution_artifacts={"champion_model_ids": {"Trend": "meta_label_classifier"}},
    )

    result = cache_runner.run_trained_regime_backtest(
        tickers=["AAPL"],
        start_date=date(2024, 1, 1),
        end_date=date(2024, 2, 1),
        cache_root=tmp_path,
        timeframe="1d",
        regime_contract=contract,
    )

    assert captured["strategy"] == "xsmom"
    assert captured["lookback_days"] == 88
    assert captured["skip_days"] == 7
    assert captured["portfolio_max_gross_exposure"] == 1.15
    assert "Champion model IDs" in result
    assert "Risk On" in result

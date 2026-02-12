from __future__ import annotations

import numpy as np

from src.backtesting.portfolio import PortfolioConstructionConfig, construct_target_weights


def test_capped_optimization_respects_constraints() -> None:
    prices = np.array(
        [
            [100.0, 50.0, 25.0, 80.0],
            [101.0, 49.0, 26.0, 82.0],
            [102.0, 48.0, 27.0, 84.0],
            [103.0, 47.0, 26.5, 85.0],
        ]
    )
    raw = np.array(
        [
            [1.0, 1.0, -1.0, -1.0],
            [1.0, -1.0, -1.0, 1.0],
            [1.0, 1.0, 1.0, -1.0],
            [0.0, 1.0, -1.0, 0.0],
        ]
    )
    symbols = ["A", "B", "C", "D"]
    cfg = PortfolioConstructionConfig(
        method="capped_optimization",
        vol_lookback_bars=3,
        max_symbol_weight=0.30,
        max_sector_weight=0.45,
        max_gross_exposure=0.90,
        min_net_exposure=-0.20,
        max_net_exposure=0.20,
        sector_map={"A": "tech", "B": "tech", "C": "fin", "D": "fin"},
    )

    result = construct_target_weights(raw_signals=raw, prices=prices, symbol_order=symbols, config=cfg)
    w = result.target_weights

    assert np.all(np.abs(w) <= 0.30 + 1e-12)
    gross = np.sum(np.abs(w), axis=1)
    net = np.sum(w, axis=1)
    assert np.all(gross <= 0.90 + 1e-12)
    assert np.all(net <= 0.20 + 1e-12)
    assert np.all(net >= -0.20 - 1e-12)

    tech = np.sum(np.abs(w[:, [0, 1]]), axis=1)
    fin = np.sum(np.abs(w[:, [2, 3]]), axis=1)
    assert np.all(tech <= 0.45 + 1e-12)
    assert np.all(fin <= 0.45 + 1e-12)


def test_missing_prices_and_sparse_universe_zero_out_untradable() -> None:
    prices = np.array(
        [
            [100.0, np.nan, 20.0],
            [101.0, np.nan, np.nan],
            [np.nan, np.nan, 22.0],
            [102.0, 30.0, np.nan],
        ]
    )
    raw = np.array(
        [
            [1.0, 1.0, -1.0],
            [1.0, -1.0, -1.0],
            [1.0, 1.0, -1.0],
            [1.0, 1.0, 1.0],
        ]
    )
    cfg = PortfolioConstructionConfig(
        method="inverse_vol",
        max_symbol_weight=0.8,
        max_gross_exposure=1.0,
    )
    result = construct_target_weights(raw_signals=raw, prices=prices, symbol_order=["X", "Y", "Z"], config=cfg)

    weights = result.target_weights
    assert np.all(weights[~np.isfinite(prices)] == 0.0)
    assert np.all(weights[prices <= 0.0] == 0.0)

    gross = result.diagnostics["gross_exposure"]
    assert np.all(gross <= 1.0 + 1e-12)

    turnover_by_symbol = result.diagnostics["turnover_by_symbol"]
    turnover = result.diagnostics["turnover"]
    assert np.allclose(turnover, np.sum(turnover_by_symbol, axis=1))


def test_factor_and_beta_neutrality_constraints() -> None:
    prices = np.array(
        [
            [100.0, 80.0, 50.0, 40.0],
            [101.0, 79.0, 50.5, 39.5],
            [102.0, 78.0, 51.0, 39.0],
            [103.0, 77.0, 52.0, 38.0],
        ]
    )
    raw = np.array(
        [
            [1.5, 0.5, -1.0, -0.5],
            [1.0, 1.0, -0.8, -1.2],
            [1.2, 0.8, -1.3, -0.7],
            [1.1, 0.9, -1.1, -0.9],
        ]
    )
    factor_exposure = np.array(
        [
            [1.0, 0.2],
            [0.7, -0.1],
            [-0.8, 0.3],
            [-0.9, -0.2],
        ]
    )
    beta = np.array([1.2, 1.0, 0.8, 0.7])

    cfg = PortfolioConstructionConfig(
        method="capped_optimization",
        vol_lookback_bars=3,
        factor_exposures=factor_exposure,
        factor_tolerance=2e-3,
        beta_vector=beta,
        beta_tolerance=2e-3,
        max_symbol_weight=0.45,
        max_gross_exposure=1.0,
        min_net_exposure=-0.1,
        max_net_exposure=0.1,
    )
    result = construct_target_weights(raw_signals=raw, prices=prices, symbol_order=["A", "B", "C", "D"], config=cfg)
    w = result.target_weights

    factor_residual = factor_exposure.T @ w[-1]
    beta_residual = float(beta @ w[-1])
    assert np.all(np.abs(factor_residual) <= cfg.factor_tolerance + 5e-4)
    assert abs(beta_residual) <= cfg.beta_tolerance + 5e-4


def test_transaction_cost_penalty_reduces_turnover() -> None:
    prices = np.array(
        [
            [100.0, 100.0, 100.0],
            [101.0, 99.0, 102.0],
            [102.0, 98.0, 101.0],
            [99.0, 103.0, 100.0],
            [98.0, 104.0, 99.0],
        ]
    )
    raw = np.array(
        [
            [1.0, -1.0, 0.5],
            [-1.0, 1.0, -0.5],
            [1.0, -1.0, 0.5],
            [-1.0, 1.0, -0.5],
            [1.0, -1.0, 0.5],
        ]
    )
    symbols = ["A", "B", "C"]

    base_cfg = PortfolioConstructionConfig(
        method="capped_optimization",
        vol_lookback_bars=3,
        max_symbol_weight=0.6,
        max_gross_exposure=1.0,
    )
    tc_cfg = PortfolioConstructionConfig(
        method="capped_optimization",
        vol_lookback_bars=3,
        max_symbol_weight=0.6,
        max_gross_exposure=1.0,
        transaction_cost_penalty=4.0,
        turnover_penalty=0.5,
    )

    base_result = construct_target_weights(raw_signals=raw, prices=prices, symbol_order=symbols, config=base_cfg)
    tc_result = construct_target_weights(raw_signals=raw, prices=prices, symbol_order=symbols, config=tc_cfg)

    base_turnover = float(np.sum(base_result.diagnostics["turnover"]))
    tc_turnover = float(np.sum(tc_result.diagnostics["turnover"]))
    assert tc_turnover < base_turnover


def test_cvar_cap_enforced_in_capped_optimization() -> None:
    prices = np.array(
        [
            [100.0, 100.0, 100.0],
            [101.0, 99.0, 100.5],
            [102.0, 98.0, 99.0],
            [103.0, 97.0, 98.0],
            [104.0, 96.0, 97.0],
        ]
    )
    raw = np.array(
        [
            [1.0, -1.0, 0.5],
            [1.0, -1.0, 0.5],
            [1.0, -1.0, 0.5],
            [1.0, -1.0, 0.5],
            [1.0, -1.0, 0.5],
        ]
    )
    scenarios = np.array(
        [
            [-0.20, -0.20, -0.10],
            [-0.15, -0.10, -0.08],
            [0.03, 0.01, 0.02],
            [0.02, 0.02, 0.01],
        ]
    )

    cfg = PortfolioConstructionConfig(
        method="capped_optimization",
        vol_lookback_bars=3,
        max_symbol_weight=0.9,
        max_gross_exposure=1.0,
        min_net_exposure=-1.0,
        max_net_exposure=1.0,
        max_expected_shortfall=0.02,
        cvar_confidence=0.75,
        tail_scenarios=scenarios,
    )
    result = construct_target_weights(raw_signals=raw, prices=prices, symbol_order=["A", "B", "C"], config=cfg)

    es = result.diagnostics["tail_expected_shortfall"]
    assert np.all(es <= 0.02 + 1e-6)
    assert np.any(result.diagnostics["tail_constraint_active"] > 0)


def test_covariance_regime_and_sample_depth_selection() -> None:
    prices = np.array(
        [
            [100.0, 90.0, 80.0],
            [101.0, 89.0, 79.5],
            [102.0, 88.0, 79.0],
            [103.0, 87.0, 78.5],
            [102.0, 86.5, 78.0],
            [101.0, 86.0, 77.5],
        ]
    )
    raw = np.array(
        [
            [1.0, 0.5, -0.5],
            [1.0, 0.5, -0.5],
            [1.0, 0.5, -0.5],
            [1.0, 0.5, -0.5],
            [1.0, 0.5, -0.5],
            [1.0, 0.5, -0.5],
        ]
    )
    regimes = ["calm", "calm", "stress", "stress", "calm", "stress"]

    cfg = PortfolioConstructionConfig(
        method="capped_optimization",
        vol_lookback_bars=5,
        covariance_estimator="sample",
        covariance_regime_overrides={"stress": "robust", "calm": "ewma"},
        covariance_shrinkage_min_samples=10,
        covariance_robust_min_samples=6,
        regime_labels=regimes,
        max_symbol_weight=0.7,
        max_gross_exposure=1.0,
    )

    result = construct_target_weights(raw_signals=raw, prices=prices, symbol_order=["A", "B", "C"], config=cfg)
    estimators = [row.get("covariance_estimator") for row in result.diagnostics["binding_constraints"]]

    assert estimators[0] == "sample"
    assert estimators[1] == "sample"
    assert all(est == "shrinkage" for est in estimators[2:])

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import numpy as np


PortfolioMethod = Literal["equal_weight", "vol_target", "inverse_vol", "capped_optimization", "hrp", "herc"]
CovarianceEstimator = Literal["sample", "ewma", "shrinkage", "ledoit_wolf", "oas", "robust"]


@dataclass(frozen=True)
class PortfolioConstructionConfig:
    method: PortfolioMethod = "equal_weight"
    vol_lookback_bars: int = 20
    target_volatility: float = 0.10
    max_symbol_weight: float = 0.25
    max_sector_weight: float = 0.60
    max_gross_exposure: float = 1.0
    min_gross_exposure: float = 0.0
    min_net_exposure: float = -1.0
    max_net_exposure: float = 1.0
    max_net_gamma: float | None = None
    max_abs_vega_bucket: float | None = None
    max_abs_delta_per_underlying: float | None = None
    vega_bucket_map: dict[str, str] | None = None
    sector_map: dict[str, str] | None = None
    country_map: dict[str, str] | None = None
    sector_caps: dict[str, float] | None = None
    country_caps: dict[str, float] | None = None

    covariance_estimator: CovarianceEstimator = "sample"
    covariance_ewma_halflife: float = 10.0
    covariance_shrinkage: float = 0.15
    rebalance_frequency_bars: int = 1
    clustering_linkage: str = "single"
    covariance_regime_overrides: dict[str, CovarianceEstimator] | None = None
    covariance_shrinkage_min_samples: int = 20
    covariance_robust_min_samples: int = 8
    regime_labels: list[str] | np.ndarray | None = None

    factor_exposures: np.ndarray | None = None
    factor_covariances: np.ndarray | None = None
    factor_targets: np.ndarray | None = None
    factor_tolerance: float = 1e-3

    beta_vector: np.ndarray | None = None
    beta_target: float = 0.0
    beta_tolerance: float = 1e-3

    risk_aversion: float = 1.0
    turnover_penalty: float = 0.0
    exposure_penalty: float = 0.0
    transaction_cost_penalty: float = 0.0
    implementation_shortfall_penalty: float = 0.0
    scenario_risk_penalty: float = 0.0
    tail_scenarios: np.ndarray | None = None
    cvar_confidence: float = 0.95
    max_expected_shortfall: float | None = None
    optimization_iters: int = 120
    optimization_step: float = 0.08


@dataclass(frozen=True)
class PortfolioConstructionResult:
    target_weights: np.ndarray
    diagnostics: dict[str, np.ndarray | list[dict[str, float | int | bool | str]]]


@dataclass(frozen=True)
class StressReplayResult:
    scenario: str
    portfolio_returns: np.ndarray
    equity_curve: np.ndarray
    max_drawdown: float
    cvar_95: float
    liquidity_breach_count: int
    attribution_by_asset: np.ndarray
    worst_bar: int
    worst_bar_contribution_by_asset: np.ndarray


def construct_target_weights(
    *,
    raw_signals: np.ndarray,
    prices: np.ndarray,
    symbol_order: list[str],
    config: PortfolioConstructionConfig,
    greek_gamma: np.ndarray | None = None,
    greek_vega: np.ndarray | None = None,
    greek_delta: np.ndarray | None = None,
    underlying_by_symbol: dict[str, str] | None = None,
) -> PortfolioConstructionResult:
    signals = np.asarray(raw_signals, dtype=float)
    px = np.asarray(prices, dtype=float)
    if signals.shape != px.shape:
        raise ValueError("raw_signals and prices must have same shape")

    periods, assets = signals.shape
    weights = np.zeros_like(signals, dtype=float)
    turnover_by_symbol = np.zeros_like(signals, dtype=float)
    prev = np.zeros(assets, dtype=float)
    sector_map = config.sector_map or {}

    gamma_arr = _exposure_array_for_horizon(greek_gamma, periods=periods, assets=assets)
    vega_arr = _exposure_array_for_horizon(greek_vega, periods=periods, assets=assets)
    delta_arr = _exposure_array_for_horizon(greek_delta, periods=periods, assets=assets)
    underlier_map = underlying_by_symbol or {}

    factor_exposures = _factor_exposure_for_horizon(config.factor_exposures, periods=periods, assets=assets)
    factor_covariances = _factor_covariance_for_horizon(config.factor_covariances, periods=periods)
    tail_scenarios = _tail_scenarios_for_horizon(config.tail_scenarios, periods=periods, assets=assets)
    regimes = _regime_for_horizon(config.regime_labels, periods=periods)
    beta_vector = None
    if config.beta_vector is not None:
        beta_vector = np.asarray(config.beta_vector, dtype=float)
        if beta_vector.shape != (assets,):
            raise ValueError("beta_vector must have shape (assets,)")

    binding_constraints: list[dict[str, float | int | bool | str]] = []
    grad_norm = np.zeros(periods, dtype=float)
    constraint_violation = np.zeros(periods, dtype=float)
    long_risk_contrib = np.zeros(periods, dtype=float)
    short_risk_contrib = np.zeros(periods, dtype=float)
    factor_risk_contrib = np.zeros(periods, dtype=float)
    tail_expected_shortfall = np.zeros(periods, dtype=float)
    tail_constraint_active = np.zeros(periods, dtype=float)
    net_gamma_exposure = np.zeros(periods, dtype=float)
    max_vega_bucket_exposure = np.zeros(periods, dtype=float)
    max_underlying_delta_exposure = np.zeros(periods, dtype=float)

    for idx in range(periods):
        row = np.nan_to_num(signals[idx], nan=0.0, posinf=0.0, neginf=0.0)
        tradable = np.isfinite(px[idx]) & (px[idx] > 0.0)
        row[~tradable] = 0.0

        rebalance_freq = max(1, int(config.rebalance_frequency_bars))
        should_rebalance = idx == 0 or (idx % rebalance_freq == 0)
        if should_rebalance:
            base = _build_base_weights(
                row=row,
                prices=px,
                idx=idx,
                config=config,
                tradable=tradable,
            )
        else:
            base = np.array(prev, dtype=float, copy=True)
            base[~tradable] = 0.0

        diag_row: dict[str, float | int | bool | str] = {}
        if config.method == "capped_optimization":
            cov, cov_estimator = _estimate_covariance(
                prices=px,
                idx=idx,
                lookback=max(2, int(config.vol_lookback_bars)),
                config=config,
                regime=regimes[idx],
            )
            fac = factor_exposures[idx] if factor_exposures is not None else None
            fac_cov = factor_covariances[idx] if factor_covariances is not None else None
            scenarios = tail_scenarios[idx] if tail_scenarios is not None else None
            cov_total = _compose_covariance(base_covariance=cov, factor_exposure=fac, factor_covariance=fac_cov)
            optimized, diag_row = _optimize_information_ratio(
                expected_returns=row,
                covariance=cov_total,
                initial=base,
                prev_weights=prev,
                tradable=tradable,
                symbol_order=symbol_order,
                sector_map=sector_map,
                country_map=config.country_map or {},
                config=config,
                factor_exposure=fac,
                beta_vector=beta_vector,
                scenario_returns=scenarios,
                gamma_exposure=(gamma_arr[idx] if gamma_arr is not None else None),
                vega_exposure=(vega_arr[idx] if vega_arr is not None else None),
                delta_exposure=(delta_arr[idx] if delta_arr is not None else None),
                underlying_by_symbol=underlier_map,
            )
            diag_row["covariance_estimator"] = cov_estimator
            constrained = optimized
            grad_norm[idx] = float(diag_row.get("gradient_norm", 0.0))
            constraint_violation[idx] = float(diag_row.get("constraint_violation", 0.0))
            tail_expected_shortfall[idx] = float(diag_row.get("expected_shortfall", 0.0))
            tail_constraint_active[idx] = float(bool(diag_row.get("tail_constraint_active", False)))
            factor_risk_contrib[idx] = _factor_risk_contribution(constrained, cov_total, fac, fac_cov)
        else:
            constrained = _apply_constraints(
                weights=base,
                prev_weights=prev,
                symbol_order=symbol_order,
                sector_map=sector_map,
                country_map=config.country_map or {},
                config=config,
                factor_exposure=(factor_exposures[idx] if factor_exposures is not None else None),
                beta_vector=beta_vector,
                scenario_returns=(tail_scenarios[idx] if tail_scenarios is not None else None),
                gamma_exposure=(gamma_arr[idx] if gamma_arr is not None else None),
                vega_exposure=(vega_arr[idx] if vega_arr is not None else None),
                delta_exposure=(delta_arr[idx] if delta_arr is not None else None),
                underlying_by_symbol=underlier_map,
            )
        constrained[~tradable] = 0.0
        weights[idx] = constrained
        turnover_by_symbol[idx] = np.abs(constrained - prev)
        prev = constrained

        cov_for_diag, _ = _estimate_covariance(
            prices=px,
            idx=idx,
            lookback=max(2, int(config.vol_lookback_bars)),
            config=config,
            regime=regimes[idx],
        )
        fac = factor_exposures[idx] if factor_exposures is not None else None
        fac_cov = factor_covariances[idx] if factor_covariances is not None else None
        cov_diag = _compose_covariance(base_covariance=cov_for_diag, factor_exposure=fac, factor_covariance=fac_cov)
        rc_long, rc_short = _risk_contribution_by_sleeve(constrained, cov_diag)
        long_risk_contrib[idx] = rc_long
        short_risk_contrib[idx] = rc_short
        gamma_row = gamma_arr[idx] if gamma_arr is not None else None
        vega_row = vega_arr[idx] if vega_arr is not None else None
        delta_row = delta_arr[idx] if delta_arr is not None else None
        net_gamma_exposure[idx] = float(np.dot(constrained, gamma_row)) if gamma_row is not None else 0.0
        max_vega_bucket_exposure[idx] = _max_abs_vega_bucket_exposure(
            constrained,
            vega_exposure=vega_row,
            symbol_order=symbol_order,
            bucket_map=(config.vega_bucket_map or {}),
        )
        max_underlying_delta_exposure[idx] = _max_abs_underlying_delta_exposure(
            constrained,
            delta_exposure=delta_row,
            symbol_order=symbol_order,
            underlying_by_symbol=underlier_map,
        )

        if config.method != "capped_optimization":
            factor_risk_contrib[idx] = _factor_risk_contribution(constrained, cov_diag, fac, fac_cov)
            scenarios = tail_scenarios[idx] if tail_scenarios is not None else None
            tail_expected_shortfall[idx] = _expected_shortfall_from_scenarios(
                constrained,
                scenario_returns=scenarios,
                confidence=float(config.cvar_confidence),
            )
            tail_constraint_active[idx] = float(
                config.max_expected_shortfall is not None
                and tail_expected_shortfall[idx] >= float(config.max_expected_shortfall) - 1e-8
            )
        binding_constraints.append(diag_row)

    gross = np.sum(np.abs(weights), axis=1)
    net = np.sum(weights, axis=1)
    concentration = np.zeros(periods, dtype=float)
    nonzero = gross > 0
    concentration[nonzero] = np.sum((np.abs(weights[nonzero]) / gross[nonzero, None]) ** 2, axis=1)
    leverage_usage = np.divide(
        gross,
        max(config.max_gross_exposure, 1e-12),
        out=np.zeros_like(gross),
        where=np.isfinite(gross),
    )

    diagnostics: dict[str, np.ndarray | list[dict[str, float | int | bool | str]]] = {
        "gross_exposure": gross,
        "net_exposure": net,
        "concentration": concentration,
        "leverage_usage": leverage_usage,
        "turnover": np.sum(turnover_by_symbol, axis=1),
        "turnover_by_symbol": turnover_by_symbol,
        "binding_constraints": binding_constraints,
        "kkt_gradient_norm": grad_norm,
        "kkt_constraint_violation": constraint_violation,
        "risk_contribution_long_sleeve": long_risk_contrib,
        "risk_contribution_short_sleeve": short_risk_contrib,
        "factor_risk_contribution": factor_risk_contrib,
        "tail_expected_shortfall": tail_expected_shortfall,
        "tail_constraint_active": tail_constraint_active,
        "net_gamma_exposure": net_gamma_exposure,
        "max_vega_bucket_exposure": max_vega_bucket_exposure,
        "max_underlying_delta_exposure": max_underlying_delta_exposure,
    }
    return PortfolioConstructionResult(target_weights=weights, diagnostics=diagnostics)


def replay_weights_under_stress(
    *,
    base_weights: np.ndarray,
    stressed_asset_returns: np.ndarray,
    stressed_available_volume: np.ndarray | None = None,
    stressed_spread_bps: np.ndarray | None = None,
    liquidity_turnover_threshold: float = 0.10,
    scenario_name: str = "stress_scenario",
) -> StressReplayResult:
    """Replay fixed strategy weights through a stressed return path."""

    weights = np.asarray(base_weights, dtype=float)
    asset_returns = np.asarray(stressed_asset_returns, dtype=float)
    if weights.shape != asset_returns.shape:
        raise ValueError("base_weights and stressed_asset_returns must have same shape")

    portfolio_returns = np.sum(weights * asset_returns, axis=1)
    equity_curve = np.cumprod(1.0 + portfolio_returns)
    running_peak = np.maximum.accumulate(equity_curve)
    drawdown = equity_curve / np.where(running_peak == 0.0, 1.0, running_peak) - 1.0
    max_drawdown = float(np.min(drawdown)) if drawdown.size else 0.0

    losses = -portfolio_returns
    if losses.size:
        threshold = float(np.quantile(losses, 0.95))
        tail = losses[losses >= threshold]
        cvar_95 = float(np.mean(tail)) if tail.size else float(max(0.0, threshold))
    else:
        cvar_95 = 0.0

    turnover_proxy = np.sum(np.abs(np.diff(weights, axis=0, prepend=weights[0:1])), axis=1)
    liquidity_multiplier = np.ones_like(turnover_proxy)
    if stressed_available_volume is not None:
        avail = np.asarray(stressed_available_volume, dtype=float)
        if avail.shape != weights.shape:
            raise ValueError("stressed_available_volume must match base_weights shape")
        mean_avail = np.mean(np.clip(avail, 1e-12, None), axis=1)
        liquidity_multiplier *= 1.0 / np.sqrt(mean_avail)
    if stressed_spread_bps is not None:
        spread = np.asarray(stressed_spread_bps, dtype=float)
        if spread.shape != weights.shape:
            raise ValueError("stressed_spread_bps must match base_weights shape")
        liquidity_multiplier *= 1.0 + (np.mean(np.clip(spread, 0.0, None), axis=1) / 10_000.0)
    liquidity_load = turnover_proxy * liquidity_multiplier
    liquidity_breach_count = int(np.sum(liquidity_load > float(liquidity_turnover_threshold)))

    attribution_by_asset = np.sum(weights * asset_returns, axis=0)
    per_bar_contrib = weights * asset_returns
    worst_bar = int(np.argmin(portfolio_returns)) if portfolio_returns.size else 0
    worst_bar_contrib = per_bar_contrib[worst_bar] if per_bar_contrib.size else np.zeros(weights.shape[1], dtype=float)

    return StressReplayResult(
        scenario=str(scenario_name),
        portfolio_returns=portfolio_returns,
        equity_curve=equity_curve,
        max_drawdown=max_drawdown,
        cvar_95=cvar_95,
        liquidity_breach_count=liquidity_breach_count,
        attribution_by_asset=attribution_by_asset,
        worst_bar=worst_bar,
        worst_bar_contribution_by_asset=np.asarray(worst_bar_contrib, dtype=float),
    )


def _build_base_weights(
    *,
    row: np.ndarray,
    prices: np.ndarray,
    idx: int,
    config: PortfolioConstructionConfig,
    tradable: np.ndarray,
) -> np.ndarray:
    active = np.flatnonzero((row != 0.0) & tradable)
    if active.size == 0:
        return np.zeros_like(row)

    signs = np.sign(row)
    if config.method == "equal_weight":
        base = np.zeros_like(row)
        base[active] = signs[active]
        return _normalize_gross(base)

    vol = _rolling_volatility(prices=prices, idx=idx, lookback=max(2, int(config.vol_lookback_bars)))
    inv_vol = np.zeros_like(row)
    valid = (vol > 0.0) & np.isfinite(vol)
    inv_vol[valid] = 1.0 / vol[valid]

    if config.method in {"inverse_vol", "capped_optimization"}:
        base = np.zeros_like(row)
        base[active] = signs[active] * inv_vol[active]
        if np.sum(np.abs(base)) <= 0:
            base[active] = signs[active]
        return _normalize_gross(base)

    if config.method in {"hrp", "herc"}:
        cov, _ = _estimate_covariance(
            prices=prices,
            idx=idx,
            lookback=max(2, int(config.vol_lookback_bars)),
            config=config,
            regime=None,
        )
        active_list = active.tolist()
        if len(active_list) == 1:
            base = np.zeros_like(row)
            base[active_list[0]] = signs[active_list[0]]
            return base
        if config.method == "hrp":
            alloc = _hrp_allocation(cov, active_list, linkage=str(config.clustering_linkage))
        else:
            alloc = _herc_allocation(cov, active_list, linkage=str(config.clustering_linkage))
        base = np.zeros_like(row)
        for asset_idx, weight in zip(active_list, alloc, strict=False):
            base[asset_idx] = signs[asset_idx] * weight
        return _normalize_gross(base)

    if config.method == "vol_target":
        base = np.zeros_like(row)
        base[active] = signs[active]
        base = _normalize_gross(base)
        if np.sum(np.abs(base)) <= 0:
            return base
        portfolio_vol = float(np.sqrt(np.sum((base * np.nan_to_num(vol, nan=0.0)) ** 2)))
        if portfolio_vol <= 0:
            return base
        scaled = base * (float(config.target_volatility) / portfolio_vol)
        return scaled

    raise ValueError(f"Unknown portfolio method: {config.method}")


def _optimize_information_ratio(
    *,
    expected_returns: np.ndarray,
    covariance: np.ndarray,
    initial: np.ndarray,
    prev_weights: np.ndarray,
    tradable: np.ndarray,
    symbol_order: list[str],
    sector_map: dict[str, str],
    country_map: dict[str, str],
    config: PortfolioConstructionConfig,
    factor_exposure: np.ndarray | None,
    beta_vector: np.ndarray | None,
    scenario_returns: np.ndarray | None,
    gamma_exposure: np.ndarray | None,
    vega_exposure: np.ndarray | None,
    delta_exposure: np.ndarray | None,
    underlying_by_symbol: dict[str, str],
) -> tuple[np.ndarray, dict[str, float | int | bool | str]]:
    w = np.array(initial, dtype=float, copy=True)
    mu = np.nan_to_num(expected_returns, nan=0.0)
    cov = np.nan_to_num(covariance, nan=0.0)
    cov = 0.5 * (cov + cov.T)
    step = max(1e-4, float(config.optimization_step))
    iters = max(1, int(config.optimization_iters))

    for _ in range(iters):
        grad = mu - float(config.risk_aversion) * (cov @ w)
        grad -= float(config.turnover_penalty + config.transaction_cost_penalty) * np.sign(w - prev_weights)

        shortfall_penalty = float(config.implementation_shortfall_penalty)
        if shortfall_penalty > 0:
            grad -= shortfall_penalty * (w - prev_weights)

        if factor_exposure is not None and float(config.exposure_penalty) > 0:
            target = _factor_target_vector(config=config, factors=factor_exposure.shape[1])
            mismatch = factor_exposure.T @ w - target
            grad -= float(config.exposure_penalty) * (factor_exposure @ mismatch)

        if scenario_returns is not None and float(config.scenario_risk_penalty) > 0:
            scenario_pnl = scenario_returns @ w
            losses = -scenario_pnl
            tail_cut = np.quantile(losses, float(np.clip(config.cvar_confidence, 0.5, 0.999)))
            tail_mask = losses >= tail_cut
            if np.any(tail_mask):
                tail_grad = -np.mean(scenario_returns[tail_mask], axis=0)
                grad -= float(config.scenario_risk_penalty) * tail_grad

        w = w + step * grad
        w = _apply_constraints(
            weights=w,
            prev_weights=prev_weights,
            symbol_order=symbol_order,
            sector_map=sector_map,
            country_map=country_map,
            config=config,
            factor_exposure=factor_exposure,
            beta_vector=beta_vector,
            scenario_returns=scenario_returns,
            gamma_exposure=gamma_exposure,
            vega_exposure=vega_exposure,
            delta_exposure=delta_exposure,
            underlying_by_symbol=underlying_by_symbol,
        )
        w[~tradable] = 0.0

    grad = mu - float(config.risk_aversion) * (cov @ w)
    if float(config.implementation_shortfall_penalty) > 0:
        grad -= float(config.implementation_shortfall_penalty) * (w - prev_weights)
    raw_violation = _constraint_violation(
        w=w,
        symbol_order=symbol_order,
        sector_map=sector_map,
        country_map=country_map,
        config=config,
        factor_exposure=factor_exposure,
        beta_vector=beta_vector,
        scenario_returns=scenario_returns,
        gamma_exposure=gamma_exposure,
        vega_exposure=vega_exposure,
        delta_exposure=delta_exposure,
        underlying_by_symbol=underlying_by_symbol,
    )
    es = _expected_shortfall_from_scenarios(w, scenario_returns=scenario_returns, confidence=float(config.cvar_confidence))
    diag = {
        "max_symbol_cap": bool(np.any(np.isclose(np.abs(w), float(config.max_symbol_weight), atol=1e-5))),
        "gross_cap": bool(np.isclose(np.sum(np.abs(w)), float(config.max_gross_exposure), atol=1e-4)),
        "net_bound": bool(
            np.isclose(np.sum(w), float(config.min_net_exposure), atol=1e-4)
            or np.isclose(np.sum(w), float(config.max_net_exposure), atol=1e-4)
        ),
        "gradient_norm": float(np.linalg.norm(grad)),
        "constraint_violation": float(raw_violation),
        "expected_shortfall": float(es),
        "tail_constraint_active": bool(
            config.max_expected_shortfall is not None and es >= float(config.max_expected_shortfall) - 1e-8
        ),
    }
    return w, diag


def _apply_constraints(
    *,
    weights: np.ndarray,
    prev_weights: np.ndarray,
    symbol_order: list[str],
    sector_map: dict[str, str],
    country_map: dict[str, str],
    config: PortfolioConstructionConfig,
    factor_exposure: np.ndarray | None,
    beta_vector: np.ndarray | None,
    scenario_returns: np.ndarray | None,
    gamma_exposure: np.ndarray | None,
    vega_exposure: np.ndarray | None,
    delta_exposure: np.ndarray | None,
    underlying_by_symbol: dict[str, str],
) -> np.ndarray:
    out = np.array(weights, dtype=float, copy=True)

    max_symbol = max(0.0, float(config.max_symbol_weight))
    if max_symbol > 0:
        out = np.clip(out, -max_symbol, max_symbol)

    _apply_group_caps(out, symbol_order=symbol_order, group_map=sector_map, default_cap=float(config.max_sector_weight), explicit_caps=config.sector_caps)
    _apply_group_caps(out, symbol_order=symbol_order, group_map=country_map, default_cap=float(config.max_gross_exposure), explicit_caps=config.country_caps)

    out = _enforce_net_and_gross(out, config=config)

    if factor_exposure is not None:
        target = _factor_target_vector(config=config, factors=factor_exposure.shape[1])
        out = _project_linear_constraint(
            out,
            matrix=factor_exposure.T,
            target=target,
            tolerance=float(max(config.factor_tolerance, 0.0)),
        )

    if beta_vector is not None:
        out = _project_linear_constraint(
            out,
            matrix=np.asarray(beta_vector, dtype=float)[None, :],
            target=np.array([float(config.beta_target)]),
            tolerance=float(max(config.beta_tolerance, 0.0)),
        )

    out = _enforce_net_and_gross(out, config=config)
    out = _enforce_tail_risk_constraint(out, scenario_returns=scenario_returns, config=config)
    out = _enforce_option_risk_constraints(
        out,
        gamma_exposure=gamma_exposure,
        vega_exposure=vega_exposure,
        delta_exposure=delta_exposure,
        symbol_order=symbol_order,
        underlying_by_symbol=underlying_by_symbol,
        config=config,
    )

    tc_penalty = float(config.transaction_cost_penalty)
    if tc_penalty > 0:
        mix = 1.0 / (1.0 + tc_penalty)
        out = mix * out + (1.0 - mix) * prev_weights
        out = _enforce_net_and_gross(out, config=config)
        out = _enforce_tail_risk_constraint(out, scenario_returns=scenario_returns, config=config)
    return out


def _enforce_net_and_gross(weights: np.ndarray, *, config: PortfolioConstructionConfig) -> np.ndarray:
    out = np.array(weights, copy=True)
    gross = float(np.sum(np.abs(out)))
    max_gross = max(0.0, float(config.max_gross_exposure))
    if gross > 0 and max_gross > 0 and gross > max_gross:
        out *= max_gross / gross
        gross = max_gross

    min_gross = max(0.0, float(config.min_gross_exposure))
    if gross > 0 and gross < min_gross:
        out *= min_gross / gross

    net = float(np.sum(out))
    min_net = float(config.min_net_exposure)
    max_net = float(config.max_net_exposure)
    if max_net < min_net:
        min_net, max_net = max_net, min_net

    target_net = min(max(net, min_net), max_net)
    if abs(target_net - net) > 1e-12:
        long_mask = out > 0
        short_mask = out < 0
        if target_net < net and np.any(long_mask):
            reduction = net - target_net
            long_sum = float(np.sum(out[long_mask]))
            if long_sum > 0:
                out[long_mask] *= max(0.0, 1.0 - reduction / long_sum)
        elif target_net > net and np.any(short_mask):
            increase = target_net - net
            short_mag = float(np.sum(np.abs(out[short_mask])))
            if short_mag > 0:
                out[short_mask] *= max(0.0, 1.0 - increase / short_mag)
    return out


def _apply_group_caps(
    weights: np.ndarray,
    *,
    symbol_order: list[str],
    group_map: dict[str, str],
    default_cap: float,
    explicit_caps: dict[str, float] | None,
) -> None:
    if not group_map:
        return
    groups: dict[str, list[int]] = {}
    for idx, symbol in enumerate(symbol_order):
        group = group_map.get(symbol)
        if group:
            groups.setdefault(group, []).append(idx)
    for group, idxs in groups.items():
        cap = float(explicit_caps[group]) if explicit_caps and group in explicit_caps else default_cap
        cap = max(0.0, cap)
        gross_group = float(np.sum(np.abs(weights[idxs])))
        if gross_group > cap > 0:
            weights[idxs] *= cap / gross_group


def _project_linear_constraint(
    values: np.ndarray,
    *,
    matrix: np.ndarray,
    target: np.ndarray,
    tolerance: float,
) -> np.ndarray:
    if matrix.size == 0:
        return values
    A = np.asarray(matrix, dtype=float)
    b = np.asarray(target, dtype=float)
    mismatch = A @ values - b
    if np.all(np.abs(mismatch) <= tolerance):
        return values
    gram = A @ A.T
    gram += np.eye(gram.shape[0]) * 1e-8
    try:
        lam = np.linalg.solve(gram, mismatch)
    except np.linalg.LinAlgError:
        lam = np.linalg.pinv(gram) @ mismatch
    return values - A.T @ lam


def _constraint_violation(
    *,
    w: np.ndarray,
    symbol_order: list[str],
    sector_map: dict[str, str],
    country_map: dict[str, str],
    config: PortfolioConstructionConfig,
    factor_exposure: np.ndarray | None,
    beta_vector: np.ndarray | None,
    scenario_returns: np.ndarray | None,
    gamma_exposure: np.ndarray | None,
    vega_exposure: np.ndarray | None,
    delta_exposure: np.ndarray | None,
    underlying_by_symbol: dict[str, str],
) -> float:
    v = 0.0
    v = max(v, float(np.max(np.abs(w)) - float(config.max_symbol_weight)))
    gross = float(np.sum(np.abs(w)))
    v = max(v, gross - float(config.max_gross_exposure))
    v = max(v, float(config.min_gross_exposure) - gross)
    net = float(np.sum(w))
    v = max(v, net - float(config.max_net_exposure), float(config.min_net_exposure) - net)

    for group_map, caps, default_cap in (
        (sector_map, config.sector_caps, float(config.max_sector_weight)),
        (country_map, config.country_caps, float(config.max_gross_exposure)),
    ):
        groups: dict[str, list[int]] = {}
        for idx, symbol in enumerate(symbol_order):
            g = group_map.get(symbol)
            if g:
                groups.setdefault(g, []).append(idx)
        for g, idxs in groups.items():
            cap = float(caps[g]) if caps and g in caps else default_cap
            v = max(v, float(np.sum(np.abs(w[idxs]))) - cap)

    if factor_exposure is not None:
        target = _factor_target_vector(config=config, factors=factor_exposure.shape[1])
        v = max(v, float(np.max(np.abs(factor_exposure.T @ w - target)) - float(config.factor_tolerance)))

    if beta_vector is not None:
        beta_gap = float(np.abs(np.dot(beta_vector, w) - float(config.beta_target)) - float(config.beta_tolerance))
        v = max(v, beta_gap)

    if config.max_expected_shortfall is not None:
        es = _expected_shortfall_from_scenarios(w, scenario_returns=scenario_returns, confidence=float(config.cvar_confidence))
        v = max(v, es - float(config.max_expected_shortfall))

    if config.max_net_gamma is not None and gamma_exposure is not None:
        v = max(v, abs(float(np.dot(w, gamma_exposure))) - float(config.max_net_gamma))
    if config.max_abs_vega_bucket is not None and vega_exposure is not None:
        v = max(v, _max_abs_vega_bucket_exposure(w, vega_exposure=vega_exposure, symbol_order=symbol_order, bucket_map=(config.vega_bucket_map or {})) - float(config.max_abs_vega_bucket))
    if config.max_abs_delta_per_underlying is not None and delta_exposure is not None:
        v = max(v, _max_abs_underlying_delta_exposure(w, delta_exposure=delta_exposure, symbol_order=symbol_order, underlying_by_symbol=underlying_by_symbol) - float(config.max_abs_delta_per_underlying))
    return max(0.0, float(v))


def _estimate_covariance(
    *, prices: np.ndarray, idx: int, lookback: int, config: PortfolioConstructionConfig, regime: str | None
) -> tuple[np.ndarray, str]:
    returns = _returns_window(prices=prices, idx=idx, lookback=lookback)
    assets = prices.shape[1]
    if returns.size == 0:
        return np.eye(assets) * 1e-6, "sample"

    clean = np.nan_to_num(returns, nan=0.0)
    if clean.shape[0] < 2:
        return np.eye(assets) * 1e-6, "sample"

    est = _select_covariance_estimator(config=config, sample_depth=clean.shape[0], regime=regime)
    if est == "sample":
        cov = np.cov(clean, rowvar=False)
    elif est == "ewma":
        cov = _ewma_covariance(clean, halflife=float(config.covariance_ewma_halflife))
    elif est in {"shrinkage", "ledoit_wolf", "oas"}:
        sample = np.cov(clean, rowvar=False)
        diag = np.diag(np.diag(sample))
        n_samples = max(clean.shape[0], 1)
        n_assets = max(clean.shape[1], 1)
        if est == "ledoit_wolf":
            alpha_auto = float(np.clip((n_assets + 1.0) / (n_samples + n_assets + 1.0), 0.0, 1.0))
        elif est == "oas":
            alpha_auto = float(np.clip((2.0 * n_assets + n_samples) / ((n_assets + 1.0) * max(n_samples - 1.0, 1.0)), 0.0, 1.0))
        else:
            alpha_auto = float(np.clip(config.covariance_shrinkage, 0.0, 1.0))
        manual = float(np.clip(config.covariance_shrinkage, 0.0, 1.0))
        alpha = manual if manual > 0 else alpha_auto
        cov = (1.0 - alpha) * sample + alpha * diag
    elif est == "robust":
        med = np.median(clean, axis=0)
        mad = np.median(np.abs(clean - med), axis=0) + 1e-8
        z = np.clip((clean - med) / (1.4826 * mad), -3.0, 3.0)
        cov = np.cov(z, rowvar=False)
    else:
        raise ValueError(f"Unknown covariance estimator: {est}")

    cov = np.nan_to_num(cov, nan=0.0)
    if cov.ndim == 0:
        cov = np.eye(assets) * float(cov)
    cov = np.asarray(cov, dtype=float)
    if cov.shape != (assets, assets):
        tmp = np.eye(assets) * 1e-6
        m = min(assets, cov.shape[0])
        tmp[:m, :m] = cov[:m, :m]
        cov = tmp
    cov = 0.5 * (cov + cov.T)
    cov += np.eye(assets) * 1e-8
    return cov, est


def _select_covariance_estimator(*, config: PortfolioConstructionConfig, sample_depth: int, regime: str | None) -> CovarianceEstimator:
    est: CovarianceEstimator = config.covariance_estimator
    if regime and config.covariance_regime_overrides and regime in config.covariance_regime_overrides:
        est = config.covariance_regime_overrides[regime]

    if est == "robust" and sample_depth < int(config.covariance_robust_min_samples):
        return "shrinkage"
    if est in {"sample", "ewma"} and sample_depth < int(config.covariance_shrinkage_min_samples):
        return "ledoit_wolf"
    return est


def _compose_covariance(
    *, base_covariance: np.ndarray, factor_exposure: np.ndarray | None, factor_covariance: np.ndarray | None
) -> np.ndarray:
    cov = np.asarray(base_covariance, dtype=float)
    if factor_exposure is not None and factor_covariance is not None:
        cov = cov + factor_exposure @ factor_covariance @ factor_exposure.T
    cov = 0.5 * (cov + cov.T)
    cov += np.eye(cov.shape[0]) * 1e-10
    return cov


def _enforce_tail_risk_constraint(
    weights: np.ndarray,
    *,
    scenario_returns: np.ndarray | None,
    config: PortfolioConstructionConfig,
) -> np.ndarray:
    if config.max_expected_shortfall is None:
        return weights
    es = _expected_shortfall_from_scenarios(weights, scenario_returns=scenario_returns, confidence=float(config.cvar_confidence))
    limit = float(config.max_expected_shortfall)
    if es <= limit + 1e-12:
        return weights
    scale = max(0.0, min(1.0, limit / max(es, 1e-12)))
    return np.asarray(weights, dtype=float) * scale



def _enforce_option_risk_constraints(
    weights: np.ndarray,
    *,
    gamma_exposure: np.ndarray | None,
    vega_exposure: np.ndarray | None,
    delta_exposure: np.ndarray | None,
    symbol_order: list[str],
    underlying_by_symbol: dict[str, str],
    config: PortfolioConstructionConfig,
) -> np.ndarray:
    out = np.asarray(weights, dtype=float).copy()
    if config.max_net_gamma is not None and gamma_exposure is not None:
        net_gamma = float(np.dot(out, gamma_exposure))
        limit = max(1e-12, float(config.max_net_gamma))
        if abs(net_gamma) > limit:
            out *= limit / abs(net_gamma)
    if config.max_abs_vega_bucket is not None and vega_exposure is not None:
        peak = _max_abs_vega_bucket_exposure(out, vega_exposure=vega_exposure, symbol_order=symbol_order, bucket_map=(config.vega_bucket_map or {}))
        limit = max(1e-12, float(config.max_abs_vega_bucket))
        if peak > limit:
            out *= limit / peak
    if config.max_abs_delta_per_underlying is not None and delta_exposure is not None:
        peak = _max_abs_underlying_delta_exposure(out, delta_exposure=delta_exposure, symbol_order=symbol_order, underlying_by_symbol=underlying_by_symbol)
        limit = max(1e-12, float(config.max_abs_delta_per_underlying))
        if peak > limit:
            out *= limit / peak
    return out


def _max_abs_vega_bucket_exposure(
    weights: np.ndarray,
    *,
    vega_exposure: np.ndarray | None,
    symbol_order: list[str],
    bucket_map: dict[str, str],
) -> float:
    if vega_exposure is None:
        return 0.0
    buckets: dict[str, float] = {}
    for idx, symbol in enumerate(symbol_order):
        bucket = bucket_map.get(symbol, symbol)
        buckets[bucket] = buckets.get(bucket, 0.0) + float(weights[idx] * vega_exposure[idx])
    return max((abs(v) for v in buckets.values()), default=0.0)


def _max_abs_underlying_delta_exposure(
    weights: np.ndarray,
    *,
    delta_exposure: np.ndarray | None,
    symbol_order: list[str],
    underlying_by_symbol: dict[str, str],
) -> float:
    if delta_exposure is None:
        return 0.0
    underliers: dict[str, float] = {}
    for idx, symbol in enumerate(symbol_order):
        under = underlying_by_symbol.get(symbol, symbol)
        underliers[under] = underliers.get(under, 0.0) + float(weights[idx] * delta_exposure[idx])
    return max((abs(v) for v in underliers.values()), default=0.0)


def _exposure_array_for_horizon(values: np.ndarray | None, *, periods: int, assets: int) -> np.ndarray | None:
    if values is None:
        return None
    arr = np.asarray(values, dtype=float)
    if arr.shape != (periods, assets):
        return None
    return arr

def _expected_shortfall_from_scenarios(
    weights: np.ndarray,
    *,
    scenario_returns: np.ndarray | None,
    confidence: float,
) -> float:
    if scenario_returns is None or scenario_returns.size == 0:
        return 0.0
    losses = -(np.asarray(scenario_returns, dtype=float) @ np.asarray(weights, dtype=float))
    q = float(np.clip(confidence, 0.5, 0.999))
    cut = np.quantile(losses, q)
    tail = losses[losses >= cut]
    if tail.size == 0:
        return float(max(0.0, cut))
    return float(max(0.0, np.mean(tail)))


def _factor_risk_contribution(
    weights: np.ndarray,
    covariance: np.ndarray,
    factor_exposure: np.ndarray | None,
    factor_covariance: np.ndarray | None,
) -> float:
    total_var = float(weights @ covariance @ weights)
    if total_var <= 0:
        return 0.0
    if factor_exposure is None or factor_covariance is None:
        return 0.0
    factor_var = float(weights @ (factor_exposure @ factor_covariance @ factor_exposure.T) @ weights)
    return float(np.clip(factor_var / total_var, 0.0, 1.0))


def _factor_covariance_for_horizon(covariances: np.ndarray | None, *, periods: int) -> np.ndarray | None:
    if covariances is None:
        return None
    arr = np.asarray(covariances, dtype=float)
    if arr.ndim == 2:
        if arr.shape[0] != arr.shape[1]:
            raise ValueError("factor_covariances with ndim=2 must be square")
        return np.repeat(arr[None, :, :], repeats=periods, axis=0)
    if arr.ndim == 3:
        if arr.shape[0] != periods or arr.shape[1] != arr.shape[2]:
            raise ValueError("factor_covariances with ndim=3 must have shape (periods, factors, factors)")
        return arr
    raise ValueError("factor_covariances must be 2D or 3D")


def _tail_scenarios_for_horizon(scenarios: np.ndarray | None, *, periods: int, assets: int) -> np.ndarray | None:
    if scenarios is None:
        return None
    arr = np.asarray(scenarios, dtype=float)
    if arr.ndim == 2:
        if arr.shape[1] != assets:
            raise ValueError("tail_scenarios with ndim=2 must have shape (scenarios, assets)")
        return np.repeat(arr[None, :, :], repeats=periods, axis=0)
    if arr.ndim == 3:
        if arr.shape[0] != periods or arr.shape[2] != assets:
            raise ValueError("tail_scenarios with ndim=3 must have shape (periods, scenarios, assets)")
        return arr
    raise ValueError("tail_scenarios must be 2D or 3D")


def _regime_for_horizon(regime_labels: list[str] | np.ndarray | None, *, periods: int) -> list[str | None]:
    if regime_labels is None:
        return [None] * periods
    if isinstance(regime_labels, list):
        labels = regime_labels
    else:
        labels = np.asarray(regime_labels, dtype=object).tolist()
    if len(labels) != periods:
        raise ValueError("regime_labels must have one entry per period")
    return [None if x is None else str(x) for x in labels]


def _returns_window(*, prices: np.ndarray, idx: int, lookback: int) -> np.ndarray:
    if idx <= 0:
        return np.empty((0, prices.shape[1]), dtype=float)
    start = max(1, idx - lookback + 1)
    window = prices[start - 1 : idx + 1]
    prev = window[:-1]
    curr = window[1:]
    rets = np.full_like(curr, np.nan, dtype=float)
    valid = np.isfinite(prev) & np.isfinite(curr) & (prev > 0)
    rets[valid] = curr[valid] / prev[valid] - 1.0
    return rets


def _rolling_volatility(*, prices: np.ndarray, idx: int, lookback: int) -> np.ndarray:
    rets = _returns_window(prices=prices, idx=idx, lookback=lookback)
    if rets.size == 0:
        return np.zeros(prices.shape[1], dtype=float)
    vol = np.zeros(prices.shape[1], dtype=float)
    for col in range(rets.shape[1]):
        series = rets[:, col]
        series = series[np.isfinite(series)]
        if series.size >= 2:
            sigma = float(np.std(series))
            if np.isfinite(sigma) and sigma > 0:
                vol[col] = sigma
    return vol


def _ewma_covariance(returns: np.ndarray, *, halflife: float) -> np.ndarray:
    lam = np.exp(np.log(0.5) / max(halflife, 1e-6))
    centered = returns - np.mean(returns, axis=0, keepdims=True)
    cov = np.zeros((returns.shape[1], returns.shape[1]), dtype=float)
    wsum = 0.0
    for i in range(centered.shape[0]):
        weight = lam ** (centered.shape[0] - 1 - i)
        v = centered[i][:, None]
        cov += weight * (v @ v.T)
        wsum += weight
    if wsum > 0:
        cov /= wsum
    return cov


def _correlation_distance(covariance: np.ndarray) -> np.ndarray:
    cov = np.asarray(covariance, dtype=float)
    diag = np.clip(np.diag(cov), 1e-12, None)
    std = np.sqrt(diag)
    corr = cov / np.outer(std, std)
    corr = np.clip(corr, -1.0, 1.0)
    dist = np.sqrt(0.5 * (1.0 - corr))
    dist = np.nan_to_num(dist, nan=1.0, posinf=1.0, neginf=1.0)
    np.fill_diagonal(dist, 0.0)
    return dist


def _linkage_pair(distance: np.ndarray, left: tuple[int, ...], right: tuple[int, ...], linkage: str) -> float:
    vals = [distance[i, j] for i in left for j in right]
    if not vals:
        return 0.0
    if linkage == "complete":
        return float(np.max(vals))
    if linkage == "average":
        return float(np.mean(vals))
    if linkage == "ward":
        return float(np.mean(vals))
    return float(np.min(vals))


def _hierarchical_cluster_order(indices: list[int], covariance: np.ndarray, linkage: str) -> list[int]:
    if len(indices) <= 1:
        return list(indices)
    dist = _correlation_distance(covariance)
    clusters: list[tuple[int, ...]] = [(idx,) for idx in indices]
    while len(clusters) > 1:
        best_i = 0
        best_j = 1
        best_val = _linkage_pair(dist, clusters[0], clusters[1], linkage)
        for i in range(len(clusters) - 1):
            for j in range(i + 1, len(clusters)):
                score = _linkage_pair(dist, clusters[i], clusters[j], linkage)
                if score < best_val:
                    best_i, best_j, best_val = i, j, score
        merged = clusters[best_i] + clusters[best_j]
        for k in sorted((best_i, best_j), reverse=True):
            clusters.pop(k)
        clusters.append(merged)
    return list(clusters[0])


def _cluster_variance(covariance: np.ndarray, cluster: list[int]) -> float:
    if not cluster:
        return 0.0
    sub = covariance[np.ix_(cluster, cluster)]
    inv_diag = 1.0 / np.clip(np.diag(sub), 1e-12, None)
    w = inv_diag / np.sum(inv_diag)
    return float(w @ sub @ w)


def _hrp_allocation(covariance: np.ndarray, indices: list[int], linkage: str) -> np.ndarray:
    order = _hierarchical_cluster_order(indices, covariance, linkage)
    weights = {idx: 1.0 for idx in order}
    queue: list[list[int]] = [order]
    while queue:
        cluster = queue.pop(0)
        if len(cluster) <= 1:
            continue
        split = len(cluster) // 2
        left = cluster[:split]
        right = cluster[split:]
        v_left = max(_cluster_variance(covariance, left), 1e-12)
        v_right = max(_cluster_variance(covariance, right), 1e-12)
        alpha = 1.0 - v_left / (v_left + v_right)
        for i in left:
            weights[i] *= alpha
        for i in right:
            weights[i] *= 1.0 - alpha
        queue.extend([left, right])
    out = np.array([weights[i] for i in indices], dtype=float)
    out = np.clip(out, 0.0, None)
    s = float(np.sum(out))
    return out / s if s > 0 else np.ones(len(indices)) / len(indices)


def _herc_allocation(covariance: np.ndarray, indices: list[int], linkage: str) -> np.ndarray:
    order = _hierarchical_cluster_order(indices, covariance, linkage)
    if len(order) <= 2:
        return _hrp_allocation(covariance, indices, linkage)
    split = max(1, int(np.sqrt(len(order))))
    groups: list[list[int]] = [order[i : i + split] for i in range(0, len(order), split)]
    group_vars = np.array([max(_cluster_variance(covariance, g), 1e-12) for g in groups], dtype=float)
    inv = 1.0 / group_vars
    group_w = inv / np.sum(inv)
    asset_weights = {idx: 0.0 for idx in indices}
    for g_w, g in zip(group_w, groups, strict=False):
        sub = covariance[np.ix_(g, g)]
        inv_diag = 1.0 / np.clip(np.diag(sub), 1e-12, None)
        local = inv_diag / np.sum(inv_diag)
        for i, lw in zip(g, local, strict=False):
            asset_weights[i] = float(g_w * lw)
    out = np.array([asset_weights[i] for i in indices], dtype=float)
    s = float(np.sum(out))
    return out / s if s > 0 else np.ones(len(indices)) / len(indices)


def _factor_exposure_for_horizon(
    exposures: np.ndarray | None,
    *,
    periods: int,
    assets: int,
) -> np.ndarray | None:
    if exposures is None:
        return None
    arr = np.asarray(exposures, dtype=float)
    if arr.ndim == 2:
        if arr.shape[0] != assets:
            raise ValueError("factor_exposures with ndim=2 must have shape (assets, factors)")
        return np.repeat(arr[None, :, :], repeats=periods, axis=0)
    if arr.ndim == 3:
        if arr.shape[0] != periods or arr.shape[1] != assets:
            raise ValueError("factor_exposures with ndim=3 must have shape (periods, assets, factors)")
        return arr
    raise ValueError("factor_exposures must be 2D or 3D")


def _factor_target_vector(*, config: PortfolioConstructionConfig, factors: int) -> np.ndarray:
    if config.factor_targets is None:
        return np.zeros(factors, dtype=float)
    target = np.asarray(config.factor_targets, dtype=float)
    if target.shape != (factors,):
        raise ValueError("factor_targets must have shape (factors,)")
    return target


def _risk_contribution_by_sleeve(weights: np.ndarray, covariance: np.ndarray) -> tuple[float, float]:
    mrc = covariance @ weights
    total_var = float(weights @ mrc)
    if total_var <= 0:
        return 0.0, 0.0
    contrib = weights * mrc / total_var
    long = float(np.sum(contrib[weights > 0]))
    short = float(np.sum(contrib[weights < 0]))
    return long, short


def _normalize_gross(values: np.ndarray) -> np.ndarray:
    gross = float(np.sum(np.abs(values)))
    if gross <= 0:
        return np.zeros_like(values)
    return values / gross

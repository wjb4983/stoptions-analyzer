from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import numpy as np


OptimizationObjective = Literal["mean_variance", "risk_parity", "cvar_minimization"]
AllocationLevel = Literal["instrument", "strategy"]


@dataclass(frozen=True)
class AllocationOptimizationConfig:
    objective: OptimizationObjective = "mean_variance"
    level: AllocationLevel = "instrument"
    risk_aversion: float = 4.0
    turnover_penalty: float = 0.0
    uncertainty_aversion: float = 1.0
    sector_neutrality: bool = False
    sector_neutrality_tolerance: float = 1e-3
    sector_neutrality_penalty: float = 15.0
    cvar_confidence: float = 0.95
    cvar_penalty: float = 8.0
    max_drawdown: float | None = None
    drawdown_penalty: float = 12.0
    default_leverage_cap: float = 1.0
    leverage_caps_by_regime: dict[str, float] | None = None
    strategy_gross_caps: dict[str, float] | None = None
    iterations: int = 180
    step_size: float = 0.07


@dataclass(frozen=True)
class AllocationOptimizationResult:
    weights: np.ndarray
    strategy_weights: dict[str, float]
    diagnostics: dict[str, float | dict[str, float]]
    shadow_prices: dict[str, float]


def optimize_allocation(
    *,
    expected_returns: np.ndarray,
    covariance: np.ndarray,
    previous_weights: np.ndarray | None,
    config: AllocationOptimizationConfig,
    uncertainty: np.ndarray | None = None,
    sectors: list[str] | None = None,
    regime: str | None = None,
    return_scenarios: np.ndarray | None = None,
    instrument_to_strategy: list[str] | None = None,
) -> AllocationOptimizationResult:
    mu = np.asarray(expected_returns, dtype=float)
    cov = np.asarray(covariance, dtype=float)
    if mu.ndim != 1:
        raise ValueError("expected_returns must be 1D")
    if cov.shape != (mu.size, mu.size):
        raise ValueError("covariance must have shape (n_assets, n_assets)")

    prev = np.zeros_like(mu) if previous_weights is None else np.asarray(previous_weights, dtype=float)
    if prev.shape != mu.shape:
        raise ValueError("previous_weights must match expected_returns shape")

    sigma = np.zeros_like(mu) if uncertainty is None else np.asarray(uncertainty, dtype=float)
    if sigma.shape != mu.shape:
        raise ValueError("uncertainty must match expected_returns shape")

    sectors = sectors or [f"asset_{i}" for i in range(mu.size)]
    if len(sectors) != mu.size:
        raise ValueError("sectors must match expected_returns shape")

    strategy_labels = instrument_to_strategy or [f"strategy_{i}" for i in range(mu.size)]
    if len(strategy_labels) != mu.size:
        raise ValueError("instrument_to_strategy must match expected_returns shape")

    w = np.array(prev, dtype=float, copy=True)
    if not np.any(np.abs(w) > 0):
        w = _normalize_gross(np.where(mu >= 0.0, 1.0, -1.0), gross=0.5)

    leverage_cap = _leverage_cap(config=config, regime=regime)
    step = max(1e-4, float(config.step_size))

    for _ in range(max(1, int(config.iterations))):
        grad = _objective_gradient(
            w=w,
            mu=mu,
            cov=cov,
            prev=prev,
            sigma=sigma,
            config=config,
            return_scenarios=return_scenarios,
            level=config.level,
            strategy_labels=strategy_labels,
        )
        w = w - step * grad

        if config.level == "strategy":
            w = _enforce_strategy_caps(w, strategy_labels=strategy_labels, caps=config.strategy_gross_caps)

        if config.sector_neutrality:
            w = _project_sector_neutral(w, sectors=sectors)

        w = _enforce_gross_cap(w, leverage_cap)
        if config.max_drawdown is not None:
            w = _enforce_drawdown_cap(w, return_scenarios=return_scenarios, max_drawdown=float(config.max_drawdown))

    strategy_weights = _strategy_weight_map(w, strategy_labels)
    diagnostics = _build_diagnostics(
        w=w,
        mu=mu,
        cov=cov,
        prev=prev,
        sigma=sigma,
        sectors=sectors,
        regime=regime,
        leverage_cap=leverage_cap,
        return_scenarios=return_scenarios,
        config=config,
    )
    shadow_prices = _estimate_shadow_prices(
        diagnostics=diagnostics,
        config=config,
        leverage_cap=leverage_cap,
    )
    return AllocationOptimizationResult(
        weights=w,
        strategy_weights=strategy_weights,
        diagnostics=diagnostics,
        shadow_prices=shadow_prices,
    )


def _objective_gradient(
    *,
    w: np.ndarray,
    mu: np.ndarray,
    cov: np.ndarray,
    prev: np.ndarray,
    sigma: np.ndarray,
    config: AllocationOptimizationConfig,
    return_scenarios: np.ndarray | None,
    level: AllocationLevel,
    strategy_labels: list[str],
) -> np.ndarray:
    adj_mu = mu - float(config.uncertainty_aversion) * sigma
    grad = -adj_mu + float(config.risk_aversion) * (cov @ w)
    if float(config.turnover_penalty) > 0:
        grad += float(config.turnover_penalty) * np.sign(w - prev)

    if config.objective == "risk_parity":
        grad += _finite_diff_gradient(
            w,
            lambda x: _risk_parity_loss(x, cov=cov, level=level, strategy_labels=strategy_labels),
        )
    elif config.objective == "cvar_minimization":
        grad += float(config.cvar_penalty) * _finite_diff_gradient(
            w,
            lambda x: _portfolio_cvar(x, return_scenarios=return_scenarios, confidence=float(config.cvar_confidence)),
        )
    return grad


def _risk_parity_loss(w: np.ndarray, *, cov: np.ndarray, level: AllocationLevel, strategy_labels: list[str]) -> float:
    if level == "strategy":
        keys = sorted(set(strategy_labels))
        sw = np.array([np.sum(w[np.array(strategy_labels) == key]) for key in keys], dtype=float)
        scov = np.diag(np.maximum(1e-8, np.array([np.mean(np.diag(cov)) for _ in keys], dtype=float)))
        return _risk_parity_loss(sw, cov=scov, level="instrument", strategy_labels=keys)

    port_var = float(w @ cov @ w)
    if port_var <= 1e-12:
        return 1.0
    mrc = cov @ w
    rc = w * mrc / max(port_var, 1e-12)
    target = np.full(w.size, 1.0 / max(1, w.size), dtype=float)
    return float(np.sum((rc - target) ** 2))


def _portfolio_cvar(weights: np.ndarray, *, return_scenarios: np.ndarray | None, confidence: float) -> float:
    if return_scenarios is None:
        return 0.0
    arr = np.asarray(return_scenarios, dtype=float)
    if arr.ndim == 2:
        losses = -(arr @ weights)
    elif arr.ndim == 3:
        path_ret = np.tensordot(arr, weights, axes=(2, 0))
        equity = np.cumprod(1.0 + path_ret, axis=1)
        losses = 1.0 - np.min(equity / np.maximum(np.maximum.accumulate(equity, axis=1), 1e-12), axis=1)
    else:
        return 0.0
    q = float(np.clip(confidence, 0.5, 0.999))
    cut = float(np.quantile(losses, q))
    tail = losses[losses >= cut]
    return float(np.mean(tail)) if tail.size else cut


def _max_drawdown(weights: np.ndarray, *, return_scenarios: np.ndarray | None) -> float:
    if return_scenarios is None:
        return 0.0
    arr = np.asarray(return_scenarios, dtype=float)
    if arr.ndim != 3:
        return max(0.0, _portfolio_cvar(weights, return_scenarios=arr, confidence=0.95))
    path_ret = np.tensordot(arr, weights, axes=(2, 0))
    equity = np.cumprod(1.0 + path_ret, axis=1)
    peak = np.maximum.accumulate(equity, axis=1)
    drawdown = equity / np.maximum(peak, 1e-12) - 1.0
    return float(np.min(drawdown))


def _project_sector_neutral(w: np.ndarray, *, sectors: list[str]) -> np.ndarray:
    out = np.array(w, dtype=float, copy=True)
    labels = np.asarray(sectors)
    for sector in sorted(set(sectors)):
        mask = labels == sector
        if np.any(mask):
            out[mask] -= float(np.mean(out[mask]))
    return out


def _enforce_strategy_caps(weights: np.ndarray, *, strategy_labels: list[str], caps: dict[str, float] | None) -> np.ndarray:
    if not caps:
        return weights
    out = np.array(weights, dtype=float, copy=True)
    labels = np.asarray(strategy_labels)
    for key, cap in caps.items():
        mask = labels == key
        if not np.any(mask):
            continue
        gross = float(np.sum(np.abs(out[mask])))
        cap_v = max(1e-9, float(cap))
        if gross > cap_v:
            out[mask] *= cap_v / gross
    return out


def _enforce_gross_cap(weights: np.ndarray, leverage_cap: float) -> np.ndarray:
    gross = float(np.sum(np.abs(weights)))
    if gross <= leverage_cap + 1e-12:
        return weights
    return np.asarray(weights, dtype=float) * (leverage_cap / max(gross, 1e-12))


def _enforce_drawdown_cap(weights: np.ndarray, *, return_scenarios: np.ndarray | None, max_drawdown: float) -> np.ndarray:
    dd = _max_drawdown(weights, return_scenarios=return_scenarios)
    floor = -abs(float(max_drawdown))
    if dd >= floor:
        return weights
    scale = min(1.0, abs(floor) / max(abs(dd), 1e-12))
    return np.asarray(weights, dtype=float) * scale


def _leverage_cap(*, config: AllocationOptimizationConfig, regime: str | None) -> float:
    if regime and config.leverage_caps_by_regime and regime in config.leverage_caps_by_regime:
        return max(1e-6, float(config.leverage_caps_by_regime[regime]))
    return max(1e-6, float(config.default_leverage_cap))


def _strategy_weight_map(weights: np.ndarray, strategy_labels: list[str]) -> dict[str, float]:
    out: dict[str, float] = {}
    labels = np.asarray(strategy_labels)
    for key in sorted(set(strategy_labels)):
        out[key] = float(np.sum(weights[labels == key]))
    return out


def _build_diagnostics(
    *,
    w: np.ndarray,
    mu: np.ndarray,
    cov: np.ndarray,
    prev: np.ndarray,
    sigma: np.ndarray,
    sectors: list[str],
    regime: str | None,
    leverage_cap: float,
    return_scenarios: np.ndarray | None,
    config: AllocationOptimizationConfig,
) -> dict[str, float | dict[str, float]]:
    pnl = float(mu @ w)
    var = float(w @ cov @ w)
    turnover = float(np.sum(np.abs(w - prev)))
    cvar = _portfolio_cvar(w, return_scenarios=return_scenarios, confidence=float(config.cvar_confidence))
    dd = _max_drawdown(w, return_scenarios=return_scenarios)
    gross = float(np.sum(np.abs(w)))
    sector_net: dict[str, float] = {}
    sec_labels = np.asarray(sectors)
    for sector in sorted(set(sectors)):
        sector_net[sector] = float(np.sum(w[sec_labels == sector]))
    return {
        "expected_return": pnl,
        "variance": var,
        "turnover": turnover,
        "uncertainty_budget": float(np.sum(np.abs(w) * sigma)),
        "gross_exposure": gross,
        "leverage_utilization": gross / max(leverage_cap, 1e-12),
        "regime_leverage_cap": leverage_cap,
        "cvar": cvar,
        "max_drawdown": dd,
        "sector_net_exposure": sector_net,
    }


def _estimate_shadow_prices(
    *,
    diagnostics: dict[str, float | dict[str, float]],
    config: AllocationOptimizationConfig,
    leverage_cap: float,
) -> dict[str, float]:
    gross = float(diagnostics["gross_exposure"])
    leverage_violation = max(0.0, gross - leverage_cap)
    dd = float(diagnostics["max_drawdown"])
    dd_violation = 0.0
    if config.max_drawdown is not None:
        dd_violation = max(0.0, abs(dd) - abs(float(config.max_drawdown)))

    shadow = {
        "gross_leverage_cap": leverage_violation * 100.0,
        "max_drawdown": dd_violation * float(config.drawdown_penalty),
        "turnover_penalty": float(config.turnover_penalty) * float(diagnostics["turnover"]),
    }

    sector_shadow = 0.0
    sector_exp = diagnostics.get("sector_net_exposure")
    if isinstance(sector_exp, dict) and config.sector_neutrality:
        for value in sector_exp.values():
            sector_shadow += max(0.0, abs(float(value)) - float(config.sector_neutrality_tolerance))
        sector_shadow *= float(config.sector_neutrality_penalty)
    shadow["sector_neutrality"] = sector_shadow
    return shadow


def _normalize_gross(weights: np.ndarray, *, gross: float = 1.0) -> np.ndarray:
    w = np.asarray(weights, dtype=float)
    total = float(np.sum(np.abs(w)))
    if total <= 1e-12:
        return np.zeros_like(w)
    return w * (gross / total)


def _finite_diff_gradient(weights: np.ndarray, fn, eps: float = 1e-5) -> np.ndarray:
    grad = np.zeros_like(weights, dtype=float)
    for idx in range(weights.size):
        up = np.array(weights, dtype=float, copy=True)
        dn = np.array(weights, dtype=float, copy=True)
        up[idx] += eps
        dn[idx] -= eps
        grad[idx] = (float(fn(up)) - float(fn(dn))) / (2.0 * eps)
    return grad

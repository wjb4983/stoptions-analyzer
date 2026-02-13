from __future__ import annotations

import csv
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import numpy as np

from .portfolio import replay_weights_under_stress
from .vectorized import (
    StressCorrelationBreak,
    StressLiquidityDrought,
    StressScenario,
    StressShockVector,
    StressVolatilityJump,
    apply_stress_scenario,
)


@dataclass(frozen=True)
class ScenarioSpec:
    """User-defined scenario specification."""

    name: str
    scenario_type: str
    params: dict[str, Any]


def build_custom_scenarios(
    *,
    specs: list[ScenarioSpec],
    n_assets: int,
    sector_by_asset: list[str] | None = None,
) -> list[dict[str, Any]]:
    """Create scenario payloads for requested stress classes."""

    sectors = list(sector_by_asset or [f"asset_{i}" for i in range(n_assets)])
    rows: list[dict[str, Any]] = []
    flat_multiplier = tuple(1.0 for _ in range(n_assets))
    flat_shift = tuple(0.0 for _ in range(n_assets))
    for spec in specs:
        stype = spec.scenario_type.lower().strip()
        params = dict(spec.params)
        if stype == "rate_shock":
            shift = float(params.get("rate_shift", -0.002))
            duration_beta = float(params.get("duration_beta", 1.0))
            shifts = tuple(shift * duration_beta for _ in range(n_assets))
            scenario = StressScenario(
                name=spec.name,
                shock_vector=StressShockVector(
                    returns_multiplier_by_asset=flat_multiplier,
                    returns_shift_by_asset=shifts,
                ),
            )
            rows.append({"spec": spec, "scenario": scenario, "path_adjustments": None})
        elif stype == "vol_shock":
            scenario = StressScenario(
                name=spec.name,
                shock_vector=StressShockVector(
                    returns_multiplier_by_asset=flat_multiplier,
                    returns_shift_by_asset=flat_shift,
                ),
                volatility_jump=StressVolatilityJump(
                    jump_multiplier=float(params.get("jump_multiplier", 1.8)),
                    trigger_quantile=float(params.get("trigger_quantile", 0.6)),
                ),
            )
            rows.append({"spec": spec, "scenario": scenario, "path_adjustments": None})
        elif stype == "gap_risk":
            jump = float(params.get("gap_jump", 0.04))
            shift = tuple(float(params.get("returns_shift", -0.001)) for _ in range(n_assets))
            scenario = StressScenario(
                name=spec.name,
                shock_vector=StressShockVector(
                    returns_multiplier_by_asset=flat_multiplier,
                    returns_shift_by_asset=shift,
                ),
                volatility_jump=StressVolatilityJump(jump_multiplier=max(1.0, 1.0 + jump * 8.0), trigger_quantile=0.0),
            )
            rows.append({"spec": spec, "scenario": scenario, "path_adjustments": None})
        elif stype == "sector_rotation":
            favored = set(str(v) for v in params.get("favored_sectors", []))
            penalized = set(str(v) for v in params.get("penalized_sectors", []))
            up = float(params.get("favored_shift", 0.0015))
            down = float(params.get("penalized_shift", -0.0015))
            neutral = float(params.get("neutral_shift", 0.0))
            shifts: list[float] = []
            for sec in sectors:
                if sec in favored:
                    shifts.append(up)
                elif sec in penalized:
                    shifts.append(down)
                else:
                    shifts.append(neutral)
            scenario = StressScenario(
                name=spec.name,
                shock_vector=StressShockVector(
                    returns_multiplier_by_asset=flat_multiplier,
                    returns_shift_by_asset=tuple(shifts),
                ),
            )
            rows.append({"spec": spec, "scenario": scenario, "path_adjustments": None})
        elif stype == "crash_rebound_path":
            crash_len = max(1, int(params.get("crash_periods", 5)))
            rebound_len = max(1, int(params.get("rebound_periods", 5)))
            crash_shift = float(params.get("crash_shift", -0.01))
            rebound_shift = float(params.get("rebound_shift", 0.007))
            rows.append(
                {
                    "spec": spec,
                    "scenario": StressScenario(
                        name=spec.name,
                        shock_vector=StressShockVector(
                            returns_multiplier_by_asset=flat_multiplier,
                            returns_shift_by_asset=flat_shift,
                        ),
                    ),
                    "path_adjustments": {
                        "crash_len": crash_len,
                        "rebound_len": rebound_len,
                        "crash_shift": crash_shift,
                        "rebound_shift": rebound_shift,
                    },
                }
            )
        else:
            raise ValueError(f"Unsupported scenario_type '{spec.scenario_type}'")
    return rows


def _apply_path_adjustments(asset_returns: np.ndarray, path_adjustments: dict[str, Any] | None) -> np.ndarray:
    adjusted = np.asarray(asset_returns, dtype=float).copy()
    if not path_adjustments or adjusted.size == 0:
        return adjusted
    n_periods = adjusted.shape[0]
    crash_len = min(n_periods, int(path_adjustments.get("crash_len", 0)))
    rebound_len = min(max(0, n_periods - crash_len), int(path_adjustments.get("rebound_len", 0)))
    crash_shift = float(path_adjustments.get("crash_shift", 0.0))
    rebound_shift = float(path_adjustments.get("rebound_shift", 0.0))
    if crash_len > 0:
        adjusted[:crash_len] += crash_shift
    if rebound_len > 0:
        adjusted[crash_len : crash_len + rebound_len] += rebound_shift
    return adjusted


def project_strategy_under_scenarios(
    *,
    base_weights: np.ndarray,
    asset_returns: np.ndarray,
    scenario_payloads: list[dict[str, Any]],
    sector_by_asset: list[str] | None = None,
    liquidity_context: dict[str, np.ndarray] | None = None,
) -> list[dict[str, Any]]:
    """Project PnL/risk/exposure of a strategy for each scenario."""

    weights = np.asarray(base_weights, dtype=float)
    returns = np.asarray(asset_returns, dtype=float)
    if weights.shape != returns.shape:
        raise ValueError("base_weights and asset_returns must have same shape")

    sector_labels = list(sector_by_asset or [f"asset_{i}" for i in range(weights.shape[1])])
    if len(sector_labels) != weights.shape[1]:
        raise ValueError("sector_by_asset must match number of assets")

    rows: list[dict[str, Any]] = []
    for payload in scenario_payloads:
        scenario = payload["scenario"]
        stress_out = apply_stress_scenario(asset_returns=returns, scenario=scenario, liquidity_context=liquidity_context)
        stressed = _apply_path_adjustments(stress_out["asset_returns"], payload.get("path_adjustments"))
        replay = replay_weights_under_stress(
            base_weights=weights,
            stressed_asset_returns=stressed,
            stressed_available_volume=stress_out["liquidity"].get("available_bar_volume"),
            stressed_spread_bps=stress_out["liquidity"].get("spread_bps"),
            scenario_name=scenario.name,
        )
        gross_exposure = float(np.mean(np.sum(np.abs(weights), axis=1)))
        net_exposure = float(np.mean(np.sum(weights, axis=1)))
        sector_exposure: dict[str, float] = {}
        abs_w = np.mean(np.abs(weights), axis=0)
        for idx, label in enumerate(sector_labels):
            sector_exposure[label] = sector_exposure.get(label, 0.0) + float(abs_w[idx])

        rows.append(
            {
                "scenario": scenario.name,
                "scenario_type": payload["spec"].scenario_type,
                "pnl_total": float(np.sum(replay.portfolio_returns)),
                "max_drawdown": float(replay.max_drawdown),
                "cvar_95": float(replay.cvar_95),
                "liquidity_breach_count": int(replay.liquidity_breach_count),
                "gross_exposure": gross_exposure,
                "net_exposure": net_exposure,
                "sector_exposure": sector_exposure,
            }
        )
    return rows


def optimize_hedges_for_scenarios(
    *,
    base_weights: np.ndarray,
    asset_returns: np.ndarray,
    scenario_payloads: list[dict[str, Any]],
    hedge_returns: np.ndarray,
    selected_scenarios: list[str] | None = None,
    n_trials: int = 400,
    hedge_bounds: tuple[float, float] = (-1.0, 1.0),
    l2_penalty: float = 0.05,
    random_seed: int = 42,
) -> dict[str, Any]:
    """Optimize hedge sizes against selected stress scenarios."""

    lo, hi = float(hedge_bounds[0]), float(hedge_bounds[1])
    if hi <= lo:
        raise ValueError("hedge_bounds must satisfy high > low")

    weights = np.asarray(base_weights, dtype=float)
    returns = np.asarray(asset_returns, dtype=float)
    hedges = np.asarray(hedge_returns, dtype=float)
    if weights.shape != returns.shape:
        raise ValueError("base_weights and asset_returns must have same shape")
    if hedges.shape[0] != returns.shape[0]:
        raise ValueError("hedge_returns must have same number of periods as asset_returns")

    scenario_names = set(selected_scenarios or [str(p["scenario"].name) for p in scenario_payloads])
    relevant = [p for p in scenario_payloads if str(p["scenario"].name) in scenario_names]
    if not relevant:
        raise ValueError("No selected scenarios were found")

    rng = np.random.default_rng(int(random_seed))
    n_hedges = hedges.shape[1]
    best_vec = np.zeros(n_hedges, dtype=float)
    best_score = float("inf")
    best_breakdown: list[dict[str, float]] = []

    for _ in range(max(1, int(n_trials))):
        candidate = rng.uniform(lo, hi, size=n_hedges)
        scenario_scores: list[dict[str, float]] = []
        objective = 0.0
        for payload in relevant:
            scenario = payload["scenario"]
            stressed = apply_stress_scenario(asset_returns=returns, scenario=scenario)["asset_returns"]
            stressed = _apply_path_adjustments(stressed, payload.get("path_adjustments"))
            base_port = np.sum(weights * stressed, axis=1)
            hedged_port = base_port + hedges @ candidate
            equity = np.cumprod(1.0 + hedged_port)
            peak = np.maximum.accumulate(equity)
            dd = equity / np.where(peak == 0.0, 1.0, peak) - 1.0
            max_dd = float(np.min(dd)) if dd.size else 0.0
            loss = -hedged_port
            q95 = float(np.quantile(loss, 0.95)) if loss.size else 0.0
            tail = loss[loss >= q95] if loss.size else np.array([], dtype=float)
            cvar = float(np.mean(tail)) if tail.size else q95
            pnl = float(np.sum(hedged_port))
            local = cvar + abs(min(0.0, max_dd)) + max(0.0, -pnl)
            objective += local
            scenario_scores.append({"scenario": scenario.name, "cvar_95": cvar, "max_drawdown": max_dd, "pnl_total": pnl, "local_objective": local})
        objective = objective / len(relevant) + float(l2_penalty) * float(np.sum(candidate**2))
        if objective < best_score:
            best_score = float(objective)
            best_vec = candidate.copy()
            best_breakdown = scenario_scores

    return {
        "selected_scenarios": sorted(scenario_names),
        "optimal_hedge_weights": best_vec.tolist(),
        "objective": float(best_score),
        "scenario_breakdown": best_breakdown,
    }


def export_scenario_comparison_report(
    *,
    scenario_projection_rows: list[dict[str, Any]],
    hedge_optimization_result: dict[str, Any] | None,
    output_dir: Path,
    basename: str = "scenario_comparison_report",
) -> dict[str, str]:
    """Export scenario comparisons to JSON/CSV/Markdown for review meetings."""

    output_dir.mkdir(parents=True, exist_ok=True)
    json_path = output_dir / f"{basename}.json"
    csv_path = output_dir / f"{basename}.csv"
    md_path = output_dir / f"{basename}.md"

    payload = {
        "scenario_projection": scenario_projection_rows,
        "hedge_optimization": hedge_optimization_result,
    }
    json_path.write_text(json.dumps(payload, indent=2))

    fieldnames = [
        "scenario",
        "scenario_type",
        "pnl_total",
        "max_drawdown",
        "cvar_95",
        "liquidity_breach_count",
        "gross_exposure",
        "net_exposure",
    ]
    with csv_path.open("w", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        for row in scenario_projection_rows:
            writer.writerow({k: row.get(k) for k in fieldnames})

    lines = ["# Scenario Comparison Report", "", "## Scenario Summary"]
    for row in scenario_projection_rows:
        lines.append(
            "- {scenario} ({stype}): pnl={pnl:.4f}, max_dd={dd:.4f}, cvar95={cvar:.4f}, gross={gross:.3f}, net={net:.3f}".format(
                scenario=row.get("scenario", "unnamed"),
                stype=row.get("scenario_type", "unknown"),
                pnl=float(row.get("pnl_total", 0.0)),
                dd=float(row.get("max_drawdown", 0.0)),
                cvar=float(row.get("cvar_95", 0.0)),
                gross=float(row.get("gross_exposure", 0.0)),
                net=float(row.get("net_exposure", 0.0)),
            )
        )

    if hedge_optimization_result:
        lines.extend(["", "## Hedge Optimization", f"- objective: {float(hedge_optimization_result.get('objective', 0.0)):.6f}"])
        lines.append(f"- selected scenarios: {', '.join(hedge_optimization_result.get('selected_scenarios', []))}")
        lines.append(f"- hedge weights: {hedge_optimization_result.get('optimal_hedge_weights', [])}")

    md_path.write_text("\n".join(lines) + "\n")
    return {"json": str(json_path), "csv": str(csv_path), "markdown": str(md_path)}

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import numpy as np


@dataclass(frozen=True)
class QualityGateThresholds:
    calibration_mae_bps_max: float = 15.0
    calibration_impact_std_bps_max: float = 20.0
    downside_max_drawdown_floor: float = -0.25
    downside_deviation_max: float = 0.03
    tail_worst_rolling_drawdown_floor: float = -0.12
    capacity_turnover_total_max: float = 150.0
    capacity_slippage_bps_max: float = 25.0
    regime_minimum_pass_rate: float = 1.0


def _bool_gate(name: str, passed: bool, details: dict[str, Any]) -> dict[str, Any]:
    return {"name": name, "pass": bool(passed), "details": details}


def _dimension(scorecard: dict[str, Any], name: str) -> dict[str, Any]:
    dimensions = scorecard.get("dimensions", {})
    if isinstance(dimensions, dict):
        value = dimensions.get(name, {})
        if isinstance(value, dict):
            return value
    return {}


def evaluate_modeling_quality_gates(
    *,
    scorecard: dict[str, Any],
    calibration_report: dict[str, Any],
    baseline_rows: list[dict[str, Any]],
    thresholds: QualityGateThresholds | None = None,
) -> list[dict[str, Any]]:
    cfg = thresholds or QualityGateThresholds()
    baseline = list(baseline_rows)
    if not baseline:
        raise ValueError("baseline_rows cannot be empty")

    fit_error = calibration_report.get("fit_error", {}) if isinstance(calibration_report.get("fit_error"), dict) else {}
    stability = calibration_report.get("stability", {}) if isinstance(calibration_report.get("stability"), dict) else {}
    mae_max = float(fit_error.get("mae_bps_max", np.inf))
    coeff_std = float(stability.get("impact_coefficient_std_bps", np.inf))
    calibration_gate = _bool_gate(
        "calibration_thresholds",
        bool(mae_max <= cfg.calibration_mae_bps_max and coeff_std <= cfg.calibration_impact_std_bps_max),
        {
            "mae_bps_max": mae_max,
            "impact_coefficient_std_bps": coeff_std,
            "thresholds": {
                "mae_bps_max": cfg.calibration_mae_bps_max,
                "impact_coefficient_std_bps": cfg.calibration_impact_std_bps_max,
            },
        },
    )

    max_drawdown = np.asarray([float(row.get("max_drawdown", np.nan)) for row in baseline], dtype=float)
    downside_deviation = np.asarray([float(row.get("downside_deviation", np.nan)) for row in baseline], dtype=float)
    rolling_tail = np.asarray([float(row.get("rolling_drawdown_worst", np.nan)) for row in baseline], dtype=float)
    downside_gate = _bool_gate(
        "downside_tail_risk_constraints",
        bool(
            np.all(np.isfinite(max_drawdown))
            and np.all(np.isfinite(downside_deviation))
            and np.all(np.isfinite(rolling_tail))
            and np.min(max_drawdown) >= cfg.downside_max_drawdown_floor
            and np.max(downside_deviation) <= cfg.downside_deviation_max
            and np.min(rolling_tail) >= cfg.tail_worst_rolling_drawdown_floor
        ),
        {
            "max_drawdown_worst": float(np.min(max_drawdown)),
            "downside_deviation_max": float(np.max(downside_deviation)),
            "rolling_drawdown_worst": float(np.min(rolling_tail)),
            "thresholds": {
                "max_drawdown_floor": cfg.downside_max_drawdown_floor,
                "downside_deviation_max": cfg.downside_deviation_max,
                "rolling_drawdown_worst_floor": cfg.tail_worst_rolling_drawdown_floor,
            },
        },
    )

    turnover_total = np.asarray([float(row.get("turnover_total", np.nan)) for row in baseline], dtype=float)
    slippage_bps = np.asarray([float(row.get("slippage_bps", np.nan)) for row in baseline], dtype=float)
    capacity_gate = _bool_gate(
        "capacity_impact_penalties",
        bool(
            np.all(np.isfinite(turnover_total))
            and np.all(np.isfinite(slippage_bps))
            and np.max(turnover_total) <= cfg.capacity_turnover_total_max
            and np.max(slippage_bps) <= cfg.capacity_slippage_bps_max
        ),
        {
            "turnover_total_max": float(np.max(turnover_total)),
            "slippage_bps_max": float(np.max(slippage_bps)),
            "thresholds": {
                "turnover_total_max": cfg.capacity_turnover_total_max,
                "slippage_bps_max": cfg.capacity_slippage_bps_max,
            },
        },
    )

    robust_oos = _dimension(scorecard, "robust_oos_performance")
    stress = _dimension(scorecard, "stress_resilience")
    regime_passes = [
        bool(robust_oos.get("pass", False)),
        bool(stress.get("pass", False)),
    ]
    pass_rate = float(np.mean(np.asarray(regime_passes, dtype=float)))
    regime_gate = _bool_gate(
        "regime_robustness_minimums",
        bool(pass_rate >= cfg.regime_minimum_pass_rate),
        {
            "robust_oos_pass": regime_passes[0],
            "stress_resilience_pass": regime_passes[1],
            "pass_rate": pass_rate,
            "minimum_pass_rate": cfg.regime_minimum_pass_rate,
        },
    )

    return [calibration_gate, downside_gate, capacity_gate, regime_gate]


def promotion_blocked(gates: list[dict[str, Any]]) -> bool:
    return any(not bool(gate.get("pass", False)) for gate in gates)

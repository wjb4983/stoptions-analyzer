from __future__ import annotations

from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

import json
import numpy as np
from numpy.typing import NDArray


@dataclass(frozen=True)
class ArbitrageDiagnostics:
    calendar_violations: int
    butterfly_violations: int
    total_variance_monotonicity_violations: int
    max_calendar_violation: float
    max_butterfly_violation: float
    max_total_variance_monotonicity_violation: float

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class NoArbitrageRepairResult:
    repaired_total_variance: NDArray[np.float64]
    diagnostics: ArbitrageDiagnostics


def detect_calendar_arbitrage(
    total_variance: NDArray[np.float64],
    *,
    tolerance: float = 1e-8,
) -> tuple[int, float]:
    """Detect calendar arbitrage as decreases in total variance across tenor."""

    diffs = np.diff(np.asarray(total_variance, dtype=np.float64), axis=1)
    violations = np.maximum(-(diffs + tolerance), 0.0)
    return int(np.count_nonzero(violations > 0.0)), float(np.max(violations, initial=0.0))


def repair_calendar_arbitrage(
    total_variance: NDArray[np.float64],
) -> NDArray[np.float64]:
    repaired = np.asarray(total_variance, dtype=np.float64).copy()
    return np.maximum.accumulate(repaired, axis=1)


def detect_total_variance_monotonicity_violations(
    total_variance: NDArray[np.float64],
    *,
    tolerance: float = 1e-8,
) -> tuple[int, float]:
    return detect_calendar_arbitrage(total_variance, tolerance=tolerance)


def detect_butterfly_arbitrage(
    total_variance: NDArray[np.float64],
    moneyness: NDArray[np.float64],
    *,
    tolerance: float = 1e-8,
) -> tuple[int, float]:
    """Detect smile convexity violations on each tenor via piecewise-linear interpolation bounds."""

    w = np.asarray(total_variance, dtype=np.float64)
    k = np.asarray(moneyness, dtype=np.float64)
    if w.shape[0] != k.shape[0]:
        raise ValueError("moneyness length must match total_variance strike dimension")

    n_strikes, n_tenors = w.shape
    if n_strikes < 3:
        return 0, 0.0

    count = 0
    max_violation = 0.0
    for t in range(n_tenors):
        for i in range(1, n_strikes - 1):
            left_w, mid_w, right_w = w[i - 1, t], w[i, t], w[i + 1, t]
            left_k, mid_k, right_k = k[i - 1], k[i], k[i + 1]
            if right_k <= left_k:
                continue
            weight = (mid_k - left_k) / (right_k - left_k)
            convex_upper = (1.0 - weight) * left_w + weight * right_w
            violation = mid_w - convex_upper - tolerance
            if violation > 0.0:
                count += 1
                max_violation = max(max_violation, float(violation))
    return count, max_violation


def repair_butterfly_arbitrage(
    total_variance: NDArray[np.float64],
    moneyness: NDArray[np.float64],
    *,
    iterations: int = 8,
) -> NDArray[np.float64]:
    repaired = np.asarray(total_variance, dtype=np.float64).copy()
    k = np.asarray(moneyness, dtype=np.float64)

    if repaired.shape[0] != k.shape[0]:
        raise ValueError("moneyness length must match total_variance strike dimension")

    n_strikes, n_tenors = repaired.shape
    if n_strikes < 3:
        return repaired

    for _ in range(max(1, iterations)):
        changed = False
        for t in range(n_tenors):
            for i in range(1, n_strikes - 1):
                left_k, mid_k, right_k = k[i - 1], k[i], k[i + 1]
                if right_k <= left_k:
                    continue
                weight = (mid_k - left_k) / (right_k - left_k)
                convex_upper = (1.0 - weight) * repaired[i - 1, t] + weight * repaired[i + 1, t]
                if repaired[i, t] > convex_upper:
                    repaired[i, t] = convex_upper
                    changed = True
        if not changed:
            break

    return repaired


def detect_and_repair_no_arb(
    total_variance: NDArray[np.float64],
    moneyness: NDArray[np.float64],
    *,
    tolerance: float = 1e-8,
) -> NoArbitrageRepairResult:
    repaired = repair_calendar_arbitrage(total_variance)
    repaired = repair_butterfly_arbitrage(repaired, moneyness)
    repaired = repair_calendar_arbitrage(repaired)

    calendar_count, calendar_mag = detect_calendar_arbitrage(repaired, tolerance=tolerance)
    butterfly_count, butterfly_mag = detect_butterfly_arbitrage(repaired, moneyness, tolerance=tolerance)
    mono_count, mono_mag = detect_total_variance_monotonicity_violations(repaired, tolerance=tolerance)

    diagnostics = ArbitrageDiagnostics(
        calendar_violations=calendar_count,
        butterfly_violations=butterfly_count,
        total_variance_monotonicity_violations=mono_count,
        max_calendar_violation=calendar_mag,
        max_butterfly_violation=butterfly_mag,
        max_total_variance_monotonicity_violation=mono_mag,
    )
    return NoArbitrageRepairResult(repaired_total_variance=repaired, diagnostics=diagnostics)


def export_no_arb_diagnostics(
    diagnostics: ArbitrageDiagnostics,
    *,
    out_path: str | Path,
    model_gate_threshold: int = 0,
) -> dict[str, Any]:
    payload = {
        "model_gate": {
            "pass": bool(
                diagnostics.calendar_violations <= model_gate_threshold
                and diagnostics.butterfly_violations <= model_gate_threshold
                and diagnostics.total_variance_monotonicity_violations <= model_gate_threshold
            ),
            "threshold": model_gate_threshold,
        },
        "diagnostics": diagnostics.to_dict(),
    }
    path = Path(out_path)
    path.write_text(json.dumps(payload, indent=2) + "\n")
    return payload

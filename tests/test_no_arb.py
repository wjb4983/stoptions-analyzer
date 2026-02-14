from __future__ import annotations

import numpy as np

from src.modeling_nextgen.features.no_arb import (
    detect_and_repair_no_arb,
    detect_butterfly_arbitrage,
    detect_calendar_arbitrage,
)


def test_detects_calendar_and_butterfly_violations() -> None:
    moneyness = np.array([0.9, 1.0, 1.1], dtype=np.float64)
    total_variance = np.array(
        [
            [0.10, 0.20, 0.30],
            [0.20, 0.10, 0.32],
            [0.10, 0.21, 0.31],
        ],
        dtype=np.float64,
    )

    calendar_count, _ = detect_calendar_arbitrage(total_variance)
    butterfly_count, _ = detect_butterfly_arbitrage(total_variance, moneyness)

    assert calendar_count > 0
    assert butterfly_count > 0


def test_repair_eliminates_violations_for_gate() -> None:
    moneyness = np.array([0.9, 1.0, 1.1], dtype=np.float64)
    total_variance = np.array(
        [
            [0.10, 0.20, 0.30],
            [0.20, 0.10, 0.32],
            [0.10, 0.21, 0.31],
        ],
        dtype=np.float64,
    )

    result = detect_and_repair_no_arb(total_variance, moneyness)

    assert result.diagnostics.calendar_violations == 0
    assert result.diagnostics.butterfly_violations == 0
    assert result.diagnostics.total_variance_monotonicity_violations == 0

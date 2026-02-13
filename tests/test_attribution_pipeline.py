from __future__ import annotations

import numpy as np
import pytest

from src.analysis.attribution import build_attribution_payload


def test_build_attribution_payload_known_components() -> None:
    timestamps = np.array([1, 2, 3], dtype=np.int64)
    prices = np.array(
        [
            [100.0, 100.0],
            [110.0, 90.0],
            [121.0, 81.0],
        ],
        dtype=float,
    )
    positions = np.array(
        [
            [0.0, 0.0],
            [1.0, -1.0],
            [1.0, -1.0],
        ],
        dtype=float,
    )

    payload = build_attribution_payload(
        timestamps=timestamps,
        prices=prices,
        positions=positions,
        slippage_drag=np.array([0.0, 0.01, 0.01]),
        fee_drag=np.array([0.0, 0.005, 0.005]),
        borrow_drag=np.array([0.0, 0.002, 0.002]),
    )

    by_variant = {}
    for row in payload.time_series:
        by_variant.setdefault(row["variant"], []).append(row)

    cross_brinson = by_variant["brinson_cross_sectional"]
    assert cross_brinson[1]["gross_alpha"] == pytest.approx(0.1)
    assert cross_brinson[1]["cost_drag"] == pytest.approx(0.017)
    assert cross_brinson[1]["borrow_drag"] == pytest.approx(0.002)
    assert cross_brinson[1]["residual_unexplained"] == pytest.approx(0.0)

    factor_ts = by_variant["factor_time_series"]
    assert factor_ts[1]["explained_component"] == pytest.approx(0.09996, abs=1e-4)
    assert factor_ts[1]["residual_unexplained"] == pytest.approx(0.0, abs=1e-4)

    summary = {row["variant"]: row for row in payload.summary}
    assert summary["brinson_time_series"]["gross_alpha_total"] == pytest.approx(0.2)
    assert summary["brinson_time_series"]["cost_drag_total"] == pytest.approx(0.034)
    assert summary["brinson_time_series"]["net_alpha_total"] == pytest.approx(0.166)

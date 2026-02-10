from __future__ import annotations

import numpy as np


def deterministic_single_asset_case() -> tuple[np.ndarray, np.ndarray]:
    """Simple deterministic price/signal arrays with non-trivial turnover."""

    prices = np.array([100.0, 101.0, 103.0, 102.0, 104.0], dtype=float)
    signals = np.array([0.0, 1.0, 1.0, -1.0, 0.0], dtype=float)
    return prices, signals


def deterministic_multi_asset_case() -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    """Two-asset deterministic scenario for trade/position invariants."""

    prices = np.array(
        [
            [100.0, 50.0],
            [101.0, 49.0],
            [102.0, 50.0],
            [100.0, 52.0],
        ],
        dtype=float,
    )
    signals = np.array(
        [
            [1.0, 0.0],
            [1.0, -1.0],
            [0.0, -1.0],
            [0.0, 0.0],
        ],
        dtype=float,
    )
    weights = np.array([0.7, 0.3], dtype=float)
    return prices, signals, weights


def deterministic_missing_bar_case() -> tuple[np.ndarray, np.ndarray, list[str]]:
    """Discontinuous timestamps to mimic missing bars in the raw feed."""

    timestamps = ["2024-01-02", "2024-01-03", "2024-01-05", "2024-01-08"]
    prices = np.array([100.0, 101.0, 100.5, 102.0], dtype=float)
    signals = np.array([0.0, 1.0, 1.0, 0.0], dtype=float)
    return prices, signals, timestamps


def deterministic_flat_prices_case() -> tuple[np.ndarray, np.ndarray]:
    prices = np.array([50.0, 50.0, 50.0, 50.0], dtype=float)
    signals = np.array([0.0, 1.0, -1.0, 0.0], dtype=float)
    return prices, signals


def deterministic_zero_volume_case() -> tuple[np.ndarray, np.ndarray]:
    """Second asset is illiquid: strategy never signals exposure to it."""

    prices = np.array(
        [
            [100.0, 10.0],
            [101.0, 10.0],
            [102.0, 10.0],
            [101.0, 10.0],
        ],
        dtype=float,
    )
    signals = np.array(
        [
            [0.0, 0.0],
            [1.0, 0.0],
            [1.0, 0.0],
            [0.0, 0.0],
        ],
        dtype=float,
    )
    return prices, signals


def deterministic_event_bars() -> list[dict[str, float | str]]:
    return [
        {"timestamp": "2024-01-02", "open": 100.0, "volume": 1_000.0},
        {"timestamp": "2024-01-03", "open": 100.0, "volume": 900.0},
        {"timestamp": "2024-01-04", "open": 100.0, "volume": 0.0},
        {"timestamp": "2024-01-05", "open": 100.0, "volume": 1_100.0},
    ]


def deterministic_event_target_positions() -> np.ndarray:
    """Target position for each bar close."""

    return np.array([0.0, 1.0, -1.0, 0.0], dtype=float)

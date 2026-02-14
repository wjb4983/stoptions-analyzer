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


def pit_membership_fixture() -> dict[str, np.ndarray]:
    """Fixture for point-in-time universe inclusion/exclusion."""

    timestamps = np.array([
        1704205800000,
        1704205860000,
        1704205920000,
        1704205980000,
    ], dtype=np.int64)
    return {
        "t": timestamps,
        "o": np.array([100.0, 101.0, 102.0, 103.0], dtype=float),
        "c": np.array([100.2, 101.2, 102.2, 103.2], dtype=float),
        "active_from": np.array([timestamps[1]], dtype=np.int64),
        "active_to": np.array([timestamps[2]], dtype=np.int64),
        "tradable": np.array([True, True, True, True], dtype=bool),
    }


def delisting_fixture() -> dict[str, np.ndarray]:
    """Fixture containing terminal bar and delisting scenario."""

    timestamps = np.array([
        1704205800000,
        1704205860000,
        1704205920000,
    ], dtype=np.int64)
    return {
        "t": timestamps,
        "o": np.array([50.0, 49.0, 48.0], dtype=float),
        "c": np.array([49.5, 48.5, 48.0], dtype=float),
        "tradable": np.array([True, True, True], dtype=bool),
    }



def synthetic_option_snapshots_dataset() -> list[dict[str, object]]:
    """Representative options payloads used across data-access contract tests."""

    return [
        {
            "details": {
                "ticker": "AAPL240119C00100000",
                "expiration_date": "2024-01-19",
                "contract_type": "call",
                "strike_price": 100.0,
            },
            "greeks": {"delta": 0.62, "gamma": 0.03, "theta": -0.08, "vega": 0.11, "rho": 0.04, "iv": 0.24},
            "day": {"close": 5.1, "volume": 1200},
            "last_quote": {"bid": 5.0, "ask": 5.2},
            "last_trade": {"price": 5.15},
            "open_interest": 800,
        },
        {
            "details": {
                "ticker": "AAPL240119P00095000",
                "expiration_date": "2024-01-19",
                "contract_type": "put",
                "strike_price": 95.0,
            },
            "greeks": {"delta": -0.38, "gamma": 0.02, "theta": -0.06, "vega": 0.09, "rho": -0.03, "iv": 0.27},
            "day": {"close": 2.8, "volume": 900},
            "last_quote": {"bid": 2.7, "ask": 2.9},
            "last_trade": {"price": 2.8},
            "open_interest": 950,
        },
    ]


def synthetic_vendor_bars_dataset() -> list[dict[str, object]]:
    """Minute bars in vendor payload shape expected by coerce_vendor_bar."""

    return [
        {"t": 1704205800000, "o": 100.0, "h": 101.0, "l": 99.5, "c": 100.5, "v": 1000.0, "n": 10},
        {"t": 1704205860000, "o": 100.5, "h": 101.3, "l": 100.0, "c": 101.0, "v": 1050.0, "n": 11},
        {"t": 1704205920000, "o": 101.0, "h": 101.8, "l": 100.7, "c": 101.5, "v": 1100.0, "n": 12},
    ]


def synthetic_corporate_actions_dataset() -> list[dict[str, object]]:
    return [
        {"action_type": "dividend", "action_date": "2024-01-15", "value": 0.25, "currency": "USD"},
        {"action_type": "split", "action_date": "2024-02-01", "value": 2.0, "ratio": 2.0, "currency": "USD"},
    ]

"""Phase A usage example: load bars, normalize, and run a vectorized backtest."""

from __future__ import annotations

from datetime import datetime, timedelta, timezone
from pathlib import Path

import numpy as np

from backtesting.vectorized import BpsSlippage, backtest_vectorized
from data_access.normalization import NormalizationConfig, normalize_bars
from data_access.providers import MassiveCacheProvider


def main() -> None:
    symbol = "NVDA"
    cache_root = Path("src/data/backtest_cache")

    provider = MassiveCacheProvider(cache_root=cache_root, symbols=[symbol])

    # Massive cache provider only supports 1-minute bars.
    bars = provider.get_bars(
        [symbol],
        start=datetime(2024, 1, 2, tzinfo=timezone.utc),
        end=datetime(2024, 1, 3, tzinfo=timezone.utc),
        timeframe="1m",
    )

    # provider.get_bars may return a DataFrame or list of dicts depending on pandas.
    if hasattr(bars, "to_dict"):
        records = bars.to_dict("records")
    else:
        records = list(bars)

    normalized = normalize_bars(
        records,
        NormalizationConfig(
            vendor_timezone="UTC",
            expected_interval=timedelta(minutes=1),
            missing_bar_policy="drop",
            adjustment_mode="raw",
        ),
    )

    if not normalized:
        print("No bars loaded. Ensure the cache root has data for the symbol/time range.")
        return

    closes = np.array([bar["close"] for bar in normalized], dtype=float)

    # Expected shapes:
    # - closes: (N,)
    # - signals: (N,)
    signals = np.zeros_like(closes)
    signals[1:] = np.where(closes[1:] > closes[:-1], 1.0, -1.0)

    result = backtest_vectorized(
        closes,
        signals,
        slippage_model=BpsSlippage(bps=1.0),
        initial_equity=1.0,
    )

    # Expected output shapes (aligned with closes):
    # - result.equity_curve: (N,)
    # - result.positions: (N,)
    # - result.returns: (N,)
    # - result.pnl: (N,)
    print("Phase A metrics:", result.metrics)


if __name__ == "__main__":
    main()

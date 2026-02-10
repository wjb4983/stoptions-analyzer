"""Phase B usage example: normalized bars + vectorized backtest on a trivial signal."""

from __future__ import annotations

from datetime import datetime, timedelta, timezone
from pathlib import Path

import numpy as np

from backtesting.vectorized import backtest_vectorized
from data_access.normalization import NormalizationConfig, normalize_bars
from data_access.providers import MassiveCacheProvider


def main() -> None:
    symbol = "NVDA"
    cache_root = Path("src/data/backtest_cache")

    provider = MassiveCacheProvider(cache_root=cache_root, symbols=[symbol])

    bars = provider.get_bars(
        [symbol],
        start=datetime(2024, 2, 1, tzinfo=timezone.utc),
        end=datetime(2024, 2, 2, tzinfo=timezone.utc),
        timeframe="1m",
    )

    if hasattr(bars, "to_dict"):
        records = bars.to_dict("records")
    else:
        records = list(bars)

    normalized = normalize_bars(
        records,
        NormalizationConfig(
            vendor_timezone="UTC",
            expected_interval=timedelta(minutes=1),
            missing_bar_policy="ffill",
            adjustment_mode="total-return",
        ),
    )

    if not normalized:
        print("No bars loaded. Ensure the cache root has data for the symbol/time range.")
        return

    closes = np.array([bar["close"] for bar in normalized], dtype=float)

    # Assumptions:
    # - adjustment_mode="total-return" only has an effect if split/dividend fields exist.
    # - signals are evaluated on close and executed on the next bar open.
    signals = np.ones_like(closes)

    # Expected shapes:
    # - closes: (N,)
    # - signals: (N,)
    result = backtest_vectorized(closes, signals)

    # Expected output shapes (aligned with closes):
    # - result.equity_curve: (N,)
    # - result.positions: (N,)
    # - result.returns: (N,)
    # - result.pnl: (N,)
    print("Phase B metrics:", result.metrics)


if __name__ == "__main__":
    main()

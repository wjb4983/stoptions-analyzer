from __future__ import annotations

import numpy as np

from .config import EntrySignalConfig, ExitSignalConfig
from .entry import build_entry_signal
from .exit import PositionState, build_exit_signal


def build_targets(
    *,
    close_prices: np.ndarray,
    missing_mask: np.ndarray,
    entry_config: EntrySignalConfig,
    exit_config: ExitSignalConfig,
) -> np.ndarray:
    n_periods, n_assets = close_prices.shape
    signals = np.zeros((n_periods, n_assets), dtype=float)
    entry_signal = build_entry_signal(entry_config)
    exit_signal = build_exit_signal(exit_config)

    for asset_idx in range(n_assets):
        prices = close_prices[:, asset_idx]
        missing = missing_mask[:, asset_idx]
        side = 0
        state: PositionState | None = None

        for idx in range(n_periods):
            if missing[idx]:
                signals[idx, asset_idx] = float(side)
                continue

            candidate = entry_signal.value_at(idx, prices, missing)

            if state is not None:
                state.bars_held += 1
                if state.side > 0:
                    state.peak_price = max(state.peak_price, float(prices[idx]))
                else:
                    state.peak_price = min(state.peak_price, float(prices[idx]))

            if side != 0 and state is not None:
                should_exit = exit_signal.should_exit(idx, prices, missing, state)
                if candidate == 0 or candidate == -side:
                    should_exit = True
                if should_exit:
                    side = 0
                    state = None

            if side == 0 and candidate != 0:
                side = int(candidate)
                state = PositionState(
                    side=side,
                    entry_price=float(prices[idx]),
                    peak_price=float(prices[idx]),
                    bars_held=0,
                )

            signals[idx, asset_idx] = float(side)

    return signals

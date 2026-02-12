from __future__ import annotations

from dataclasses import dataclass

import numpy as np

from .config import EntrySignalConfig, ExitSignalConfig, required_lookback_window
from .entry import build_entry_signal
from .exit import PositionState, build_exit_signal


@dataclass(frozen=True)
class StandardSignalMetadata:
    confidence: float
    horizon_bars: int


@dataclass(frozen=True)
class StandardSignalPoint:
    value: float
    metadata: StandardSignalMetadata


@dataclass(frozen=True)
class StandardSignalOutput:
    values: np.ndarray
    confidence: np.ndarray
    horizon_bars: np.ndarray


def _standardize_signal(candidate: int, horizon_bars: int) -> StandardSignalPoint:
    return StandardSignalPoint(
        value=float(candidate),
        metadata=StandardSignalMetadata(confidence=float(min(1.0, abs(candidate))), horizon_bars=horizon_bars),
    )


def build_standardized_targets(
    *,
    close_prices: np.ndarray,
    missing_mask: np.ndarray,
    entry_config: EntrySignalConfig,
    exit_config: ExitSignalConfig,
) -> StandardSignalOutput:
    n_periods, n_assets = close_prices.shape
    values = np.zeros((n_periods, n_assets), dtype=float)
    confidence = np.zeros((n_periods, n_assets), dtype=float)
    horizon = np.zeros((n_periods, n_assets), dtype=int)
    entry_signal = build_entry_signal(entry_config)
    exit_signal = build_exit_signal(exit_config)
    default_horizon = required_lookback_window(entry_config, exit_config)

    for asset_idx in range(n_assets):
        prices = close_prices[:, asset_idx]
        missing = missing_mask[:, asset_idx]
        side = 0
        state: PositionState | None = None

        for idx in range(n_periods):
            if missing[idx]:
                values[idx, asset_idx] = float(side)
                confidence[idx, asset_idx] = 0.0
                horizon[idx, asset_idx] = 0
                continue

            candidate = entry_signal.value_at(idx, prices, missing)
            standardized = _standardize_signal(candidate, default_horizon)

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

            values[idx, asset_idx] = float(side)
            confidence[idx, asset_idx] = standardized.metadata.confidence
            horizon[idx, asset_idx] = standardized.metadata.horizon_bars

    return StandardSignalOutput(values=values, confidence=confidence, horizon_bars=horizon)


def build_targets(
    *,
    close_prices: np.ndarray,
    missing_mask: np.ndarray,
    entry_config: EntrySignalConfig,
    exit_config: ExitSignalConfig,
) -> np.ndarray:
    return build_standardized_targets(
        close_prices=close_prices,
        missing_mask=missing_mask,
        entry_config=entry_config,
        exit_config=exit_config,
    ).values

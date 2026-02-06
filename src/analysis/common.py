from __future__ import annotations

import math
from typing import Iterable

import numpy as np


def is_valid_price(value: object) -> bool:
    return isinstance(value, (int, float)) and value > 0


def extract_close_volume(series: Iterable[object]) -> tuple[list[float], list[float]]:
    closes: list[float] = []
    volumes: list[float] = []
    for item in series:
        if isinstance(item, dict):
            close_value = item.get("close")
            volume_value = item.get("volume")
        else:
            close_value = item
            volume_value = None
        if isinstance(close_value, (int, float)):
            closes.append(float(close_value))
            if isinstance(volume_value, (int, float)):
                volumes.append(float(volume_value))
    return closes, volumes


def estimate_volatility(closes: list[float]) -> float | None:
    if len(closes) < 2:
        return None
    returns = np.diff(np.log(np.asarray(closes, dtype=float)))
    if returns.size == 0:
        return None
    return float(np.std(returns, ddof=1))


def multi_horizon_return(closes: list[float], end_index: int, windows: Iterable[int]) -> float | None:
    values: list[float] = []
    for window in windows:
        start_index = end_index - window
        if start_index < 0:
            continue
        start_price = closes[start_index]
        end_price = closes[end_index]
        if not is_valid_price(start_price) or not is_valid_price(end_price):
            continue
        values.append((end_price / start_price) - 1.0)
    if not values:
        return None
    return float(np.mean(values))


def quantile_bucket_sizes(total: int, top_quantile: float, bottom_quantile: float) -> tuple[int, int]:
    if total <= 0:
        return 0, 0
    top = max(1, int(math.ceil(total * top_quantile))) if top_quantile > 0 else 0
    bottom = max(1, int(math.ceil(total * bottom_quantile))) if bottom_quantile > 0 else 0
    if top + bottom > total:
        bottom = max(0, total - top)
    return top, bottom

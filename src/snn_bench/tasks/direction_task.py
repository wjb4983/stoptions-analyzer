from __future__ import annotations

import numpy as np


class DirectionClassificationTask:
    """Binary task: predict whether next-bar return is positive."""

    def make_dataset(self, features: np.ndarray, bars: dict) -> tuple[np.ndarray, np.ndarray]:
        close = np.asarray(bars.get("c", np.array([])), dtype=float)
        if close.size < 2 or features.size == 0:
            width = features.shape[1] if features.ndim == 2 else 0
            return np.empty((0, width)), np.array([], dtype=int)

        future_ret = np.diff(close) / np.where(close[:-1] == 0, 1.0, close[:-1])
        y = (future_ret > 0).astype(int)
        x = features[:-1]
        min_len = min(len(x), len(y))
        return x[:min_len], y[:min_len]

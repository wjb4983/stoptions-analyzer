from __future__ import annotations

import numpy as np


class BasicFeaturePipeline:
    """Simple numpy-based feature pipeline suitable for smoke tests."""

    def transform(self, bars: dict) -> np.ndarray:
        close = np.asarray(bars.get("c", np.array([])), dtype=float)
        if close.size == 0:
            return np.empty((0, 3), dtype=float)

        ret_1 = np.zeros_like(close)
        ret_5 = np.zeros_like(close)
        vol_10 = np.zeros_like(close)

        ret_1[1:] = np.diff(close) / np.where(close[:-1] == 0, 1.0, close[:-1])
        ret_5[5:] = (close[5:] - close[:-5]) / np.where(close[:-5] == 0, 1.0, close[:-5])

        for i in range(10, len(close)):
            window = ret_1[i - 9 : i + 1]
            vol_10[i] = float(np.std(window))

        return np.column_stack([ret_1, ret_5, vol_10])

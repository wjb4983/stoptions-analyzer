from __future__ import annotations

import numpy as np


def classification_metrics(y_true: np.ndarray, y_pred: np.ndarray) -> dict:
    if y_true.size == 0:
        return {"accuracy": 0.0, "count": 0}
    acc = float((y_true == y_pred).mean())
    return {"accuracy": acc, "count": int(y_true.size)}

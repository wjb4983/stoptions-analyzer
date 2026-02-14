from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import numpy as np

from backtesting.strategies.alpha_model import probability_calibrated_position_size
from backtesting.walk_forward import WalkForwardFold, build_walk_forward_folds


@dataclass(frozen=True)
class BacktestingBridge:
    """Adapter exposing next-gen validation hooks into existing backtesting APIs."""

    def build_walk_forward_folds(self, **kwargs: Any) -> list[WalkForwardFold]:
        return build_walk_forward_folds(**kwargs)

    def interval_aware_position_size(
        self,
        probabilities: np.ndarray,
        *,
        confidence_lower: np.ndarray | None = None,
        confidence_upper: np.ndarray | None = None,
        predictive_std: np.ndarray | None = None,
        neutral_probability: float = 0.5,
        max_leverage: float = 1.0,
        gamma: float = 1.0,
        interval_aversion: float = 1.0,
        min_scale: float = 0.05,
    ) -> np.ndarray:
        """Risk-throttle position sizes with posterior interval information.

        Narrow intervals preserve more base size; wider intervals shrink exposures.
        """

        base_size = probability_calibrated_position_size(
            probabilities,
            neutral_probability=neutral_probability,
            max_leverage=max_leverage,
            gamma=gamma,
        )

        probs = np.asarray(probabilities, dtype=float)
        scale = np.ones_like(probs, dtype=float)
        risk_aversion = max(float(interval_aversion), 0.0)

        if confidence_lower is not None and confidence_upper is not None:
            lower = np.asarray(confidence_lower, dtype=float)
            upper = np.asarray(confidence_upper, dtype=float)
            if lower.shape != probs.shape or upper.shape != probs.shape:
                raise ValueError("confidence bounds must match probabilities shape")
            interval_width = np.clip(upper - lower, 0.0, None)
            confidence_score = 1.0 - np.clip(interval_width, 0.0, 1.0)
            scale = np.power(confidence_score, risk_aversion)
        elif predictive_std is not None:
            sigma = np.asarray(predictive_std, dtype=float)
            if sigma.shape != probs.shape:
                raise ValueError("predictive_std must match probabilities shape")
            scale = np.exp(-risk_aversion * np.clip(sigma, 0.0, None))

        scale = np.clip(scale, float(min_scale), 1.0)
        return base_size * scale

from __future__ import annotations

from dataclasses import dataclass

import numpy as np


_Z_SCORES = {
    0.8: 1.2815515655446004,
    0.85: 1.4395314709384563,
    0.9: 1.6448536269514722,
    0.95: 1.959963984540054,
    0.975: 2.241402727604947,
    0.99: 2.5758293035489004,
}


@dataclass(frozen=True)
class BayesianUncertaintyEstimate:
    """Posterior summary for predictive distributions."""

    posterior_mean: np.ndarray
    confidence_lower: np.ndarray
    confidence_upper: np.ndarray
    total_std: np.ndarray
    epistemic_std: np.ndarray
    aleatoric_std: np.ndarray


class BayesianUncertaintyCalibrator:
    """Calibrate and aggregate predictive samples into posterior uncertainty bands.

    Supports two common approximations:
    - `bma`: Bayesian model averaging over an ensemble of model samples.
    - `mc_dropout`: MC-dropout style sampling treated as draws from posterior weights.
    """

    name = "bayesian_uncertainty"

    def __init__(self, mode: str = "bma", confidence_level: float = 0.9) -> None:
        mode_value = str(mode).strip().lower()
        if mode_value not in {"bma", "mc_dropout"}:
            raise ValueError("mode must be one of {'bma', 'mc_dropout'}")
        if not 0.0 < float(confidence_level) < 1.0:
            raise ValueError("confidence_level must be in (0, 1)")

        self.mode = mode_value
        self.confidence_level = float(confidence_level)
        self._weights: np.ndarray | None = None
        self._aleatoric_scale = 1.0

    @staticmethod
    def _as_2d(samples: np.ndarray) -> np.ndarray:
        arr = np.asarray(samples, dtype=float)
        if arr.ndim == 1:
            return arr[None, :]
        if arr.ndim != 2:
            raise ValueError("samples must be shape (n_draws, n_observations)")
        return arr

    def fit(
        self,
        samples: np.ndarray,
        realized: np.ndarray | None = None,
        aleatoric_std: np.ndarray | None = None,
    ) -> None:
        draws = self._as_2d(samples)
        n_draws = draws.shape[0]

        if realized is not None:
            target = np.asarray(realized, dtype=float).reshape(-1)
            if target.shape[0] != draws.shape[1]:
                raise ValueError("realized must match n_observations")
            mse = np.mean((draws - target[None, :]) ** 2, axis=1)
            precision = 1.0 / np.clip(mse, 1e-12, None)
            self._weights = precision / np.sum(precision)
        else:
            self._weights = np.full(n_draws, 1.0 / max(n_draws, 1), dtype=float)

        if aleatoric_std is not None and realized is not None:
            aleatoric = np.asarray(aleatoric_std, dtype=float)
            expected_var = np.mean(np.square(np.clip(aleatoric, 1e-8, None)))
            residual_var = np.mean((target - np.mean(draws, axis=0)) ** 2)
            if expected_var > 0.0:
                self._aleatoric_scale = max(residual_var / expected_var, 1e-6) ** 0.5

    def transform(
        self,
        samples: np.ndarray,
        aleatoric_std: np.ndarray | None = None,
        confidence_level: float | None = None,
    ) -> BayesianUncertaintyEstimate:
        draws = self._as_2d(samples)
        weights = self._weights
        if weights is None or weights.shape[0] != draws.shape[0]:
            weights = np.full(draws.shape[0], 1.0 / max(draws.shape[0], 1), dtype=float)

        posterior_mean = np.average(draws, axis=0, weights=weights)
        centered = draws - posterior_mean[None, :]
        epistemic_var = np.average(centered**2, axis=0, weights=weights)

        if aleatoric_std is None:
            aleatoric_var = np.zeros_like(epistemic_var)
        else:
            aleatoric = np.asarray(aleatoric_std, dtype=float)
            if aleatoric.shape != posterior_mean.shape:
                raise ValueError("aleatoric_std must match posterior shape")
            aleatoric_var = np.square(np.clip(aleatoric * self._aleatoric_scale, 1e-8, None))

        total_std = np.sqrt(np.clip(epistemic_var + aleatoric_var, 0.0, None))

        level = float(confidence_level if confidence_level is not None else self.confidence_level)
        if not 0.0 < level < 1.0:
            raise ValueError("confidence_level must be in (0, 1)")
        z_value = _Z_SCORES.get(level)
        if z_value is None:
            # Rational approximation is unnecessary here; nearest tabulated level is stable for sizing.
            nearest_level = min(_Z_SCORES.keys(), key=lambda x: abs(x - level))
            z_value = _Z_SCORES[nearest_level]

        confidence_lower = posterior_mean - z_value * total_std
        confidence_upper = posterior_mean + z_value * total_std

        return BayesianUncertaintyEstimate(
            posterior_mean=posterior_mean,
            confidence_lower=confidence_lower,
            confidence_upper=confidence_upper,
            total_std=total_std,
            epistemic_std=np.sqrt(np.clip(epistemic_var, 0.0, None)),
            aleatoric_std=np.sqrt(np.clip(aleatoric_var, 0.0, None)),
        )

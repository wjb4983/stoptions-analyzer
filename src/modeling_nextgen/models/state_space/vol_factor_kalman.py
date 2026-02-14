from __future__ import annotations

from dataclasses import dataclass, field

import numpy as np
from numpy.typing import NDArray


@dataclass(frozen=True)
class VolFactorKalmanConfig:
    """Configuration for a 2-factor volatility state-space model.

    State ordering:
        x_t = [latent_vol_level, latent_vol_of_vol]
    """

    transition_matrix: NDArray[np.float64] = field(
        default_factory=lambda: np.array([[0.97, 0.05], [0.0, 0.90]], dtype=np.float64)
    )
    process_noise: NDArray[np.float64] = field(default_factory=lambda: np.array([0.015, 0.02], dtype=np.float64))
    measurement_noise: float = 0.05
    initial_mean: NDArray[np.float64] = field(default_factory=lambda: np.zeros(2, dtype=np.float64))
    initial_covariance: NDArray[np.float64] = field(default_factory=lambda: np.eye(2, dtype=np.float64))


@dataclass(frozen=True)
class VolFactorKalmanOutput:
    filtered_mean: NDArray[np.float64]
    filtered_covariance: NDArray[np.float64]
    smoothed_mean: NDArray[np.float64]
    smoothed_covariance: NDArray[np.float64]
    innovation: NDArray[np.float64]
    observation_matrix: NDArray[np.float64]

    def as_feature_dict(self, prefix: str = "vol_factor") -> dict[str, NDArray[np.float64]]:
        """Expose posterior factors as 1D arrays for ensemble-ready feature ingestion."""

        return {
            f"{prefix}_vol_level_filtered": self.filtered_mean[:, 0],
            f"{prefix}_vol_of_vol_filtered": self.filtered_mean[:, 1],
            f"{prefix}_vol_level_smoothed": self.smoothed_mean[:, 0],
            f"{prefix}_vol_of_vol_smoothed": self.smoothed_mean[:, 1],
            f"{prefix}_vol_level_var": self.smoothed_covariance[:, 0, 0],
            f"{prefix}_vol_of_vol_var": self.smoothed_covariance[:, 1, 1],
            f"{prefix}_cross_cov": self.smoothed_covariance[:, 0, 1],
            f"{prefix}_innovation_norm": np.linalg.norm(self.innovation, axis=1),
        }


def infer_observation_matrix(
    feature_names: list[str],
    *,
    n_factors: int = 2,
) -> NDArray[np.float64]:
    """Infer an observation loading matrix from common volatility-surface feature names."""

    if n_factors != 2:
        raise ValueError("Only 2-factor inference is supported: [vol_level, vol_of_vol].")

    h = np.zeros((len(feature_names), 2), dtype=np.float64)
    for idx, name in enumerate(feature_names):
        token = name.lower()

        if any(k in token for k in ("level", "atm", "mean_iv", "surface_mean")):
            h[idx, 0] += 1.0
        if any(k in token for k in ("slope", "skew", "term", "moneyness")):
            h[idx, 0] += 0.6
            h[idx, 1] += 0.35
        if any(k in token for k in ("curvature", "convexity", "butterfly", "kurtosis")):
            h[idx, 1] += 0.9
        if any(k in token for k in ("vol_of_vol", "vv", "volvol", "realized_vol_of_vol")):
            h[idx, 1] += 1.2
        if np.allclose(h[idx], 0.0):
            h[idx] = np.array([0.7, 0.3], dtype=np.float64)

    return h


def run_vol_factor_kalman(
    observations: NDArray[np.float64],
    *,
    feature_names: list[str] | None = None,
    observation_matrix: NDArray[np.float64] | None = None,
    config: VolFactorKalmanConfig | None = None,
) -> VolFactorKalmanOutput:
    """Run online Kalman filter and RTS smoother on volatility-surface feature observations.

    Args:
        observations: 2D array with shape (n_timestamps, n_features).
        feature_names: Optional names for observation features; used to infer loadings.
        observation_matrix: Optional explicit loading matrix H with shape (n_features, 2).
        config: Optional model configuration.
    """

    obs = np.asarray(observations, dtype=np.float64)
    if obs.ndim != 2:
        raise ValueError("observations must be 2D with shape (n_timestamps, n_features)")

    n_steps, n_features = obs.shape
    cfg = config or VolFactorKalmanConfig()

    if observation_matrix is not None:
        h = np.asarray(observation_matrix, dtype=np.float64)
    else:
        names = feature_names or [f"surface_feature_{i}" for i in range(n_features)]
        if len(names) != n_features:
            raise ValueError("feature_names length must equal observations.shape[1]")
        h = infer_observation_matrix(names)

    if h.shape != (n_features, 2):
        raise ValueError("observation_matrix must have shape (n_features, 2)")

    f = np.asarray(cfg.transition_matrix, dtype=np.float64)
    q = np.diag(np.asarray(cfg.process_noise, dtype=np.float64) ** 2)
    r_scalar = float(cfg.measurement_noise)

    filtered_mean = np.zeros((n_steps, 2), dtype=np.float64)
    filtered_cov = np.zeros((n_steps, 2, 2), dtype=np.float64)
    predicted_mean = np.zeros((n_steps, 2), dtype=np.float64)
    predicted_cov = np.zeros((n_steps, 2, 2), dtype=np.float64)
    innovation = np.zeros((n_steps, n_features), dtype=np.float64)

    x_prev = np.asarray(cfg.initial_mean, dtype=np.float64).copy()
    p_prev = np.asarray(cfg.initial_covariance, dtype=np.float64).copy()

    identity_state = np.eye(2, dtype=np.float64)

    for t in range(n_steps):
        x_pred = f @ x_prev
        p_pred = f @ p_prev @ f.T + q

        predicted_mean[t] = x_pred
        predicted_cov[t] = p_pred

        y_t = obs[t]
        valid = np.isfinite(y_t)

        if valid.any():
            y_obs = y_t[valid]
            h_obs = h[valid]

            innov = y_obs - (h_obs @ x_pred)
            innovation[t, valid] = innov

            r_t = np.eye(y_obs.shape[0], dtype=np.float64) * (r_scalar**2)
            s_t = h_obs @ p_pred @ h_obs.T + r_t
            k_t = p_pred @ h_obs.T @ np.linalg.pinv(s_t)

            x_upd = x_pred + k_t @ innov
            p_upd = (identity_state - k_t @ h_obs) @ p_pred
            p_upd = 0.5 * (p_upd + p_upd.T)
        else:
            x_upd = x_pred
            p_upd = p_pred

        filtered_mean[t] = x_upd
        filtered_cov[t] = p_upd

        x_prev = x_upd
        p_prev = p_upd

    smoothed_mean = filtered_mean.copy()
    smoothed_cov = filtered_cov.copy()

    for t in range(n_steps - 2, -1, -1):
        p_filt = filtered_cov[t]
        p_pred_next = predicted_cov[t + 1]

        gain = p_filt @ f.T @ np.linalg.pinv(p_pred_next)
        smoothed_mean[t] = filtered_mean[t] + gain @ (smoothed_mean[t + 1] - predicted_mean[t + 1])
        smoothed_cov[t] = p_filt + gain @ (smoothed_cov[t + 1] - p_pred_next) @ gain.T
        smoothed_cov[t] = 0.5 * (smoothed_cov[t] + smoothed_cov[t].T)

    return VolFactorKalmanOutput(
        filtered_mean=filtered_mean,
        filtered_covariance=filtered_cov,
        smoothed_mean=smoothed_mean,
        smoothed_covariance=smoothed_cov,
        innovation=innovation,
        observation_matrix=h,
    )


__all__ = [
    "VolFactorKalmanConfig",
    "VolFactorKalmanOutput",
    "infer_observation_matrix",
    "run_vol_factor_kalman",
]

from __future__ import annotations

from dataclasses import dataclass

import numpy as np
from numpy.typing import NDArray


@dataclass(frozen=True)
class RegimeSwitchingConfig:
    regime_names: tuple[str, ...] = ("calm", "stressed", "dislocated")
    em_iterations: int = 20
    eps: float = 1e-10


@dataclass(frozen=True)
class RegimeSwitchingOutput:
    regime_names: NDArray[np.object_]
    transition_matrix: NDArray[np.float64]
    posterior_probabilities: NDArray[np.float64]
    filtered_probabilities: NDArray[np.float64]
    regime_state_argmax: NDArray[np.int64]
    regime_labels: NDArray[np.object_]
    dates: NDArray[np.object_]

    def as_feature_dict(self, prefix: str = "markov_regime") -> dict[str, NDArray[np.float64]]:
        feature_map: dict[str, NDArray[np.float64]] = {
            f"{prefix}_state_argmax": self.regime_state_argmax.astype(np.float64),
        }
        for idx, regime in enumerate(self.regime_names.tolist()):
            feature_map[f"{prefix}_posterior_{regime}"] = self.posterior_probabilities[:, idx]
            feature_map[f"{prefix}_filtered_{regime}"] = self.filtered_probabilities[:, idx]
        return feature_map


def _normalize_rows(matrix: NDArray[np.float64], eps: float) -> NDArray[np.float64]:
    row_sum = np.maximum(np.sum(matrix, axis=1, keepdims=True), eps)
    return matrix / row_sum


def estimate_transition_matrix(
    posterior_probabilities: NDArray[np.float64],
    *,
    eps: float = 1e-10,
) -> NDArray[np.float64]:
    posterior = np.asarray(posterior_probabilities, dtype=np.float64)
    if posterior.ndim != 2:
        raise ValueError("posterior_probabilities must be a 2D array")

    n_obs, n_states = posterior.shape
    transition = np.full((n_states, n_states), eps, dtype=np.float64)
    for t in range(1, n_obs):
        transition += np.outer(posterior[t - 1], posterior[t])
    return _normalize_rows(transition, eps)


def _gaussian_log_likelihood(
    observations: NDArray[np.float64],
    means: NDArray[np.float64],
    variances: NDArray[np.float64],
) -> NDArray[np.float64]:
    n_obs = observations.shape[0]
    n_states = means.shape[0]
    log_like = np.zeros((n_obs, n_states), dtype=np.float64)
    for k in range(n_states):
        var = np.maximum(variances[k], 1e-5)
        diff = observations - means[k]
        log_like[:, k] = -0.5 * np.sum((diff * diff) / var + np.log(2.0 * np.pi * var), axis=1)
    return log_like


def _forward_backward(
    log_emission: NDArray[np.float64],
    transition: NDArray[np.float64],
    initial_probs: NDArray[np.float64],
    eps: float,
) -> tuple[NDArray[np.float64], NDArray[np.float64]]:
    n_obs, n_states = log_emission.shape
    emission = np.exp(log_emission - np.max(log_emission, axis=1, keepdims=True))
    emission = _normalize_rows(np.maximum(emission, eps), eps)

    forward = np.zeros_like(emission)
    forward[0] = initial_probs * emission[0]
    forward[0] /= np.maximum(np.sum(forward[0]), eps)

    for t in range(1, n_obs):
        forward[t] = emission[t] * (forward[t - 1] @ transition)
        forward[t] /= np.maximum(np.sum(forward[t]), eps)

    backward = np.ones_like(emission)
    for t in range(n_obs - 2, -1, -1):
        backward[t] = transition @ (emission[t + 1] * backward[t + 1])
        backward[t] /= np.maximum(np.sum(backward[t]), eps)

    posterior = forward * backward
    posterior = _normalize_rows(np.maximum(posterior, eps), eps)
    return posterior, forward


def fit_regime_switching_model(
    observations: NDArray[np.float64],
    *,
    dates: NDArray[np.object_] | list[object] | None = None,
    config: RegimeSwitchingConfig | None = None,
) -> RegimeSwitchingOutput:
    cfg = config or RegimeSwitchingConfig()
    x = np.asarray(observations, dtype=np.float64)
    if x.ndim != 2:
        raise ValueError("observations must be 2D with shape (n_dates, n_features)")

    n_obs, n_features = x.shape
    regime_names = np.asarray(cfg.regime_names, dtype=object)
    n_states = regime_names.size

    if n_obs == 0:
        return RegimeSwitchingOutput(
            regime_names=regime_names,
            transition_matrix=np.full((n_states, n_states), 1.0 / max(n_states, 1), dtype=np.float64),
            posterior_probabilities=np.zeros((0, n_states), dtype=np.float64),
            filtered_probabilities=np.zeros((0, n_states), dtype=np.float64),
            regime_state_argmax=np.zeros(0, dtype=np.int64),
            regime_labels=np.asarray([], dtype=object),
            dates=np.asarray([], dtype=object),
        )

    mean = np.nanmean(x, axis=0)
    std = np.nanstd(x, axis=0)
    std = np.where(std <= 1e-8, 1.0, std)
    z = (np.where(np.isfinite(x), x, mean) - mean) / std

    scores = np.mean(z, axis=1)
    quantiles = np.linspace(0.0, 1.0, n_states + 2)[1:-1]
    seeds = np.quantile(scores, quantiles)
    means = np.column_stack([seeds for _ in range(n_features)])
    variances = np.full((n_states, n_features), 1.0, dtype=np.float64)
    initial_probs = np.full(n_states, 1.0 / n_states, dtype=np.float64)
    transition = np.full((n_states, n_states), 1.0 / n_states, dtype=np.float64)

    posterior = np.full((n_obs, n_states), 1.0 / n_states, dtype=np.float64)
    filtered = posterior.copy()

    for _ in range(max(1, int(cfg.em_iterations))):
        log_emission = _gaussian_log_likelihood(z, means, variances)
        posterior, filtered = _forward_backward(log_emission, transition, initial_probs, cfg.eps)

        state_mass = np.maximum(np.sum(posterior, axis=0), cfg.eps)
        for k in range(n_states):
            weight = posterior[:, [k]]
            means[k] = np.sum(weight * z, axis=0) / state_mass[k]
            diff = z - means[k]
            variances[k] = np.maximum(np.sum(weight * (diff * diff), axis=0) / state_mass[k], 1e-5)

        initial_probs = np.maximum(posterior[0], cfg.eps)
        initial_probs /= np.sum(initial_probs)
        transition = estimate_transition_matrix(posterior, eps=cfg.eps)

    argmax = np.argmax(posterior, axis=1).astype(np.int64)
    labels = regime_names[argmax]

    if dates is None:
        date_index = np.asarray(np.arange(n_obs), dtype=object)
    else:
        supplied = np.asarray(dates, dtype=object).reshape(-1)
        if supplied.size != n_obs:
            raise ValueError("dates must have the same length as observations")
        date_index = supplied

    return RegimeSwitchingOutput(
        regime_names=regime_names,
        transition_matrix=transition,
        posterior_probabilities=posterior,
        filtered_probabilities=filtered,
        regime_state_argmax=argmax,
        regime_labels=labels,
        dates=date_index,
    )


__all__ = [
    "RegimeSwitchingConfig",
    "RegimeSwitchingOutput",
    "estimate_transition_matrix",
    "fit_regime_switching_model",
]

from __future__ import annotations

from dataclasses import dataclass
import math

import numpy as np
from numpy.typing import NDArray

from .regime_switching import _gaussian_log_likelihood, _normalize_rows, estimate_transition_matrix


@dataclass(frozen=True)
class SemiMarkovConfig:
    regime_names: tuple[str, ...] = ("calm", "stressed", "dislocated")
    em_iterations: int = 20
    max_duration: int = 20
    duration_prior_mean: float = 6.0
    duration_prior_strength: float = 2.0
    duration_feature_beta: float = 1.0
    eps: float = 1e-10


@dataclass(frozen=True)
class SemiMarkovOutput:
    regime_names: NDArray[np.object_]
    transition_matrix: NDArray[np.float64]
    duration_distributions: NDArray[np.float64]
    posterior_probabilities: NDArray[np.float64]
    filtered_probabilities: NDArray[np.float64]
    regime_state_argmax: NDArray[np.int64]
    regime_labels: NDArray[np.object_]
    duration_signal: NDArray[np.float64]
    dates: NDArray[np.object_]

    def as_feature_dict(self, prefix: str = "semi_markov_regime") -> dict[str, NDArray[np.float64]]:
        feature_map: dict[str, NDArray[np.float64]] = {
            f"{prefix}_state_argmax": self.regime_state_argmax.astype(np.float64),
            f"{prefix}_duration_signal": self.duration_signal,
        }
        for idx, regime in enumerate(self.regime_names.tolist()):
            feature_map[f"{prefix}_posterior_{regime}"] = self.posterior_probabilities[:, idx]
            feature_map[f"{prefix}_filtered_{regime}"] = self.filtered_probabilities[:, idx]
        return feature_map


def _compute_duration_signal(
    duration_features: NDArray[np.float64] | None,
    n_obs: int,
) -> NDArray[np.float64]:
    if duration_features is None:
        return np.zeros(n_obs, dtype=np.float64)

    values = np.asarray(duration_features, dtype=np.float64)
    if values.ndim == 1:
        signal = values
    elif values.ndim == 2:
        signal = np.mean(values, axis=1)
    else:
        raise ValueError("duration_features must be 1D or 2D")

    if signal.shape[0] != n_obs:
        raise ValueError("duration_features must have same number of rows as observations")

    signal = np.where(np.isfinite(signal), signal, np.nanmean(signal))
    std = np.nanstd(signal)
    if std <= 1e-8:
        return np.zeros(n_obs, dtype=np.float64)
    return np.clip((signal - np.nanmean(signal)) / std, -3.0, 3.0)


def _build_duration_prior(max_duration: int, mean_duration: float) -> NDArray[np.float64]:
    d = np.arange(1, max_duration + 1, dtype=np.float64)
    lam = max(mean_duration, 1.0)
    log_p = d * np.log(lam) - lam - np.array([math.lgamma(v + 1.0) for v in d])
    p = np.exp(log_p - np.max(log_p))
    return p / np.sum(p)


def _duration_pmf_from_path(
    path: NDArray[np.int64],
    n_states: int,
    max_duration: int,
    prior: NDArray[np.float64],
    prior_strength: float,
) -> NDArray[np.float64]:
    counts = np.tile(prior * max(prior_strength, 0.0), (n_states, 1))
    if path.size == 0:
        return _normalize_rows(np.maximum(counts, 1e-12), 1e-12)

    starts = np.r_[0, np.where(path[1:] != path[:-1])[0] + 1]
    ends = np.r_[starts[1:], path.size]
    states = path[starts]
    lengths = ends - starts

    for state, length in zip(states, lengths, strict=False):
        idx = int(min(max(length, 1), max_duration) - 1)
        counts[int(state), idx] += 1.0

    return _normalize_rows(np.maximum(counts, 1e-12), 1e-12)


def _hazard_from_duration_pmf(duration_pmf: NDArray[np.float64], eps: float) -> NDArray[np.float64]:
    n_states, max_duration = duration_pmf.shape
    hazard = np.zeros_like(duration_pmf)
    for state in range(n_states):
        survival = 1.0
        for d in range(max_duration):
            p = duration_pmf[state, d]
            hazard[state, d] = np.clip(p / max(survival, eps), eps, 1.0 - eps)
            survival = max(survival - p, eps)
    return hazard


def _forward_backward_semi_markov(
    log_emission: NDArray[np.float64],
    transition: NDArray[np.float64],
    initial_probs: NDArray[np.float64],
    hazard: NDArray[np.float64],
    duration_signal: NDArray[np.float64],
    duration_feature_beta: float,
    eps: float,
) -> tuple[NDArray[np.float64], NDArray[np.float64]]:
    n_obs, n_states = log_emission.shape
    max_duration = hazard.shape[1]

    emission = np.exp(log_emission - np.max(log_emission, axis=1, keepdims=True))
    emission = _normalize_rows(np.maximum(emission, eps), eps)

    expanded_size = n_states * max_duration
    forward = np.zeros((n_obs, expanded_size), dtype=np.float64)

    for k in range(n_states):
        forward[0, k * max_duration] = initial_probs[k] * emission[0, k]
    forward[0] /= np.maximum(np.sum(forward[0]), eps)

    switch_template = transition.copy()
    np.fill_diagonal(switch_template, 0.0)
    switch_template = _normalize_rows(np.maximum(switch_template, eps), eps)

    for t in range(1, n_obs):
        curr = np.zeros(expanded_size, dtype=np.float64)
        signal_adj = duration_feature_beta * duration_signal[t - 1]
        for k in range(n_states):
            for d in range(max_duration):
                idx = k * max_duration + d
                mass = forward[t - 1, idx]
                if mass <= eps:
                    continue

                base_h = hazard[k, d]
                logit_h = np.log(base_h / (1.0 - base_h))
                adjusted_h = 1.0 / (1.0 + np.exp(-(logit_h - signal_adj)))
                adjusted_h = np.clip(adjusted_h, eps, 1.0 - eps)
                stay_prob = 1.0 - adjusted_h

                stay_d = min(d + 1, max_duration - 1)
                curr[k * max_duration + stay_d] += mass * stay_prob * emission[t, k]

                switch_mass = mass * adjusted_h
                for j in range(n_states):
                    if j == k:
                        continue
                    curr[j * max_duration] += switch_mass * switch_template[k, j] * emission[t, j]

        forward[t] = curr / np.maximum(np.sum(curr), eps)

    backward = np.ones((n_obs, expanded_size), dtype=np.float64)
    for t in range(n_obs - 2, -1, -1):
        nxt = np.zeros(expanded_size, dtype=np.float64)
        signal_adj = duration_feature_beta * duration_signal[t]
        for k in range(n_states):
            for d in range(max_duration):
                idx = k * max_duration + d
                base_h = hazard[k, d]
                logit_h = np.log(base_h / (1.0 - base_h))
                adjusted_h = 1.0 / (1.0 + np.exp(-(logit_h - signal_adj)))
                adjusted_h = np.clip(adjusted_h, eps, 1.0 - eps)
                stay_prob = 1.0 - adjusted_h

                stay_d = min(d + 1, max_duration - 1)
                total = stay_prob * emission[t + 1, k] * backward[t + 1, k * max_duration + stay_d]
                for j in range(n_states):
                    if j == k:
                        continue
                    total += adjusted_h * switch_template[k, j] * emission[t + 1, j] * backward[
                        t + 1, j * max_duration
                    ]
                nxt[idx] = total

        backward[t] = nxt / np.maximum(np.sum(nxt), eps)

    smoothed = np.maximum(forward * backward, eps)
    smoothed = _normalize_rows(smoothed, eps)

    posterior = np.zeros((n_obs, n_states), dtype=np.float64)
    filtered = np.zeros((n_obs, n_states), dtype=np.float64)
    for k in range(n_states):
        start = k * max_duration
        stop = (k + 1) * max_duration
        posterior[:, k] = np.sum(smoothed[:, start:stop], axis=1)
        filtered[:, k] = np.sum(forward[:, start:stop], axis=1)

    posterior = _normalize_rows(np.maximum(posterior, eps), eps)
    filtered = _normalize_rows(np.maximum(filtered, eps), eps)
    return posterior, filtered


def fit_semi_markov_model(
    observations: NDArray[np.float64],
    *,
    duration_features: NDArray[np.float64] | None = None,
    dates: NDArray[np.object_] | list[object] | None = None,
    config: SemiMarkovConfig | None = None,
) -> SemiMarkovOutput:
    cfg = config or SemiMarkovConfig()
    x = np.asarray(observations, dtype=np.float64)
    if x.ndim != 2:
        raise ValueError("observations must be 2D with shape (n_dates, n_features)")

    n_obs, n_features = x.shape
    regime_names = np.asarray(cfg.regime_names, dtype=object)
    n_states = regime_names.size
    max_duration = max(int(cfg.max_duration), 2)

    if n_obs == 0:
        return SemiMarkovOutput(
            regime_names=regime_names,
            transition_matrix=np.full((n_states, n_states), 1.0 / max(n_states, 1), dtype=np.float64),
            duration_distributions=np.full((n_states, max_duration), 1.0 / max_duration, dtype=np.float64),
            posterior_probabilities=np.zeros((0, n_states), dtype=np.float64),
            filtered_probabilities=np.zeros((0, n_states), dtype=np.float64),
            regime_state_argmax=np.zeros(0, dtype=np.int64),
            regime_labels=np.asarray([], dtype=object),
            duration_signal=np.zeros(0, dtype=np.float64),
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

    duration_signal = _compute_duration_signal(duration_features, n_obs)
    duration_prior = _build_duration_prior(max_duration, cfg.duration_prior_mean)
    duration_pmf = np.tile(duration_prior, (n_states, 1))

    posterior = np.full((n_obs, n_states), 1.0 / n_states, dtype=np.float64)
    filtered = posterior.copy()

    for _ in range(max(1, int(cfg.em_iterations))):
        hazard = _hazard_from_duration_pmf(duration_pmf, cfg.eps)
        log_emission = _gaussian_log_likelihood(z, means, variances)
        posterior, filtered = _forward_backward_semi_markov(
            log_emission,
            transition,
            initial_probs,
            hazard,
            duration_signal,
            cfg.duration_feature_beta,
            cfg.eps,
        )

        state_mass = np.maximum(np.sum(posterior, axis=0), cfg.eps)
        for k in range(n_states):
            weight = posterior[:, [k]]
            means[k] = np.sum(weight * z, axis=0) / state_mass[k]
            diff = z - means[k]
            variances[k] = np.maximum(np.sum(weight * (diff * diff), axis=0) / state_mass[k], 1e-5)

        initial_probs = np.maximum(posterior[0], cfg.eps)
        initial_probs /= np.sum(initial_probs)
        transition = estimate_transition_matrix(posterior, eps=cfg.eps)

        path = np.argmax(posterior, axis=1).astype(np.int64)
        duration_pmf = _duration_pmf_from_path(
            path,
            n_states,
            max_duration,
            duration_prior,
            cfg.duration_prior_strength,
        )

    argmax = np.argmax(posterior, axis=1).astype(np.int64)
    labels = regime_names[argmax]

    if dates is None:
        date_index = np.asarray(np.arange(n_obs), dtype=object)
    else:
        supplied = np.asarray(dates, dtype=object).reshape(-1)
        if supplied.size != n_obs:
            raise ValueError("dates must have the same length as observations")
        date_index = supplied

    return SemiMarkovOutput(
        regime_names=regime_names,
        transition_matrix=transition,
        duration_distributions=duration_pmf,
        posterior_probabilities=posterior,
        filtered_probabilities=filtered,
        regime_state_argmax=argmax,
        regime_labels=labels,
        duration_signal=duration_signal,
        dates=date_index,
    )


__all__ = [
    "SemiMarkovConfig",
    "SemiMarkovOutput",
    "fit_semi_markov_model",
]

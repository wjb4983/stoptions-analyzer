from __future__ import annotations

from dataclasses import dataclass

import numpy as np


@dataclass(frozen=True)
class HierarchicalMoEOutput:
    probability: np.ndarray
    signal: np.ndarray
    gating_weights: np.ndarray
    fallback_mask: np.ndarray
    expert_average_weights: dict[str, float]
    model_contributions: dict[str, np.ndarray]


class HierarchicalMoE:
    """Hierarchical Mixture-of-Experts for regime/microstate-aware ensembling.

    Experts are organized into pools keyed by a market state (regime or microstate).
    At inference, a gating network blends experts using:
      * posterior probability of each state,
      * uncertainty penalty,
      * expert fit inside each state pool.

    If gating confidence is low, weights back off to learned global fallback weights.
    """

    def __init__(self, *, fallback_confidence_threshold: float = 0.55, uncertainty_key: str = "regime_uncertainty") -> None:
        self.fallback_confidence_threshold = float(np.clip(fallback_confidence_threshold, 0.0, 1.0))
        self.uncertainty_key = uncertainty_key

        self.state_names_: list[str] = []
        self.state_to_index_: dict[str, int] = {}
        self.expert_names_: list[str] = []

        self.fallback_weights_: np.ndarray | None = None
        self.pool_weights_: np.ndarray | None = None

    def fit(
        self,
        *,
        expert_predictions: dict[str, np.ndarray],
        labels: np.ndarray,
        features: dict[str, np.ndarray],
    ) -> None:
        self.expert_names_ = list(expert_predictions.keys())
        if not self.expert_names_:
            raise ValueError("expert_predictions must be non-empty")

        y = np.asarray(labels, dtype=float)
        n_samples = y.shape[0]
        pred_matrix = self._build_prediction_matrix(expert_predictions=expert_predictions, n_samples=n_samples)

        states = self._extract_states(features=features, n_samples=n_samples)
        self.state_names_ = sorted(set(states.tolist()))
        self.state_to_index_ = {state: idx for idx, state in enumerate(self.state_names_)}

        correct_matrix = (pred_matrix >= 0.5) == (y[:, None] >= 0.5)
        perf = np.mean(correct_matrix, axis=0)
        self.fallback_weights_ = self._normalize(perf)

        self.pool_weights_ = np.full((len(self.state_names_), len(self.expert_names_)), 1.0 / len(self.expert_names_), dtype=float)
        for state_idx, state in enumerate(self.state_names_):
            mask = states == state
            if not np.any(mask):
                continue
            state_perf = np.mean(correct_matrix[mask], axis=0)
            self.pool_weights_[state_idx] = self._normalize(state_perf)

    def predict(
        self,
        *,
        expert_predictions: dict[str, np.ndarray],
        features: dict[str, np.ndarray],
    ) -> HierarchicalMoEOutput:
        if self.pool_weights_ is None or self.fallback_weights_ is None:
            raise RuntimeError("HierarchicalMoE must be fit before predict")

        n_samples = len(next(iter(expert_predictions.values())))
        pred_matrix = self._build_prediction_matrix(expert_predictions=expert_predictions, n_samples=n_samples)
        state_posterior = self._extract_state_posterior(features=features, n_samples=n_samples)
        uncertainty = self._extract_uncertainty(features=features, posterior=state_posterior)

        pool_weight_by_sample = np.dot(state_posterior, self.pool_weights_)
        uncertainty_scale = np.clip(1.0 - uncertainty, 0.05, 1.0)[:, None]
        weighted = pool_weight_by_sample * uncertainty_scale
        weighted = np.apply_along_axis(self._normalize, axis=1, arr=weighted)

        gate_confidence = np.max(weighted, axis=1)
        fallback_mask = gate_confidence < self.fallback_confidence_threshold
        if np.any(fallback_mask):
            weighted[fallback_mask] = self.fallback_weights_

        blended_prob = np.sum(pred_matrix * weighted, axis=1)
        signal = np.where(blended_prob >= 0.5, 1.0, -1.0)
        contributions = {
            expert: pred_matrix[:, idx] * weighted[:, idx]
            for idx, expert in enumerate(self.expert_names_)
        }
        avg_weights = {
            expert: float(np.mean(weighted[:, idx])) for idx, expert in enumerate(self.expert_names_)
        }

        return HierarchicalMoEOutput(
            probability=blended_prob,
            signal=signal,
            gating_weights=weighted,
            fallback_mask=fallback_mask,
            expert_average_weights=avg_weights,
            model_contributions=contributions,
        )

    def _build_prediction_matrix(self, *, expert_predictions: dict[str, np.ndarray], n_samples: int) -> np.ndarray:
        cols: list[np.ndarray] = []
        for expert in self.expert_names_:
            preds = np.asarray(expert_predictions[expert], dtype=float)
            if preds.ndim != 1 or preds.shape[0] != n_samples:
                raise ValueError("each expert prediction array must be 1D and aligned")
            cols.append(preds)
        return np.column_stack(cols)

    def _extract_states(self, *, features: dict[str, np.ndarray], n_samples: int) -> np.ndarray:
        state_feature = features.get("market_microstate", features.get("regime"))
        if state_feature is None:
            return np.full(n_samples, "global", dtype=object)
        states = np.asarray(state_feature).astype(str)
        if states.shape[0] != n_samples:
            raise ValueError("regime/market_microstate feature must align with predictions")
        return states

    def _extract_state_posterior(self, *, features: dict[str, np.ndarray], n_samples: int) -> np.ndarray:
        posterior = features.get("market_microstate_posterior", features.get("regime_posterior"))
        if posterior is None:
            states = self._extract_states(features=features, n_samples=n_samples)
            mat = np.zeros((n_samples, len(self.state_names_)), dtype=float)
            for idx, state in enumerate(states):
                state_idx = self.state_to_index_.get(str(state), None)
                if state_idx is None:
                    mat[idx] = 1.0 / len(self.state_names_)
                else:
                    mat[idx, state_idx] = 1.0
            return mat

        posterior_arr = np.asarray(posterior, dtype=float)
        if posterior_arr.ndim != 2 or posterior_arr.shape[0] != n_samples:
            raise ValueError("regime_posterior/market_microstate_posterior must be 2D and aligned")
        if posterior_arr.shape[1] != len(self.state_names_):
            raise ValueError("posterior column count must match number of fitted states")

        row_sums = np.sum(posterior_arr, axis=1, keepdims=True)
        row_sums = np.where(row_sums <= 0.0, 1.0, row_sums)
        return posterior_arr / row_sums

    def _extract_uncertainty(self, *, features: dict[str, np.ndarray], posterior: np.ndarray) -> np.ndarray:
        uncertainty_feature = features.get(self.uncertainty_key)
        if uncertainty_feature is not None:
            uncertainty = np.asarray(uncertainty_feature, dtype=float)
            if uncertainty.ndim != 1 or uncertainty.shape[0] != posterior.shape[0]:
                raise ValueError(f"{self.uncertainty_key} must be 1D and aligned with predictions")
            return np.clip(uncertainty, 0.0, 1.0)

        entropy = -np.sum(posterior * np.log(np.clip(posterior, 1e-12, 1.0)), axis=1)
        max_entropy = np.log(max(2, posterior.shape[1]))
        return np.clip(entropy / max_entropy, 0.0, 1.0)

    @staticmethod
    def _normalize(values: np.ndarray) -> np.ndarray:
        vals = np.asarray(values, dtype=float)
        vals = np.clip(vals, 0.0, None)
        total = float(np.sum(vals))
        if total <= 0.0:
            return np.full(vals.shape, 1.0 / max(1, vals.size), dtype=float)
        return vals / total


__all__ = ["HierarchicalMoE", "HierarchicalMoEOutput"]

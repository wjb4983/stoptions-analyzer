from __future__ import annotations

from dataclasses import dataclass

import numpy as np


@dataclass(frozen=True)
class AcceptancePolicy:
    """Rule set for accepting/rejecting base trade signals."""

    min_base_confidence: float = 0.5
    min_risk_filter_score: float = 0.0
    min_conformal_p_value: float = 0.05


@dataclass(frozen=True)
class PolicyDecision:
    accepted_mask: np.ndarray
    rejected_mask: np.ndarray
    gated_signal: np.ndarray
    p_values: np.ndarray
    empirical_risk: float
    empirical_coverage: float


class MetaLabelConformalModel:
    """Conformal meta-labeling model with risk/coverage-aware trade filtering."""

    def __init__(
        self,
        *,
        acceptance_policy: AcceptancePolicy | None = None,
        target_risk: float = 0.2,
        min_coverage: float = 0.0,
    ) -> None:
        self.acceptance_policy = acceptance_policy or AcceptancePolicy()
        self.target_risk = float(np.clip(target_risk, 0.0, 1.0))
        self.min_coverage = float(np.clip(min_coverage, 0.0, 1.0))

        self.feature_names = ("base_signal", "base_confidence", "risk_filter_score")
        self.feature_importances_: dict[str, float] = {}

        self._weights: np.ndarray | None = None
        self._calibration_nonconformity: np.ndarray = np.array([], dtype=float)
        self._adaptive_min_p_value: float = self.acceptance_policy.min_conformal_p_value

    def fit(self, features: dict[str, np.ndarray], labels: np.ndarray) -> None:
        x = self._build_feature_matrix(features)
        y = np.asarray(labels, dtype=float)
        if y.ndim != 1 or y.shape[0] != x.shape[0]:
            raise ValueError("labels must be 1D and aligned to feature rows")

        binary_y = (y >= 0.5).astype(float)

        centered_x = x - np.mean(x, axis=0)
        centered_y = binary_y - np.mean(binary_y)
        raw_importances = np.abs(np.dot(centered_x.T, centered_y))
        total = float(np.sum(raw_importances))
        if total <= 0.0:
            weights = np.full(raw_importances.shape, 1.0 / max(1, raw_importances.size), dtype=float)
        else:
            weights = raw_importances / total
        self._weights = weights
        self.feature_importances_ = dict(zip(self.feature_names, weights, strict=True))

        probs = self.predict_proba(features)
        self._calibration_nonconformity = self._nonconformity_from_probs(probs, binary_y)
        self._adaptive_min_p_value = self._select_p_value_for_risk_coverage(
            p_values=self.conformal_p_values(probs),
            labels=binary_y,
            probs=probs,
        )

    def predict_proba(self, features: dict[str, np.ndarray]) -> np.ndarray:
        if self._weights is None:
            raise RuntimeError("Model must be fit before predict_proba")
        x = self._build_feature_matrix(features)
        logits = np.dot(x, self._weights)
        logits = logits - np.mean(logits)
        return 1.0 / (1.0 + np.exp(-logits))

    def conformal_p_values(self, probabilities: np.ndarray) -> np.ndarray:
        probs = np.asarray(probabilities, dtype=float)
        if self._calibration_nonconformity.size == 0:
            # Pre-fit fallback: confidence-derived proxy p-values.
            return np.clip(np.maximum(probs, 1.0 - probs), 0.0, 1.0)

        test_scores = 1.0 - np.maximum(probs, 1.0 - probs)
        cal = self._calibration_nonconformity
        p_values = np.array([(1.0 + np.sum(cal >= s)) / (cal.size + 1.0) for s in test_scores], dtype=float)
        return np.clip(p_values, 0.0, 1.0)

    def apply_policy(self, features: dict[str, np.ndarray]) -> PolicyDecision:
        probs = self.predict_proba(features)
        p_values = self.conformal_p_values(probs)

        base_signal = np.asarray(features["base_signal"], dtype=float)
        base_conf = np.asarray(features["base_confidence"], dtype=float)
        risk_score = np.asarray(features["risk_filter_score"], dtype=float)

        min_p = max(self.acceptance_policy.min_conformal_p_value, self._adaptive_min_p_value)
        accepted = (
            (np.abs(base_signal) > 0.0)
            & (base_conf >= self.acceptance_policy.min_base_confidence)
            & (risk_score >= self.acceptance_policy.min_risk_filter_score)
            & (p_values >= min_p)
        )
        rejected = ~accepted
        gated = np.where(accepted, base_signal, 0.0)

        predicted = probs >= 0.5
        risk = float(np.mean(np.where(accepted, ~predicted, False))) if accepted.size else 0.0
        coverage = float(np.mean(accepted)) if accepted.size else 0.0

        return PolicyDecision(
            accepted_mask=accepted,
            rejected_mask=rejected,
            gated_signal=gated,
            p_values=p_values,
            empirical_risk=risk,
            empirical_coverage=coverage,
        )

    def _build_feature_matrix(self, features: dict[str, np.ndarray]) -> np.ndarray:
        cols: list[np.ndarray] = []
        n: int | None = None
        for name in self.feature_names:
            if name not in features:
                raise KeyError(f"Feature '{name}' is required")
            values = np.asarray(features[name], dtype=float)
            if values.ndim != 1:
                raise ValueError(f"Feature '{name}' must be 1D")
            if n is None:
                n = values.shape[0]
            elif values.shape[0] != n:
                raise ValueError("All features must have the same number of samples")
            cols.append(values)
        return np.column_stack(cols)

    @staticmethod
    def _nonconformity_from_probs(probs: np.ndarray, labels: np.ndarray) -> np.ndarray:
        p = np.asarray(probs, dtype=float)
        y = np.asarray(labels, dtype=float)
        true_prob = np.where(y >= 0.5, p, 1.0 - p)
        return 1.0 - np.clip(true_prob, 0.0, 1.0)

    def _select_p_value_for_risk_coverage(self, *, p_values: np.ndarray, labels: np.ndarray, probs: np.ndarray) -> float:
        candidates = np.unique(np.concatenate(([0.0], p_values)))
        predicted = probs >= 0.5
        actual = labels >= 0.5

        best_threshold = float(self.acceptance_policy.min_conformal_p_value)
        best_coverage = -1.0
        for thr in candidates:
            accepted = p_values >= thr
            coverage = float(np.mean(accepted)) if accepted.size else 0.0
            if coverage < self.min_coverage:
                continue

            if np.any(accepted):
                risk = float(np.mean(predicted[accepted] != actual[accepted]))
            else:
                risk = 0.0

            if risk <= self.target_risk and coverage > best_coverage:
                best_threshold = float(thr)
                best_coverage = coverage

        return float(np.clip(best_threshold, 0.0, 1.0))

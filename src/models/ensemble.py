from __future__ import annotations

import logging
from dataclasses import dataclass

import numpy as np

from .base import ModelInterface
from backtesting.validation import generate_purged_kfold_splits

LOGGER = logging.getLogger(__name__)


@dataclass(frozen=True)
class EnsembleOutput:
    signal: np.ndarray
    probability: np.ndarray
    confidence_scores: np.ndarray
    feature_importances: dict[str, dict[str, float]]
    model_weights: dict[str, float]
    model_contributions: dict[str, np.ndarray]


@dataclass(frozen=True)
class ContributionSnapshot:
    index: int
    regime: str
    model_name: str
    weight: float
    contribution: float


class ModelEnsembler:
    def __init__(self, models: list[tuple[ModelInterface, float]], recent_window: int = 20) -> None:
        if not models:
            raise ValueError("models must be non-empty")
        self.models = models
        self.recent_window = int(max(5, recent_window))
        self.stacking_weights_: np.ndarray | None = None
        self.base_performance_: dict[str, float] = {}
        self.regime_fit_: dict[str, dict[str, float]] = {}
        self.contribution_history_: list[ContributionSnapshot] = []

    def fit(self, features: dict[str, np.ndarray], labels: np.ndarray) -> None:
        y = np.asarray(labels, dtype=float)
        regime_labels = _extract_regimes(features, y.shape[0])
        for model, _ in self.models:
            model.fit(features, y)
            LOGGER.info("fit model=%s importances=%s", model.name, model.feature_importances_)
            probs = model.predict_proba(features)
            recent_start = max(0, y.size - self.recent_window)
            recent_acc = np.mean((probs[recent_start:] >= 0.5) == (y[recent_start:] >= 0.5))
            self.base_performance_[model.name] = float(recent_acc)

            per_regime: dict[str, float] = {}
            for regime in sorted(set(regime_labels.tolist())):
                mask = regime_labels == regime
                if np.any(mask):
                    per_regime[regime] = float(np.mean((probs[mask] >= 0.5) == (y[mask] >= 0.5)))
            self.regime_fit_[model.name] = per_regime

    def weighted_vote(self, features: dict[str, np.ndarray], labels: np.ndarray | None = None) -> EnsembleOutput:
        probs = []
        base_weights = np.asarray([weight for _, weight in self.models], dtype=float)
        weights = self._dynamic_weights(features, base_weights, labels)

        feature_importances: dict[str, dict[str, float]] = {}
        confidences = []
        model_contributions: dict[str, np.ndarray] = {}
        for model, _ in self.models:
            model_probs = model.predict_proba(features)
            probs.append(model_probs)
            feature_importances[model.name] = dict(model.feature_importances_)
            confidences.append(np.abs(model_probs - 0.5) * 2.0)
            LOGGER.info("predict model=%s mean_confidence=%.4f", model.name, float(np.mean(confidences[-1])))

        prob_matrix = np.column_stack(probs)
        blended_prob = np.dot(prob_matrix, weights)
        confidence = np.dot(np.column_stack(confidences), weights)
        regimes = _extract_regimes(features, prob_matrix.shape[0])
        for model_idx, (model, _) in enumerate(self.models):
            contrib = prob_matrix[:, model_idx] * weights[model_idx]
            model_contributions[model.name] = contrib
            for idx, value in enumerate(contrib):
                self.contribution_history_.append(
                    ContributionSnapshot(
                        index=int(idx),
                        regime=str(regimes[idx]),
                        model_name=model.name,
                        weight=float(weights[model_idx]),
                        contribution=float(value),
                    )
                )
        signal = np.where(blended_prob >= 0.5, 1.0, -1.0)
        return EnsembleOutput(
            signal=signal,
            probability=blended_prob,
            confidence_scores=confidence,
            feature_importances=feature_importances,
            model_weights={model.name: float(weights[idx]) for idx, (model, _) in enumerate(self.models)},
            model_contributions=model_contributions,
        )

    def fit_stacking(
        self,
        features: dict[str, np.ndarray],
        labels: np.ndarray,
        *,
        n_splits: int = 5,
        purge_window_bars: int = 1,
        embargo_window_bars: int = 1,
    ) -> None:
        self.fit(features, labels)
        y = np.asarray(labels, dtype=float)
        base_prob_matrix = self._oof_base_prob_matrix(
            features,
            y,
            n_splits=n_splits,
            purge_window_bars=purge_window_bars,
            embargo_window_bars=embargo_window_bars,
        )
        centered_y = y - np.mean(y)
        centered_x = base_prob_matrix - np.mean(base_prob_matrix, axis=0)
        coeffs, *_ = np.linalg.lstsq(centered_x, centered_y, rcond=None)
        if np.allclose(coeffs, 0.0):
            coeffs = np.full(coeffs.shape, 1.0 / max(1, coeffs.size), dtype=float)
        self.stacking_weights_ = coeffs
        LOGGER.info("stacking weights=%s", coeffs)

    def stacking_predict(self, features: dict[str, np.ndarray]) -> EnsembleOutput:
        if self.stacking_weights_ is None:
            raise RuntimeError("fit_stacking must be called before stacking_predict")

        probs = np.column_stack([model.predict_proba(features) for model, _ in self.models])
        logits = np.dot(probs, self.stacking_weights_)
        centered = logits - np.mean(logits)
        ensemble_prob = 1.0 / (1.0 + np.exp(-centered))
        signal = np.where(ensemble_prob >= 0.5, 1.0, -1.0)

        feature_importances = {model.name: dict(model.feature_importances_) for model, _ in self.models}
        confidence = np.abs(ensemble_prob - 0.5) * 2.0
        return EnsembleOutput(
            signal=signal,
            probability=ensemble_prob,
            confidence_scores=confidence,
            feature_importances=feature_importances,
            model_weights={model.name: float(self.stacking_weights_[idx]) for idx, (model, _) in enumerate(self.models)},
            model_contributions={model.name: probs[:, idx] * float(self.stacking_weights_[idx]) for idx, (model, _) in enumerate(self.models)},
        )

    def contribution_by_regime(self) -> dict[str, dict[str, float]]:
        summary: dict[str, dict[str, list[float]]] = {}
        for snapshot in self.contribution_history_:
            summary.setdefault(snapshot.regime, {}).setdefault(snapshot.model_name, []).append(snapshot.contribution)
        return {
            regime: {model: float(np.mean(values)) for model, values in model_values.items()}
            for regime, model_values in summary.items()
        }

    def _dynamic_weights(
        self,
        features: dict[str, np.ndarray],
        base_weights: np.ndarray,
        labels: np.ndarray | None,
    ) -> np.ndarray:
        gross = float(np.sum(np.abs(base_weights)))
        if gross <= 0.0:
            raise ValueError("ensemble weights must have non-zero gross")
        adjusted = np.abs(base_weights / gross)
        regimes = _extract_regimes(features, len(next(iter(features.values()))))
        latest_regime = str(regimes[-1]) if regimes.size else "global"

        for idx, (model, _) in enumerate(self.models):
            uncertainty = 1.0 - float(np.mean(np.abs(model.predict_proba(features) - 0.5) * 2.0))
            recent_perf = self.base_performance_.get(model.name, 0.5)
            regime_perf = self.regime_fit_.get(model.name, {}).get(latest_regime, recent_perf)
            label_perf = 1.0
            if labels is not None and labels.size:
                preds = model.predict_proba(features)
                label_perf = float(np.mean((preds >= 0.5) == (labels >= 0.5)))
            score = (0.45 * recent_perf) + (0.35 * regime_perf) + (0.2 * label_perf)
            adjusted[idx] = adjusted[idx] * max(0.0, score) * max(0.05, 1.0 - uncertainty)

        total = float(np.sum(adjusted))
        if total <= 0.0:
            return np.full(adjusted.shape, 1.0 / max(1, adjusted.size), dtype=float)
        return adjusted / total

    def _oof_base_prob_matrix(
        self,
        features: dict[str, np.ndarray],
        labels: np.ndarray,
        *,
        n_splits: int,
        purge_window_bars: int,
        embargo_window_bars: int,
    ) -> np.ndarray:
        n_samples = labels.shape[0]
        splits = generate_purged_kfold_splits(
            n_samples=n_samples,
            n_splits=max(2, min(int(n_splits), n_samples)),
            purge_window_bars=int(purge_window_bars),
            embargo_window_bars=int(embargo_window_bars),
            label_horizon_bars=1,
        )
        oof = np.full((n_samples, len(self.models)), 0.5, dtype=float)
        seen = np.zeros(n_samples, dtype=bool)
        for split in splits:
            train_idx = split.train_indices
            test_idx = split.test_indices
            train_features = {name: np.asarray(values)[train_idx] for name, values in features.items()}
            test_features = {name: np.asarray(values)[test_idx] for name, values in features.items()}
            y_train = labels[train_idx]

            for model_idx, (model, _) in enumerate(self.models):
                cloned_model = type(model)()
                cloned_model.fit(train_features, y_train)
                oof[test_idx, model_idx] = cloned_model.predict_proba(test_features)
            seen[test_idx] = True

        if not np.all(seen):
            LOGGER.warning("OOF coverage incomplete; filling missing rows with full-fit predictions")
            for model_idx, (model, _) in enumerate(self.models):
                oof[~seen, model_idx] = model.predict_proba(features)[~seen]
        return oof


def _extract_regimes(features: dict[str, np.ndarray], n_samples: int) -> np.ndarray:
    regimes = features.get("regime")
    if regimes is None:
        return np.full(n_samples, "global", dtype=object)
    regime_array = np.asarray(regimes)
    if regime_array.shape[0] != n_samples:
        raise ValueError("regime feature must have same number of rows as all model features")
    return regime_array.astype(str)

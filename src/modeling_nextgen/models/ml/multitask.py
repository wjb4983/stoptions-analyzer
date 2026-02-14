from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import numpy as np


ArrayLike = np.ndarray


@dataclass(frozen=True)
class TaskSpec:
    """Task metadata for a multiclass classification head."""

    name: str
    n_classes: int

    def __post_init__(self) -> None:
        if self.n_classes < 2:
            raise ValueError(f"Task {self.name!r} must have at least 2 classes")


@dataclass
class CalibrationArtifact:
    """Reliability statistics for a single task."""

    expected_calibration_error: float
    bin_edges: list[float]
    bin_accuracy: list[float]
    bin_confidence: list[float]
    bin_count: list[int]


class MultiTaskRiskModel:
    """
    Joint multi-task classifier for:
      1) directional returns,
      2) realized-vol buckets,
      3) drawdown-risk class.

    Architecture:
      - Shared backbone: Linear -> ReLU -> Dropout
      - Per-task heads: independent linear classification layers

    The model also learns per-task homoscedastic uncertainty weights
    (log-variance parameters), and emits post-fit calibration artifacts.
    """

    def __init__(
        self,
        *,
        input_dim: int,
        hidden_dim: int = 64,
        dropout_rate: float = 0.10,
        learning_rate: float = 1e-2,
        epochs: int = 300,
        seed: int = 7,
        tasks: tuple[TaskSpec, ...] | None = None,
    ) -> None:
        if input_dim <= 0:
            raise ValueError("input_dim must be > 0")
        if hidden_dim <= 0:
            raise ValueError("hidden_dim must be > 0")
        if not 0.0 <= dropout_rate < 1.0:
            raise ValueError("dropout_rate must be in [0, 1)")
        if learning_rate <= 0.0:
            raise ValueError("learning_rate must be > 0")
        if epochs <= 0:
            raise ValueError("epochs must be > 0")

        self.input_dim = int(input_dim)
        self.hidden_dim = int(hidden_dim)
        self.dropout_rate = float(dropout_rate)
        self.learning_rate = float(learning_rate)
        self.epochs = int(epochs)
        self._rng = np.random.default_rng(seed)

        default_tasks = (
            TaskSpec("directional_return", 3),
            TaskSpec("realized_vol_bucket", 3),
            TaskSpec("drawdown_risk", 3),
        )
        self.tasks = tasks if tasks is not None else default_tasks
        self._task_idx = {task.name: i for i, task in enumerate(self.tasks)}

        self.W_shared = self._rng.normal(0.0, 0.05, size=(self.input_dim, self.hidden_dim))
        self.b_shared = np.zeros(self.hidden_dim, dtype=float)

        self.W_heads: dict[str, ArrayLike] = {}
        self.b_heads: dict[str, ArrayLike] = {}
        for task in self.tasks:
            self.W_heads[task.name] = self._rng.normal(0.0, 0.05, size=(self.hidden_dim, task.n_classes))
            self.b_heads[task.name] = np.zeros(task.n_classes, dtype=float)

        self.log_variance = np.zeros(len(self.tasks), dtype=float)
        self.calibration_artifacts: dict[str, CalibrationArtifact] = {}

    def fit(
        self,
        X: ArrayLike,
        y_by_task: dict[str, ArrayLike],
        *,
        calibration_data: tuple[ArrayLike, dict[str, ArrayLike]] | None = None,
    ) -> "MultiTaskRiskModel":
        X = np.asarray(X, dtype=float)
        if X.ndim != 2 or X.shape[1] != self.input_dim:
            raise ValueError(f"X must have shape (n_samples, {self.input_dim})")

        y_clean = self._validate_labels(y_by_task=y_by_task, n_samples=X.shape[0])
        n = X.shape[0]

        for _ in range(self.epochs):
            hidden_linear = X @ self.W_shared + self.b_shared
            hidden = np.maximum(hidden_linear, 0.0)

            if self.dropout_rate > 0.0:
                keep_prob = 1.0 - self.dropout_rate
                dropout_mask = (self._rng.random(hidden.shape) < keep_prob).astype(float)
                hidden_train = hidden * dropout_mask / keep_prob
            else:
                dropout_mask = np.ones_like(hidden)
                hidden_train = hidden

            grad_W_shared = np.zeros_like(self.W_shared)
            grad_b_shared = np.zeros_like(self.b_shared)

            hidden_grad_accum = np.zeros_like(hidden)
            for task_idx, task in enumerate(self.tasks):
                y = y_clean[task.name]
                logits = hidden_train @ self.W_heads[task.name] + self.b_heads[task.name]
                probs = _softmax(logits)
                y_one_hot = _one_hot(y, task.n_classes)

                ce_grad_logits = (probs - y_one_hot) / n
                ce_loss = -np.mean(np.sum(y_one_hot * np.log(probs + 1e-12), axis=1))

                precision = np.exp(-self.log_variance[task_idx])
                weighted_grad_logits = precision * ce_grad_logits

                grad_W_head = hidden_train.T @ weighted_grad_logits
                grad_b_head = np.sum(weighted_grad_logits, axis=0)
                self.W_heads[task.name] -= self.learning_rate * grad_W_head
                self.b_heads[task.name] -= self.learning_rate * grad_b_head

                hidden_grad_accum += weighted_grad_logits @ self.W_heads[task.name].T

                grad_log_var = -precision * ce_loss + 1.0
                self.log_variance[task_idx] -= self.learning_rate * grad_log_var

            hidden_grad_accum *= (hidden_linear > 0.0)
            hidden_grad_accum *= dropout_mask / (1.0 - self.dropout_rate if self.dropout_rate < 1.0 else 1.0)
            grad_W_shared += X.T @ hidden_grad_accum
            grad_b_shared += np.sum(hidden_grad_accum, axis=0)

            self.W_shared -= self.learning_rate * grad_W_shared
            self.b_shared -= self.learning_rate * grad_b_shared

        calib_X, calib_y = calibration_data if calibration_data is not None else (X, y_clean)
        probs = self.predict_proba(calib_X)
        self.calibration_artifacts = {
            task.name: _calibration_from_predictions(
                y_true=np.asarray(calib_y[task.name], dtype=int),
                probs=np.asarray(probs[task.name], dtype=float),
            )
            for task in self.tasks
        }

        return self

    def predict_proba(self, X: ArrayLike) -> dict[str, ArrayLike]:
        X = np.asarray(X, dtype=float)
        if X.ndim != 2 or X.shape[1] != self.input_dim:
            raise ValueError(f"X must have shape (n_samples, {self.input_dim})")

        hidden = np.maximum(X @ self.W_shared + self.b_shared, 0.0)
        return {
            task.name: _softmax(hidden @ self.W_heads[task.name] + self.b_heads[task.name])
            for task in self.tasks
        }

    def predict(self, X: ArrayLike) -> dict[str, ArrayLike]:
        probs = self.predict_proba(X)
        return {task_name: np.argmax(task_probs, axis=1) for task_name, task_probs in probs.items()}

    def task_uncertainty(self, X: ArrayLike) -> dict[str, dict[str, ArrayLike | float]]:
        probs = self.predict_proba(X)
        output: dict[str, dict[str, ArrayLike | float]] = {}
        for idx, task in enumerate(self.tasks):
            task_probs = probs[task.name]
            confidence = np.max(task_probs, axis=1)
            entropy = -np.sum(task_probs * np.log(task_probs + 1e-12), axis=1)
            normalized_entropy = entropy / np.log(task.n_classes)
            output[task.name] = {
                "confidence": confidence,
                "entropy": entropy,
                "normalized_entropy": normalized_entropy,
                "aleatoric_scale": float(np.exp(0.5 * self.log_variance[idx])),
            }
        return output

    def artifacts(self, X: ArrayLike) -> dict[str, Any]:
        """Bundle inference probabilities, uncertainty and calibration artifacts."""

        return {
            "probabilities": self.predict_proba(X),
            "uncertainty": self.task_uncertainty(X),
            "calibration": self.calibration_artifacts,
            "task_log_variance": {task.name: float(self.log_variance[idx]) for idx, task in enumerate(self.tasks)},
        }

    def _validate_labels(self, *, y_by_task: dict[str, ArrayLike], n_samples: int) -> dict[str, ArrayLike]:
        missing = [task.name for task in self.tasks if task.name not in y_by_task]
        if missing:
            raise ValueError(f"Missing task labels: {missing}")

        y_clean: dict[str, ArrayLike] = {}
        for task in self.tasks:
            y = np.asarray(y_by_task[task.name], dtype=int)
            if y.ndim != 1 or y.shape[0] != n_samples:
                raise ValueError(f"Labels for task {task.name!r} must have shape ({n_samples},)")
            if np.min(y) < 0 or np.max(y) >= task.n_classes:
                raise ValueError(f"Task {task.name!r} labels must be in [0, {task.n_classes - 1}]")
            y_clean[task.name] = y
        return y_clean


def _softmax(logits: ArrayLike) -> ArrayLike:
    shifted = logits - np.max(logits, axis=1, keepdims=True)
    exp = np.exp(shifted)
    return exp / np.sum(exp, axis=1, keepdims=True)


def _one_hot(y: ArrayLike, n_classes: int) -> ArrayLike:
    out = np.zeros((y.shape[0], n_classes), dtype=float)
    out[np.arange(y.shape[0]), y] = 1.0
    return out


def _calibration_from_predictions(*, y_true: ArrayLike, probs: ArrayLike, n_bins: int = 10) -> CalibrationArtifact:
    pred = np.argmax(probs, axis=1)
    confidence = np.max(probs, axis=1)
    accuracy = (pred == y_true).astype(float)

    bin_edges = np.linspace(0.0, 1.0, n_bins + 1)
    bin_ids = np.digitize(confidence, bin_edges[1:-1], right=False)

    bin_acc: list[float] = []
    bin_conf: list[float] = []
    bin_count: list[int] = []
    ece = 0.0

    for b in range(n_bins):
        mask = bin_ids == b
        count = int(np.sum(mask))
        bin_count.append(count)
        if count == 0:
            bin_acc.append(0.0)
            bin_conf.append(0.0)
            continue

        acc_b = float(np.mean(accuracy[mask]))
        conf_b = float(np.mean(confidence[mask]))
        bin_acc.append(acc_b)
        bin_conf.append(conf_b)
        ece += (count / len(confidence)) * abs(acc_b - conf_b)

    return CalibrationArtifact(
        expected_calibration_error=float(ece),
        bin_edges=bin_edges.tolist(),
        bin_accuracy=bin_acc,
        bin_confidence=bin_conf,
        bin_count=bin_count,
    )


__all__ = ["TaskSpec", "CalibrationArtifact", "MultiTaskRiskModel"]

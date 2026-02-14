from __future__ import annotations

from dataclasses import dataclass
from typing import Iterator, Literal

import numpy as np


TaskType = Literal["regression", "classification"]


@dataclass(frozen=True)
class PanelSplit:
    """A single walk-forward split over panel data."""

    fold_id: int
    train_idx: np.ndarray
    test_idx: np.ndarray
    train_start_date: np.datetime64
    train_end_date: np.datetime64
    test_start_date: np.datetime64
    test_end_date: np.datetime64
    purge_dates: int
    embargo_dates: int
    label_horizon_dates: int

    @property
    def leakage_checks(self) -> dict[str, bool]:
        return {
            "strict_time_order": bool(self.train_end_date < self.test_start_date),
            "no_index_overlap": bool(np.intersect1d(self.train_idx, self.test_idx).size == 0),
        }


class PanelWalkForwardSplitter:
    """Date-aware, asset-group-compatible walk-forward splitter with leakage guards."""

    def __init__(
        self,
        *,
        train_dates: int,
        test_dates: int,
        step_dates: int | None = None,
        purge_dates: int = 0,
        embargo_dates: int = 0,
        label_horizon_dates: int = 1,
    ) -> None:
        if train_dates <= 0 or test_dates <= 0:
            raise ValueError("train_dates and test_dates must be positive")
        if label_horizon_dates <= 0:
            raise ValueError("label_horizon_dates must be positive")
        if purge_dates < 0 or embargo_dates < 0:
            raise ValueError("purge_dates and embargo_dates must be non-negative")

        self.train_dates = int(train_dates)
        self.test_dates = int(test_dates)
        self.step_dates = int(step_dates) if step_dates is not None else int(test_dates)
        self.purge_dates = int(purge_dates)
        self.embargo_dates = int(embargo_dates)
        self.label_horizon_dates = int(label_horizon_dates)

        if self.step_dates <= 0:
            raise ValueError("step_dates must be positive")

    def split(self, *, asset_ids: np.ndarray, dates: np.ndarray) -> Iterator[PanelSplit]:
        asset_ids = np.asarray(asset_ids)
        dates = np.asarray(dates)
        if asset_ids.shape[0] != dates.shape[0]:
            raise ValueError("asset_ids and dates must have the same length")
        if dates.ndim != 1:
            raise ValueError("dates must be 1D")

        unique_dates = np.unique(dates)
        unique_dates.sort()

        effective_gap = max(self.purge_dates, self.label_horizon_dates)
        cursor = 0
        fold_id = 0

        while cursor + self.train_dates + effective_gap + self.embargo_dates + self.test_dates <= unique_dates.size:
            train_start_pos = cursor
            train_end_pos = train_start_pos + self.train_dates
            test_start_pos = train_end_pos + effective_gap + self.embargo_dates
            test_end_pos = test_start_pos + self.test_dates

            train_date_values = unique_dates[train_start_pos:train_end_pos]
            test_date_values = unique_dates[test_start_pos:test_end_pos]

            train_mask = np.isin(dates, train_date_values)
            test_mask = np.isin(dates, test_date_values)
            train_idx = np.flatnonzero(train_mask)
            test_idx = np.flatnonzero(test_mask)

            split = PanelSplit(
                fold_id=fold_id,
                train_idx=train_idx,
                test_idx=test_idx,
                train_start_date=np.min(train_date_values),
                train_end_date=np.max(train_date_values),
                test_start_date=np.min(test_date_values),
                test_end_date=np.max(test_date_values),
                purge_dates=self.purge_dates,
                embargo_dates=self.embargo_dates,
                label_horizon_dates=self.label_horizon_dates,
            )
            _validate_split_leakage(asset_ids=asset_ids, dates=dates, split=split)

            yield split
            fold_id += 1
            cursor += self.step_dates


def _validate_split_leakage(*, asset_ids: np.ndarray, dates: np.ndarray, split: PanelSplit) -> None:
    if not all(split.leakage_checks.values()):
        raise ValueError(f"Fold {split.fold_id} violates basic leakage checks: {split.leakage_checks}")

    train_pairs = set(zip(asset_ids[split.train_idx].tolist(), dates[split.train_idx].tolist()))
    test_pairs = set(zip(asset_ids[split.test_idx].tolist(), dates[split.test_idx].tolist()))
    if train_pairs.intersection(test_pairs):
        raise ValueError(f"Fold {split.fold_id} has overlapping asset/date rows between train and test")


@dataclass
class _SklearnPanelBaseline:
    """Unified interface for sklearn panel baselines with strict split checks."""

    estimator_name: str
    task: TaskType
    estimator_params: dict

    def __post_init__(self) -> None:
        self._estimator = None

    def fit(self, X: np.ndarray, y: np.ndarray) -> "_SklearnPanelBaseline":
        X = np.asarray(X, dtype=float)
        y = np.asarray(y)
        if X.ndim != 2:
            raise ValueError("X must be a 2D array")
        if y.ndim != 1 or y.shape[0] != X.shape[0]:
            raise ValueError("y must be 1D and aligned with X")

        self._estimator = _build_sklearn_estimator(
            estimator_name=self.estimator_name,
            task=self.task,
            estimator_params=self.estimator_params,
        )
        self._estimator.fit(X, y)
        return self

    def predict(self, X: np.ndarray) -> np.ndarray:
        if self._estimator is None:
            raise RuntimeError("Model is not fitted")
        X = np.asarray(X, dtype=float)
        if self.task == "classification" and hasattr(self._estimator, "predict_proba"):
            return np.asarray(self._estimator.predict_proba(X)[:, 1], dtype=float)
        return np.asarray(self._estimator.predict(X), dtype=float)


class ElasticNetBaseline(_SklearnPanelBaseline):
    def __init__(self, **estimator_params: float) -> None:
        super().__init__(estimator_name="elastic_net", task="regression", estimator_params=estimator_params)


class LogitBaseline(_SklearnPanelBaseline):
    def __init__(self, **estimator_params: float) -> None:
        super().__init__(estimator_name="logit", task="classification", estimator_params=estimator_params)


class TreeBoostingBaseline(_SklearnPanelBaseline):
    def __init__(self, task: TaskType = "regression", **estimator_params: float) -> None:
        super().__init__(estimator_name="tree_boosting", task=task, estimator_params=estimator_params)


class RandomForestBaseline(_SklearnPanelBaseline):
    def __init__(self, task: TaskType = "regression", **estimator_params: float) -> None:
        super().__init__(estimator_name="random_forest", task=task, estimator_params=estimator_params)


def _build_sklearn_estimator(*, estimator_name: str, task: TaskType, estimator_params: dict):
    try:
        from sklearn.ensemble import (
            GradientBoostingClassifier,
            GradientBoostingRegressor,
            RandomForestClassifier,
            RandomForestRegressor,
        )
        from sklearn.linear_model import ElasticNet, LogisticRegression
    except ImportError as exc:  # pragma: no cover - environment-dependent
        raise ImportError(
            "scikit-learn is required for panel baseline models; install scikit-learn to use these classes"
        ) from exc

    if estimator_name == "elastic_net":
        return ElasticNet(**estimator_params)
    if estimator_name == "logit":
        return LogisticRegression(**estimator_params)
    if estimator_name == "tree_boosting":
        return GradientBoostingRegressor(**estimator_params) if task == "regression" else GradientBoostingClassifier(**estimator_params)
    if estimator_name == "random_forest":
        return RandomForestRegressor(**estimator_params) if task == "regression" else RandomForestClassifier(**estimator_params)

    raise ValueError(f"Unknown estimator_name: {estimator_name}")


__all__ = [
    "ElasticNetBaseline",
    "LogitBaseline",
    "TreeBoostingBaseline",
    "RandomForestBaseline",
    "PanelSplit",
    "PanelWalkForwardSplitter",
]

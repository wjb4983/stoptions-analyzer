from __future__ import annotations

from dataclasses import dataclass

import numpy as np


@dataclass(frozen=True)
class ReliabilityBin:
    """Single-bin summary point for reliability diagrams."""

    lower_bound: float
    upper_bound: float
    avg_prediction: float
    empirical_frequency: float
    count: int


@dataclass(frozen=True)
class ProbabilityCalibrationReport:
    """Calibration diagnostics, including reliability diagram data and scalar metrics."""

    method: str
    sample_size: int
    expected_calibration_error: float
    brier_score: float
    reliability_bins: tuple[ReliabilityBin, ...]
    strategy: str | None = None
    window_start: int | None = None
    window_end: int | None = None


class ProbabilityCalibrator:
    """Binary probability calibration with optional walk-forward re-training.

    Supported methods:
    - ``identity``: no-op passthrough.
    - ``platt``: logistic (Platt-style) calibration on logit(probability).
    - ``isotonic``: monotonic piecewise-constant calibration (PAV algorithm).
    - ``histogram``: equal-width binning calibration.
    """

    name = "probability"

    def __init__(self, method: str = "identity", n_bins: int = 10, eps: float = 1e-8) -> None:
        method_value = str(method).strip().lower()
        if method_value not in {"identity", "platt", "isotonic", "histogram"}:
            raise ValueError("method must be one of {'identity', 'platt', 'isotonic', 'histogram'}")
        if int(n_bins) <= 1:
            raise ValueError("n_bins must be greater than 1")

        self.method = method_value
        self.n_bins = int(n_bins)
        self.eps = float(eps)

        self._fitted = False
        self._platt_coef: tuple[float, float] | None = None
        self._iso_x: np.ndarray | None = None
        self._iso_y: np.ndarray | None = None
        self._hist_bin_edges: np.ndarray | None = None
        self._hist_bin_values: np.ndarray | None = None

    @staticmethod
    def _as_1d(array: np.ndarray) -> np.ndarray:
        arr = np.asarray(array, dtype=float).reshape(-1)
        if arr.size == 0:
            raise ValueError("inputs cannot be empty")
        return arr

    def fit(self, probabilities: np.ndarray, labels: np.ndarray) -> None:
        probs = np.clip(self._as_1d(probabilities), self.eps, 1.0 - self.eps)
        y = self._as_1d(labels)
        if probs.shape[0] != y.shape[0]:
            raise ValueError("probabilities and labels must have identical length")

        if self.method == "identity":
            self._fitted = True
            return

        if self.method == "platt":
            self._fit_platt(probs, y)
        elif self.method == "isotonic":
            self._fit_isotonic(probs, y)
        elif self.method == "histogram":
            self._fit_histogram(probs, y)

        self._fitted = True

    def transform(self, probabilities: np.ndarray) -> np.ndarray:
        probs = np.clip(self._as_1d(probabilities), self.eps, 1.0 - self.eps)
        if self.method == "identity" or not self._fitted:
            return probs

        if self.method == "platt":
            assert self._platt_coef is not None
            a, b = self._platt_coef
            logits = np.log(probs / (1.0 - probs))
            return 1.0 / (1.0 + np.exp(-(a * logits + b)))

        if self.method == "isotonic":
            assert self._iso_x is not None and self._iso_y is not None
            out = np.interp(probs, self._iso_x, self._iso_y, left=self._iso_y[0], right=self._iso_y[-1])
            return np.clip(out, 0.0, 1.0)

        assert self._hist_bin_edges is not None and self._hist_bin_values is not None
        idx = np.digitize(probs, self._hist_bin_edges[1:-1], right=False)
        return np.clip(self._hist_bin_values[idx], 0.0, 1.0)

    def fit_transform(self, probabilities: np.ndarray, labels: np.ndarray) -> np.ndarray:
        self.fit(probabilities, labels)
        return self.transform(probabilities)

    def report(
        self,
        probabilities: np.ndarray,
        labels: np.ndarray,
        calibrated_probabilities: np.ndarray | None = None,
    ) -> ProbabilityCalibrationReport:
        y = self._as_1d(labels)
        probs = self.transform(probabilities) if calibrated_probabilities is None else self._as_1d(calibrated_probabilities)
        if probs.shape[0] != y.shape[0]:
            raise ValueError("probabilities and labels must have identical length")

        reliability_bins = self._reliability_bins(probs, y, self.n_bins)
        ece = float(
            sum(abs(bin_.avg_prediction - bin_.empirical_frequency) * (bin_.count / max(len(y), 1)) for bin_ in reliability_bins)
        )
        brier = float(np.mean((probs - y) ** 2))

        return ProbabilityCalibrationReport(
            method=self.method,
            sample_size=int(y.size),
            expected_calibration_error=ece,
            brier_score=brier,
            reliability_bins=tuple(reliability_bins),
        )

    def walk_forward_retraining(
        self,
        probabilities: np.ndarray,
        labels: np.ndarray,
        *,
        strategy: str = "expanding",
        min_train_size: int = 100,
        window_size: int | None = None,
    ) -> tuple[np.ndarray, list[ProbabilityCalibrationReport]]:
        """Calibrate sequentially with expanding/rolling retraining windows.

        Returns calibrated probabilities and per-step reports computed on each training window.
        """

        probs = self._as_1d(probabilities)
        y = self._as_1d(labels)
        if probs.shape[0] != y.shape[0]:
            raise ValueError("probabilities and labels must have identical length")
        if int(min_train_size) < 2:
            raise ValueError("min_train_size must be >= 2")

        strategy_value = str(strategy).strip().lower()
        if strategy_value not in {"expanding", "rolling"}:
            raise ValueError("strategy must be one of {'expanding', 'rolling'}")
        if strategy_value == "rolling" and (window_size is None or int(window_size) < 2):
            raise ValueError("window_size must be >= 2 for rolling strategy")

        calibrated = np.full_like(probs, fill_value=np.nan, dtype=float)
        reports: list[ProbabilityCalibrationReport] = []

        for end in range(min_train_size, len(probs)):
            if strategy_value == "rolling":
                start = max(0, end - int(window_size))
            else:
                start = 0

            train_probs = probs[start:end]
            train_y = y[start:end]

            step_calibrator = ProbabilityCalibrator(method=self.method, n_bins=self.n_bins, eps=self.eps)
            step_calibrator.fit(train_probs, train_y)
            calibrated[end] = step_calibrator.transform(np.array([probs[end]], dtype=float))[0]

            train_calibrated = step_calibrator.transform(train_probs)
            report = step_calibrator.report(train_probs, train_y, calibrated_probabilities=train_calibrated)
            reports.append(
                ProbabilityCalibrationReport(
                    method=report.method,
                    sample_size=report.sample_size,
                    expected_calibration_error=report.expected_calibration_error,
                    brier_score=report.brier_score,
                    reliability_bins=report.reliability_bins,
                    strategy=strategy_value,
                    window_start=start,
                    window_end=end,
                )
            )

        return calibrated, reports

    def _fit_platt(self, probs: np.ndarray, y: np.ndarray) -> None:
        z = np.log(probs / (1.0 - probs))
        a = 1.0
        b = 0.0

        for _ in range(100):
            logits = a * z + b
            p = 1.0 / (1.0 + np.exp(-logits))

            grad_a = np.mean((p - y) * z)
            grad_b = np.mean(p - y)
            h_aa = np.mean(p * (1.0 - p) * z * z) + 1e-10
            h_ab = np.mean(p * (1.0 - p) * z)
            h_bb = np.mean(p * (1.0 - p)) + 1e-10

            det = h_aa * h_bb - h_ab * h_ab
            if abs(det) < 1e-12:
                break

            delta_a = (h_bb * grad_a - h_ab * grad_b) / det
            delta_b = (-h_ab * grad_a + h_aa * grad_b) / det
            a -= delta_a
            b -= delta_b

            if abs(delta_a) + abs(delta_b) < 1e-8:
                break

        self._platt_coef = (float(a), float(b))

    def _fit_isotonic(self, probs: np.ndarray, y: np.ndarray) -> None:
        order = np.argsort(probs)
        x = probs[order]
        t = y[order]

        values: list[float] = []
        weights: list[int] = []
        ends: list[int] = []

        for i, label in enumerate(t):
            values.append(float(label))
            weights.append(1)
            ends.append(i)

            while len(values) >= 2 and values[-2] > values[-1]:
                new_weight = weights[-2] + weights[-1]
                new_value = (values[-2] * weights[-2] + values[-1] * weights[-1]) / new_weight
                new_end = ends[-1]
                values[-2:] = [new_value]
                weights[-2:] = [new_weight]
                ends[-2:] = [new_end]

        yhat = np.empty_like(t)
        start = 0
        for value, end in zip(values, ends, strict=False):
            yhat[start : end + 1] = value
            start = end + 1

        unique_x, inverse = np.unique(x, return_inverse=True)
        mapped = np.zeros_like(unique_x)
        counts = np.zeros_like(unique_x)
        np.add.at(mapped, inverse, yhat)
        np.add.at(counts, inverse, 1)
        mapped /= np.clip(counts, 1.0, None)

        self._iso_x = unique_x
        self._iso_y = mapped

    def _fit_histogram(self, probs: np.ndarray, y: np.ndarray) -> None:
        edges = np.linspace(0.0, 1.0, self.n_bins + 1)
        idx = np.digitize(probs, edges[1:-1], right=False)
        values = np.zeros(self.n_bins, dtype=float)
        global_mean = float(np.mean(y))

        for bin_idx in range(self.n_bins):
            mask = idx == bin_idx
            values[bin_idx] = float(np.mean(y[mask])) if np.any(mask) else global_mean

        self._hist_bin_edges = edges
        self._hist_bin_values = values

    @staticmethod
    def _reliability_bins(probs: np.ndarray, labels: np.ndarray, n_bins: int) -> list[ReliabilityBin]:
        edges = np.linspace(0.0, 1.0, n_bins + 1)
        idx = np.digitize(probs, edges[1:-1], right=False)
        bins: list[ReliabilityBin] = []

        for i in range(n_bins):
            mask = idx == i
            count = int(np.sum(mask))
            if count == 0:
                bins.append(
                    ReliabilityBin(
                        lower_bound=float(edges[i]),
                        upper_bound=float(edges[i + 1]),
                        avg_prediction=0.0,
                        empirical_frequency=0.0,
                        count=0,
                    )
                )
                continue

            bins.append(
                ReliabilityBin(
                    lower_bound=float(edges[i]),
                    upper_bound=float(edges[i + 1]),
                    avg_prediction=float(np.mean(probs[mask])),
                    empirical_frequency=float(np.mean(labels[mask])),
                    count=count,
                )
            )

        return bins


class IdentityProbabilityCalibrator(ProbabilityCalibrator):
    name = "identity_probability"

    def __init__(self) -> None:
        super().__init__(method="identity")

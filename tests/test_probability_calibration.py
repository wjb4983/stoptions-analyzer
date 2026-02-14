from __future__ import annotations

import numpy as np

from src.modeling_nextgen.calibration.probability import ProbabilityCalibrator


def test_probability_calibrator_supports_multiple_methods() -> None:
    probs = np.array([0.05, 0.1, 0.2, 0.35, 0.6, 0.7, 0.8, 0.9], dtype=float)
    labels = np.array([0, 0, 0, 0, 1, 1, 1, 1], dtype=float)

    for method in ("identity", "platt", "isotonic", "histogram"):
        calibrator = ProbabilityCalibrator(method=method, n_bins=5)
        calibrated = calibrator.fit_transform(probs, labels)

        assert calibrated.shape == probs.shape
        assert np.all(calibrated >= 0.0)
        assert np.all(calibrated <= 1.0)


def test_report_contains_ece_brier_and_reliability_data() -> None:
    probs = np.array([0.05, 0.15, 0.2, 0.35, 0.65, 0.72, 0.81, 0.95], dtype=float)
    labels = np.array([0, 0, 1, 0, 1, 1, 1, 1], dtype=float)

    calibrator = ProbabilityCalibrator(method="histogram", n_bins=4)
    calibrated = calibrator.fit_transform(probs, labels)
    report = calibrator.report(probs, labels, calibrated_probabilities=calibrated)

    assert report.sample_size == len(labels)
    assert report.expected_calibration_error >= 0.0
    assert report.brier_score >= 0.0
    assert len(report.reliability_bins) == 4
    assert sum(bin_.count for bin_ in report.reliability_bins) == len(labels)


def test_walk_forward_retraining_supports_expanding_and_rolling() -> None:
    probs = np.linspace(0.05, 0.95, 20)
    labels = (probs > 0.5).astype(float)

    expanding = ProbabilityCalibrator(method="platt", n_bins=5)
    calibrated_expanding, reports_expanding = expanding.walk_forward_retraining(
        probs,
        labels,
        strategy="expanding",
        min_train_size=5,
    )

    rolling = ProbabilityCalibrator(method="platt", n_bins=5)
    calibrated_rolling, reports_rolling = rolling.walk_forward_retraining(
        probs,
        labels,
        strategy="rolling",
        min_train_size=5,
        window_size=6,
    )

    assert np.isnan(calibrated_expanding[:5]).all()
    assert np.isnan(calibrated_rolling[:5]).all()
    assert np.all((calibrated_expanding[5:] >= 0.0) & (calibrated_expanding[5:] <= 1.0))
    assert np.all((calibrated_rolling[5:] >= 0.0) & (calibrated_rolling[5:] <= 1.0))

    assert len(reports_expanding) == len(probs) - 5
    assert len(reports_rolling) == len(probs) - 5
    assert all(report.strategy == "expanding" for report in reports_expanding)
    assert all(report.strategy == "rolling" for report in reports_rolling)
    assert all(report.window_start is not None for report in reports_rolling)

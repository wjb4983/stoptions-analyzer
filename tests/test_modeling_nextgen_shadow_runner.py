from __future__ import annotations

import numpy as np

from src.modeling_nextgen.core.contracts import PredictionResult
from src.modeling_nextgen.serving.shadow_runner import ShadowRunner


class _StubModel:
    def __init__(self, predictions: np.ndarray, uncertainty: np.ndarray | None) -> None:
        self.name = "stub"
        self._predictions = predictions
        self._uncertainty = uncertainty

    def fit(self, features: dict[str, np.ndarray], labels: np.ndarray) -> None:
        _ = features
        _ = labels

    def predict(self, features: dict[str, np.ndarray]) -> PredictionResult:
        _ = features
        return PredictionResult(
            predictions=self._predictions,
            uncertainty=self._uncertainty,
            metadata={"source": "prod"},
        )


def test_shadow_runner_returns_production_signal_and_logs_governance() -> None:
    prod = _StubModel(np.array([1.0, -2.0, 3.0, -4.0]), np.array([0.2, 0.1, 0.3, 0.5]))
    shadow = _StubModel(np.array([1.5, -1.0, -2.0, -3.0]), np.array([0.3, 0.1, 0.2, 0.9]))
    runner = ShadowRunner(prod, shadow, rolling_window=2)

    result = runner.predict(features={"x": np.ones((4, 1))})

    np.testing.assert_allclose(result.predictions, prod._predictions)
    np.testing.assert_allclose(result.uncertainty, prod._uncertainty)

    governance = result.metadata["shadow_governance"]
    assert governance["rolling_window"] == 2
    assert len(governance["windows"]) == 2
    assert governance["prediction_mean_abs_delta"] > 0.0
    assert governance["uncertainty_mean_abs_delta"] > 0.0
    assert 0.0 <= governance["direction_agreement"] <= 1.0
    assert len(runner.diagnostic_log) == 1


def test_shadow_runner_handles_missing_uncertainty_as_zero_baseline() -> None:
    prod = _StubModel(np.array([0.5, -0.5, 1.0]), None)
    shadow = _StubModel(np.array([0.5, -1.0, 1.5]), None)
    runner = ShadowRunner(prod, shadow, rolling_window=2)

    diagnostic = runner.predict_with_shadow(features={"x": np.ones((3, 1))}).governance_diagnostics

    assert diagnostic["uncertainty_mean_abs_delta"] == 0.0
    assert len(diagnostic["windows"]) == 2


def test_shadow_runner_rejects_mismatched_prediction_shapes() -> None:
    prod = _StubModel(np.array([1.0, 2.0]), np.array([0.1, 0.2]))
    shadow = _StubModel(np.array([1.0]), np.array([0.1]))
    runner = ShadowRunner(prod, shadow)

    try:
        runner.predict(features={"x": np.ones((2, 1))})
    except ValueError as exc:
        assert "matching shapes" in str(exc)
    else:
        raise AssertionError("Expected ValueError for mismatched shapes")

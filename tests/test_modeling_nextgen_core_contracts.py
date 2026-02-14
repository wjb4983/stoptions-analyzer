from __future__ import annotations

from pathlib import Path
from typing import Any

import numpy as np
import pytest

from src.modeling_nextgen.adapters.backtesting_adapter import BacktestingBridge
from src.modeling_nextgen.core.interfaces import NextGenModelInterface
from src.modeling_nextgen.core.schemas import (
    OptionSurfaceTensorPayload,
    PanelFeaturesPayload,
    RegimeLabelsPayload,
    UncertaintyOutputPayload,
)
from src.modeling_nextgen.validation.schemes import (
    PurgedCrossValidator,
    StressValidator,
    WalkForwardValidator,
)


class FakeNextGenModel:
    name = "fake-nextgen"

    def fit(
        self,
        panel_features: PanelFeaturesPayload,
        regime_labels: RegimeLabelsPayload,
        option_surfaces: OptionSurfaceTensorPayload | None = None,
    ) -> None:
        _ = option_surfaces
        assert panel_features.dates.shape[0] == regime_labels.dates.shape[0]

    def predict(
        self,
        panel_features: PanelFeaturesPayload,
        option_surfaces: OptionSurfaceTensorPayload | None = None,
    ) -> RegimeLabelsPayload:
        _ = option_surfaces
        labels = np.zeros(panel_features.dates.shape[0], dtype=int)
        probs = np.column_stack([1.0 - labels, labels]).astype(float)
        return RegimeLabelsPayload(dates=panel_features.dates, labels=labels, probabilities=probs)

    def predict_proba(
        self,
        panel_features: PanelFeaturesPayload,
        option_surfaces: OptionSurfaceTensorPayload | None = None,
    ) -> UncertaintyOutputPayload:
        _ = option_surfaces
        std = np.full(panel_features.dates.shape[0], 0.2, dtype=float)
        return UncertaintyOutputPayload(dates=panel_features.dates, predictive_std=std)

    def predict_distribution(
        self,
        panel_features: PanelFeaturesPayload,
        option_surfaces: OptionSurfaceTensorPayload | None = None,
    ) -> UncertaintyOutputPayload:
        _ = option_surfaces
        std = np.full(panel_features.dates.shape[0], 0.1, dtype=float)
        return UncertaintyOutputPayload(
            dates=panel_features.dates,
            predictive_std=std,
            confidence_lower=np.full(panel_features.dates.shape[0], 0.4, dtype=float),
            confidence_upper=np.full(panel_features.dates.shape[0], 0.6, dtype=float),
        )

    def explain(
        self,
        panel_features: PanelFeaturesPayload,
        option_surfaces: OptionSurfaceTensorPayload | None = None,
    ) -> dict[str, Any]:
        _ = option_surfaces
        return {"n_dates": int(panel_features.dates.shape[0])}

    def calibrate(
        self,
        panel_features: PanelFeaturesPayload,
        regime_labels: RegimeLabelsPayload,
    ) -> None:
        assert panel_features.dates.shape[0] == regime_labels.labels.shape[0]

    def save(self, path: str | Path) -> None:
        Path(path).write_text(self.name)

    def load(self, path: str | Path) -> None:
        _ = Path(path).read_text()


def _requires_nextgen_interface(model: NextGenModelInterface) -> str:
    return model.name


def _build_dates(n_samples: int = 4) -> np.ndarray:
    return np.array([f"2024-01-{day:02d}" for day in range(1, n_samples + 1)], dtype="datetime64[D]")


def test_dataclass_payload_integrity_and_alignment_conventions() -> None:
    dates = _build_dates(4)
    panel = PanelFeaturesPayload(
        assets=np.array(["SPY", "QQQ"], dtype=str),
        dates=dates,
        features={
            "momentum": np.array([[0.1, 0.2, 0.3, 0.4], [0.0, -0.1, -0.2, -0.3]], dtype=float),
            "volatility": np.array([[0.2, 0.3, 0.4, 0.5], [0.3, 0.4, 0.5, 0.6]], dtype=float),
        },
    )
    surface = OptionSurfaceTensorPayload(
        moneyness=np.array([0.9, 1.0], dtype=float),
        tenors=np.array([0.25, 0.5, 1.0], dtype=float),
        dates=dates,
        values=np.ones((2, 3, 4), dtype=float),
    )
    labels = RegimeLabelsPayload(
        dates=dates,
        labels=np.array([0, 1, 0, 1], dtype=int),
        probabilities=np.array(
            [
                [0.9, 0.1],
                [0.2, 0.8],
                [0.7, 0.3],
                [0.1, 0.9],
            ],
            dtype=float,
        ),
    )
    uncertainty = UncertaintyOutputPayload(
        dates=dates,
        predictive_std=np.array([0.05, 0.08, 0.1, 0.07], dtype=float),
        confidence_lower=np.array([0.45, 0.40, 0.35, 0.42], dtype=float),
        confidence_upper=np.array([0.55, 0.60, 0.65, 0.58], dtype=float),
        regime_uncertainty=np.array([0.1, 0.2, 0.15, 0.05], dtype=float),
    )

    for feature_name, feature_values in panel.features.items():
        assert feature_values.shape == (panel.assets.shape[0], panel.dates.shape[0]), feature_name
    assert surface.values.shape == (
        surface.moneyness.shape[0],
        surface.tenors.shape[0],
        surface.dates.shape[0],
    )
    assert labels.labels.shape == labels.dates.shape
    assert labels.probabilities is not None
    assert labels.probabilities.shape[0] == labels.dates.shape[0]
    assert uncertainty.predictive_std.shape == uncertainty.dates.shape
    assert uncertainty.confidence_lower is not None
    assert uncertainty.confidence_upper is not None
    assert uncertainty.confidence_lower.shape == uncertainty.dates.shape
    assert uncertainty.confidence_upper.shape == uncertainty.dates.shape


def test_interval_aware_position_size_confidence_interval_path_and_clipping() -> None:
    bridge = BacktestingBridge()
    probabilities = np.array([0.8, 0.8], dtype=float)

    sized = bridge.interval_aware_position_size(
        probabilities,
        confidence_lower=np.array([0.79, 0.2], dtype=float),
        confidence_upper=np.array([0.81, 0.8], dtype=float),
        interval_aversion=2.0,
        min_scale=0.25,
    )

    assert sized[0] > sized[1]
    assert np.all(np.abs(sized) <= np.abs(probabilities - 0.5) * 2.0)


def test_interval_aware_position_size_predictive_std_and_negative_aversion_handling() -> None:
    bridge = BacktestingBridge()
    probabilities = np.array([0.9, 0.9, 0.9], dtype=float)

    default_sized = bridge.interval_aware_position_size(probabilities)
    std_sized = bridge.interval_aware_position_size(
        probabilities,
        predictive_std=np.array([0.05, 0.5, 1.5], dtype=float),
        interval_aversion=-3.0,
        min_scale=0.2,
    )

    np.testing.assert_allclose(std_sized, default_sized)


def test_interval_aware_position_size_default_and_shape_mismatch_failures() -> None:
    bridge = BacktestingBridge()
    probabilities = np.array([0.2, 0.8, 0.6], dtype=float)

    default_sized = bridge.interval_aware_position_size(probabilities)
    assert default_sized.shape == probabilities.shape

    with pytest.raises(ValueError, match="confidence bounds must match probabilities shape"):
        bridge.interval_aware_position_size(
            probabilities,
            confidence_lower=np.array([0.1, 0.2], dtype=float),
            confidence_upper=np.array([0.9, 0.8], dtype=float),
        )

    with pytest.raises(ValueError, match="predictive_std must match probabilities shape"):
        bridge.interval_aware_position_size(probabilities, predictive_std=np.array([0.1, 0.2], dtype=float))


def test_validator_scheme_splitters_return_expected_structure() -> None:
    n_samples = 7
    validators = [PurgedCrossValidator(), WalkForwardValidator(), StressValidator()]

    for validator in validators:
        splits = validator.split(n_samples)
        assert isinstance(splits, list)
        assert len(splits) == 1
        train_idx, test_idx = splits[0]
        assert isinstance(train_idx, np.ndarray)
        assert isinstance(test_idx, np.ndarray)
        assert train_idx.size == n_samples
        assert test_idx.size == n_samples
        np.testing.assert_array_equal(train_idx, np.arange(n_samples))
        np.testing.assert_array_equal(test_idx, np.arange(n_samples))


def test_protocol_conformance_smoke_for_fake_model(tmp_path: Path) -> None:
    dates = _build_dates(3)
    panel = PanelFeaturesPayload(
        assets=np.array(["SPY"], dtype=str),
        dates=dates,
        features={"signal": np.array([[0.1, 0.2, 0.3]], dtype=float)},
    )
    labels = RegimeLabelsPayload(dates=dates, labels=np.array([0, 1, 0], dtype=int))

    model = FakeNextGenModel()
    assert _requires_nextgen_interface(model) == "fake-nextgen"

    for method_name in (
        "fit",
        "predict",
        "predict_proba",
        "predict_distribution",
        "explain",
        "calibrate",
        "save",
        "load",
    ):
        assert callable(getattr(model, method_name))

    model.fit(panel, labels)
    predicted = model.predict(panel)
    proba = model.predict_proba(panel)
    distribution = model.predict_distribution(panel)
    explanation = model.explain(panel)
    model.calibrate(panel, labels)

    checkpoint = tmp_path / "fake-nextgen.chk"
    model.save(checkpoint)
    model.load(checkpoint)

    assert predicted.labels.shape == labels.labels.shape
    assert proba.predictive_std.shape == labels.labels.shape
    assert distribution.confidence_lower is not None
    assert distribution.confidence_upper is not None
    assert explanation["n_dates"] == 3

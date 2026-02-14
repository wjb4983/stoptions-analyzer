from __future__ import annotations

import inspect
from pathlib import Path

import numpy as np

from src.modeling_nextgen.core.contracts import Model, PredictionResult
from src.modeling_nextgen.core.interfaces import NextGenModelInterface
from src.modeling_nextgen.core.schemas import (
    OptionSurfaceTensorPayload,
    PanelFeaturesPayload,
    RegimeLabelsPayload,
    UncertaintyOutputPayload,
)
from src.modeling_nextgen.models import BayesModel, DeepModel, MLModel, MarkovModel, StateSpaceModel


class _InterfaceCompliantModel:
    name = "interface_compliant"

    def fit(
        self,
        panel_features: PanelFeaturesPayload,
        regime_labels: RegimeLabelsPayload,
        option_surfaces: OptionSurfaceTensorPayload | None = None,
    ) -> None:
        _ = (panel_features, regime_labels, option_surfaces)

    def predict(
        self,
        panel_features: PanelFeaturesPayload,
        option_surfaces: OptionSurfaceTensorPayload | None = None,
    ) -> RegimeLabelsPayload:
        _ = option_surfaces
        n_dates = panel_features.dates.shape[0]
        return RegimeLabelsPayload(dates=panel_features.dates, labels=np.zeros(n_dates, dtype=int))

    def predict_proba(
        self,
        panel_features: PanelFeaturesPayload,
        option_surfaces: OptionSurfaceTensorPayload | None = None,
    ) -> UncertaintyOutputPayload:
        _ = option_surfaces
        n_dates = panel_features.dates.shape[0]
        std = np.full(n_dates, 0.1, dtype=float)
        return UncertaintyOutputPayload(dates=panel_features.dates, predictive_std=std)

    def predict_distribution(
        self,
        panel_features: PanelFeaturesPayload,
        option_surfaces: OptionSurfaceTensorPayload | None = None,
    ) -> UncertaintyOutputPayload:
        return self.predict_proba(panel_features, option_surfaces)

    def explain(
        self,
        panel_features: PanelFeaturesPayload,
        option_surfaces: OptionSurfaceTensorPayload | None = None,
    ) -> dict[str, float]:
        _ = (panel_features, option_surfaces)
        return {"importance": 1.0}

    def calibrate(self, panel_features: PanelFeaturesPayload, regime_labels: RegimeLabelsPayload) -> None:
        _ = (panel_features, regime_labels)

    def save(self, path: str | Path) -> None:
        _ = path

    def load(self, path: str | Path) -> None:
        _ = path


def _sample_features_labels() -> tuple[dict[str, np.ndarray], np.ndarray]:
    features = {"x": np.array([[1.0], [2.0], [3.0]], dtype=float)}
    labels = np.array([0.0, 1.0, 0.0], dtype=float)
    return features, labels


def test_registered_model_classes_honor_core_model_contract() -> None:
    features, labels = _sample_features_labels()

    for model_cls in (MLModel, DeepModel, MarkovModel, StateSpaceModel, BayesModel):
        model = model_cls()
        assert isinstance(model.name, str) and model.name

        fit_sig = inspect.signature(model.fit)
        predict_sig = inspect.signature(model.predict)
        assert list(fit_sig.parameters) == ["features", "labels"]
        assert list(predict_sig.parameters) == ["features"]

        model.fit(features, labels)
        prediction = model.predict(features)
        assert isinstance(prediction, PredictionResult)

        contract_view: Model = model
        contract_prediction = contract_view.predict(features)
        assert isinstance(contract_prediction, PredictionResult)


def test_schema_payloads_and_nextgen_interface_contract() -> None:
    dates = np.array([np.datetime64("2024-01-01"), np.datetime64("2024-01-02")])
    panel = PanelFeaturesPayload(
        assets=np.array(["AAPL", "MSFT"]),
        dates=dates,
        features={"ret": np.array([[0.1, 0.2], [0.0, -0.1]], dtype=float)},
    )
    surfaces = OptionSurfaceTensorPayload(
        moneyness=np.array([0.9, 1.0, 1.1]),
        tenors=np.array([0.1, 0.5]),
        dates=dates,
        values=np.zeros((3, 2, 2), dtype=float),
    )
    regimes = RegimeLabelsPayload(dates=dates, labels=np.array([0, 1], dtype=int))

    model = _InterfaceCompliantModel()
    iface: NextGenModelInterface = model

    iface.fit(panel, regimes, surfaces)
    pred = iface.predict(panel, surfaces)
    proba = iface.predict_proba(panel, surfaces)
    dist = iface.predict_distribution(panel, surfaces)
    explanation = iface.explain(panel, surfaces)

    assert pred.labels.shape == (2,)
    assert proba.predictive_std.shape == (2,)
    assert dist.predictive_std.shape == (2,)
    assert "importance" in explanation

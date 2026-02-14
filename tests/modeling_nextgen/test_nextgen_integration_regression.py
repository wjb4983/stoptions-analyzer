from __future__ import annotations

from pathlib import Path
from typing import Any

import numpy as np

from src.modeling_nextgen.adapters.backtesting_adapter import BacktestingBridge
from src.modeling_nextgen.calibration.probability import ProbabilityCalibrator
from src.modeling_nextgen.core.interfaces import NextGenModelInterface
from src.modeling_nextgen.core.schemas import (
    OptionSurfaceTensorPayload,
    PanelFeaturesPayload,
    RegimeLabelsPayload,
    UncertaintyOutputPayload,
)
from src.modeling_nextgen.features.no_arb import (
    detect_and_repair_no_arb,
    export_no_arb_diagnostics,
)
from src.modeling_nextgen.models.markov.regime_switching import fit_regime_switching_model
from src.modeling_nextgen.models.markov.semi_markov import fit_semi_markov_model
from src.models.ensemble import ModelEnsembler
from src.models.paradigms import MomentumModel, OptionsDirectionalModel


class DummyNextGenModel:
    name = "dummy_nextgen"

    def __init__(self) -> None:
        self._threshold = 0.0

    def fit(
        self,
        panel_features: PanelFeaturesPayload,
        regime_labels: RegimeLabelsPayload,
        option_surfaces: OptionSurfaceTensorPayload | None = None,
    ) -> None:
        _ = option_surfaces
        signal = panel_features.features["signal"]
        targets = regime_labels.labels.astype(float)
        self._threshold = float(np.mean(signal) - np.mean(targets))

    def predict(
        self,
        panel_features: PanelFeaturesPayload,
        option_surfaces: OptionSurfaceTensorPayload | None = None,
    ) -> RegimeLabelsPayload:
        _ = option_surfaces
        signal = panel_features.features["signal"]
        labels = (signal >= self._threshold).astype(int)
        probs = np.column_stack([1.0 - labels, labels]).astype(float)
        return RegimeLabelsPayload(dates=panel_features.dates, labels=labels, probabilities=probs)

    def predict_proba(
        self,
        panel_features: PanelFeaturesPayload,
        option_surfaces: OptionSurfaceTensorPayload | None = None,
    ) -> UncertaintyOutputPayload:
        _ = option_surfaces
        signal = panel_features.features["signal"]
        centered = signal - np.mean(signal)
        probs = 1.0 / (1.0 + np.exp(-centered))
        std = np.sqrt(np.clip(probs * (1.0 - probs), 0.0, 1.0))
        return UncertaintyOutputPayload(dates=panel_features.dates, predictive_std=std)

    def predict_distribution(
        self,
        panel_features: PanelFeaturesPayload,
        option_surfaces: OptionSurfaceTensorPayload | None = None,
    ) -> UncertaintyOutputPayload:
        output = self.predict_proba(panel_features, option_surfaces)
        std = output.predictive_std
        assert std is not None
        return UncertaintyOutputPayload(
            dates=output.dates,
            predictive_std=std,
            confidence_lower=np.clip(0.5 - std, 0.0, 1.0),
            confidence_upper=np.clip(0.5 + std, 0.0, 1.0),
        )

    def explain(
        self,
        panel_features: PanelFeaturesPayload,
        option_surfaces: OptionSurfaceTensorPayload | None = None,
    ) -> dict[str, Any]:
        _ = option_surfaces
        return {
            "model": self.name,
            "feature_stats": {
                key: float(np.mean(value)) for key, value in panel_features.features.items()
            },
        }

    def calibrate(
        self,
        panel_features: PanelFeaturesPayload,
        regime_labels: RegimeLabelsPayload,
    ) -> None:
        self.fit(panel_features, regime_labels)

    def save(self, path: str | Path) -> None:
        Path(path).write_text(f"{self._threshold:.8f}\n")

    def load(self, path: str | Path) -> None:
        self._threshold = float(Path(path).read_text().strip())


def _accepts_nextgen_interface(model: NextGenModelInterface) -> str:
    return model.name


def test_interface_conformance_round_trip(tmp_path: Path) -> None:
    model = DummyNextGenModel()
    assert _accepts_nextgen_interface(model) == "dummy_nextgen"

    dates = np.array(["2024-01-01", "2024-01-02", "2024-01-03"], dtype="datetime64[D]")
    panel = PanelFeaturesPayload(
        assets=np.array(["SPY", "QQQ"], dtype=str),
        dates=dates,
        features={"signal": np.array([0.1, -0.2, 0.3], dtype=float)},
    )
    labels = RegimeLabelsPayload(dates=dates, labels=np.array([1, 0, 1], dtype=int))

    model.fit(panel, labels)
    pred = model.predict(panel)
    uncertainty = model.predict_distribution(panel)
    model_path = tmp_path / "dummy_nextgen.txt"
    model.save(model_path)

    reloaded = DummyNextGenModel()
    reloaded.load(model_path)

    assert pred.labels.shape == labels.labels.shape
    assert uncertainty.predictive_std.shape == labels.labels.shape
    assert uncertainty.confidence_lower is not None
    assert uncertainty.confidence_upper is not None
    assert reloaded._threshold == model._threshold


def test_no_arb_repair_and_gate_export(tmp_path: Path) -> None:
    moneyness = np.array([0.8, 1.0, 1.2], dtype=np.float64)
    total_variance = np.array(
        [
            [0.10, 0.09, 0.12],
            [0.25, 0.11, 0.15],
            [0.08, 0.10, 0.11],
        ],
        dtype=np.float64,
    )

    repaired = detect_and_repair_no_arb(total_variance, moneyness)
    payload = export_no_arb_diagnostics(repaired.diagnostics, out_path=tmp_path / "no_arb_report.json")

    assert np.all(np.diff(repaired.repaired_total_variance, axis=1) >= -1e-10)
    assert repaired.diagnostics.calendar_violations == 0
    assert repaired.diagnostics.butterfly_violations == 0
    assert payload["model_gate"]["pass"] is True


def test_regime_model_outputs_are_well_formed() -> None:
    observations = np.array(
        [
            [0.1, 0.2],
            [0.2, 0.1],
            [1.4, 1.0],
            [1.2, 1.1],
            [-0.8, -1.0],
            [-0.9, -0.7],
        ],
        dtype=float,
    )
    dates = np.array([f"2024-01-0{i+1}" for i in range(observations.shape[0])], dtype=object)

    rs = fit_regime_switching_model(observations, dates=dates)
    sm = fit_semi_markov_model(observations, dates=dates, duration_features=observations[:, 0])

    assert rs.posterior_probabilities.shape[0] == observations.shape[0]
    assert sm.posterior_probabilities.shape[0] == observations.shape[0]
    assert np.allclose(np.sum(rs.posterior_probabilities, axis=1), 1.0, atol=1e-6)
    assert np.allclose(np.sum(sm.posterior_probabilities, axis=1), 1.0, atol=1e-6)
    assert np.allclose(np.sum(sm.duration_distributions, axis=1), 1.0, atol=1e-6)


def test_calibration_reliability_stability() -> None:
    probs = np.array([0.05, 0.15, 0.25, 0.35, 0.65, 0.75, 0.85, 0.95], dtype=float)
    labels = np.array([0, 0, 0, 0, 1, 1, 1, 1], dtype=float)

    calibrator = ProbabilityCalibrator(method="isotonic", n_bins=4)
    calibrated = calibrator.fit_transform(probs, labels)
    report = calibrator.report(probs, labels, calibrated_probabilities=calibrated)

    assert np.all(np.diff(calibrated) >= -1e-10)
    assert sum(bin_.count for bin_ in report.reliability_bins) == labels.size
    assert report.expected_calibration_error <= 0.2


def test_adapter_compatibility_with_ensemble_and_backtesting() -> None:
    n = 12
    idx = np.linspace(-1.0, 1.0, n)
    features = {
        "returns_1m": idx,
        "returns_3m": idx * 0.5,
        "returns_6m": np.sin(idx),
        "regime": np.where(idx >= 0, "risk_on", "risk_off"),
    }
    labels = (idx > 0).astype(float)

    model = MomentumModel()
    ensembler = ModelEnsembler(models=[(model, 1.0)])
    ensembler.fit(features, labels)
    vote = ensembler.weighted_vote(features, labels=labels)

    bridge = BacktestingBridge()
    sizes = bridge.interval_aware_position_size(
        vote.probability,
        confidence_lower=np.clip(vote.probability - 0.05, 0.0, 1.0),
        confidence_upper=np.clip(vote.probability + 0.05, 0.0, 1.0),
        max_leverage=2.0,
    )
    folds = bridge.build_walk_forward_folds(total_bars=n, train_bars=5, validation_bars=2, test_bars=2, step_bars=1)

    assert sizes.shape == vote.probability.shape
    assert np.max(np.abs(sizes)) <= 2.0
    assert folds


def test_legacy_paradigms_regression_contract() -> None:
    momentum = MomentumModel()
    options_directional = OptionsDirectionalModel()

    assert momentum.required_feature_names() == ("returns_1m", "returns_3m", "returns_6m")
    assert options_directional.required_feature_names() == (
        "skew_z",
        "put_call_flow_imbalance_z",
        "dealer_positioning_proxy_z",
        "gamma_exposure_proxy_rank",
        "unusual_volume_signature_z",
    )

    features = {
        "returns_1m": np.array([-1.0, -0.5, 0.0, 0.5, 1.0]),
        "returns_3m": np.array([-0.7, -0.2, 0.1, 0.4, 0.9]),
        "returns_6m": np.array([-0.4, -0.1, 0.2, 0.3, 0.8]),
    }
    labels = np.array([0, 0, 0, 1, 1], dtype=float)

    momentum.fit(features, labels)
    probs = momentum.predict_proba(features)

    expected = np.array([0.30293239, 0.40624107, 0.50230768, 0.58706233, 0.70095179])
    assert np.allclose(probs, expected, atol=1e-6)

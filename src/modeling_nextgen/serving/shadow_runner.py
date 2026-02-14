from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import numpy as np

from ..core.contracts import Model, PredictionResult


@dataclass(frozen=True)
class DivergenceDiagnostic:
    window_index: int
    start: int
    stop: int
    prediction_mean_abs_delta: float
    uncertainty_mean_abs_delta: float
    direction_agreement: float


@dataclass(frozen=True)
class ShadowInferenceResult:
    production: PredictionResult
    shadow: PredictionResult
    governance_diagnostics: dict[str, Any]


class ShadowRunner:
    """Execute production and shadow inference while keeping production signal unchanged.

    The production prediction is the only signal returned to downstream consumers.
    Shadow outputs and rolling-window divergence diagnostics are stored for governance.
    """

    def __init__(self, production_model: Model, shadow_model: Model, *, rolling_window: int = 50) -> None:
        if rolling_window <= 0:
            raise ValueError("rolling_window must be positive")
        self._production_model = production_model
        self._shadow_model = shadow_model
        self._rolling_window = rolling_window
        self._diagnostic_log: list[dict[str, Any]] = []

    def predict(self, features: dict[str, np.ndarray]) -> PredictionResult:
        """Run shadow inference side-by-side and return production prediction only."""
        production = self._production_model.predict(features)
        shadow = self._shadow_model.predict(features)

        governance = self._build_governance_diagnostics(production=production, shadow=shadow)
        self._diagnostic_log.append(governance)

        production_metadata = dict(production.metadata or {})
        production_metadata["shadow_governance"] = governance
        return PredictionResult(
            predictions=np.asarray(production.predictions, dtype=float),
            probabilities=None if production.probabilities is None else np.asarray(production.probabilities, dtype=float),
            uncertainty=None if production.uncertainty is None else np.asarray(production.uncertainty, dtype=float),
            metadata=production_metadata,
        )

    def predict_with_shadow(self, features: dict[str, np.ndarray]) -> ShadowInferenceResult:
        """Run production and shadow inference and return full diagnostics payload."""
        production = self._production_model.predict(features)
        shadow = self._shadow_model.predict(features)
        governance = self._build_governance_diagnostics(production=production, shadow=shadow)
        self._diagnostic_log.append(governance)
        return ShadowInferenceResult(production=production, shadow=shadow, governance_diagnostics=governance)

    @property
    def diagnostic_log(self) -> tuple[dict[str, Any], ...]:
        return tuple(self._diagnostic_log)

    def _build_governance_diagnostics(self, *, production: PredictionResult, shadow: PredictionResult) -> dict[str, Any]:
        prod_preds = np.asarray(production.predictions, dtype=float)
        shadow_preds = np.asarray(shadow.predictions, dtype=float)
        if prod_preds.shape != shadow_preds.shape:
            raise ValueError("production and shadow predictions must have matching shapes")

        prod_unc = np.zeros_like(prod_preds) if production.uncertainty is None else np.asarray(production.uncertainty, dtype=float)
        shadow_unc = np.zeros_like(shadow_preds) if shadow.uncertainty is None else np.asarray(shadow.uncertainty, dtype=float)
        if prod_unc.shape != shadow_unc.shape:
            raise ValueError("production and shadow uncertainty must have matching shapes")

        prediction_abs_delta = np.abs(prod_preds - shadow_preds)
        uncertainty_abs_delta = np.abs(prod_unc - shadow_unc)
        agreement = float(np.mean(np.sign(prod_preds) == np.sign(shadow_preds))) if prod_preds.size else 1.0

        windows: list[dict[str, float | int]] = []
        for idx, start in enumerate(range(0, prod_preds.shape[0], self._rolling_window)):
            stop = min(start + self._rolling_window, prod_preds.shape[0])
            if start >= stop:
                continue
            window_diag = DivergenceDiagnostic(
                window_index=idx,
                start=start,
                stop=stop,
                prediction_mean_abs_delta=float(np.mean(prediction_abs_delta[start:stop])),
                uncertainty_mean_abs_delta=float(np.mean(uncertainty_abs_delta[start:stop])),
                direction_agreement=float(np.mean(np.sign(prod_preds[start:stop]) == np.sign(shadow_preds[start:stop]))),
            )
            windows.append(window_diag.__dict__)

        return {
            "rolling_window": self._rolling_window,
            "prediction_mean_abs_delta": float(np.mean(prediction_abs_delta)) if prediction_abs_delta.size else 0.0,
            "uncertainty_mean_abs_delta": float(np.mean(uncertainty_abs_delta)) if uncertainty_abs_delta.size else 0.0,
            "direction_agreement": agreement,
            "windows": windows,
        }

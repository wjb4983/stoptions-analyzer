from __future__ import annotations

import hashlib
import json
import shutil
from dataclasses import asdict, dataclass, field, replace
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Literal, Protocol

import numpy as np

from config import BACKTEST_CACHE_DIR
from data_access.cache import _safe_ticker_name
from backtesting.scenario_toolkit import ScenarioSpec, build_custom_scenarios

from modeling_nextgen.calibration.probability import ProbabilityCalibrator
from modeling_nextgen.validation.quality_gates import evaluate_modeling_quality_gates, promotion_blocked
from models.deployment import PromotionGates
from models.regime_catalog import get_model_descriptor, list_models_for_leg, validate_model_leg_pairing
from models.registry import create_model
from models.robustness import build_robustness_scorecards
from backtesting.walk_forward import build_walk_forward_folds, run_walk_forward_optimization

DEFAULT_REGIME_TRAINING_OUTPUT_DIR = Path("data/regime_training_runs")


@dataclass(frozen=True)
class RegimeLegTrainingConfig:
    name: str
    model_type: str
    controls: dict[str, float]
    model_id: str = ""
    selected_model_id: str = ""
    hyperparameters: dict[str, Any] = field(default_factory=dict)
    architecture_spec: dict[str, Any] | None = None
    calibration_spec: dict[str, Any] | None = None
    event_process_spec: dict[str, Any] | None = None

    @property
    def resolved_model_id(self) -> str:
        model_id = self.model_id.strip()
        if model_id:
            return model_id
        return self.selected_model_id.strip()


@dataclass(frozen=True)
class RegimeTrainingRequest:
    schema_version: int
    regime_id: str
    regime_name: str
    legs: tuple[RegimeLegTrainingConfig, ...]
    model_choice: str
    training_window: dict[str, int]
    risk_limits: dict[str, float]
    output_dir: str | None = None
    training_data_settings: dict[str, Any] = field(default_factory=dict)

    @property
    def regime_label(self) -> str:
        """Backwards-compatible alias used by current UI code."""
        return self.regime_name


@dataclass(frozen=True)
class RegimeTrainingResult:
    run_id: str
    status: str
    metrics: dict[str, float]
    artifact_paths: dict[str, str]
    timestamps: dict[str, str]
    warnings: tuple[str, ...] = field(default_factory=tuple)
    errors: tuple[str, ...] = field(default_factory=tuple)
    error_payload: dict[str, Any] | None = None
    summary: str = ""
    metadata: dict[str, Any] = field(default_factory=dict)
    logs: tuple[str, ...] = field(default_factory=tuple)

    @property
    def started_at(self) -> str:
        return self.timestamps.get("started_at", "")

    @property
    def completed_at(self) -> str:
        return self.timestamps.get("completed_at", "")

    @property
    def artifact_path(self) -> str:
        return self.artifact_paths.get("manifest", "")


class RegimeTrainingAdapter(Protocol):
    def fit_and_backtest(self, request: RegimeTrainingRequest, run_dir: Path) -> "RegimeTrainingAdapterOutput":
        """Run fitting + backtest for a regime and return structured adapter output."""


@dataclass(frozen=True)
class AdapterIssue:
    level: Literal["warning", "error"]
    model_id: str
    message: str
    leg_name: str | None = None


@dataclass(frozen=True)
class TrainedArtifactLocations:
    model_weights: str
    calibration_object: str
    diagnostics: str


@dataclass(frozen=True)
class LegOutOfSampleMetrics:
    leg_name: str
    model_id: str
    metrics: dict[str, float]


@dataclass(frozen=True)
class RegimeTrainingAdapterOutput:
    per_leg_artifacts: dict[str, TrainedArtifactLocations]
    per_leg_oos_metrics: dict[str, LegOutOfSampleMetrics]
    portfolio_oos_metrics: dict[str, float]
    issues: tuple[AdapterIssue, ...] = ()
    adapter_name: str = ""
    candidate_leaderboards: dict[str, list[dict[str, Any]]] = field(default_factory=dict)
    champion_by_leg: dict[str, str] = field(default_factory=dict)
    governance_by_leg: dict[str, dict[str, Any]] = field(default_factory=dict)
    training_data_audit: dict[str, Any] = field(default_factory=dict)

    def warnings(self) -> tuple[str, ...]:
        return tuple(
            f"[{issue.model_id}] {issue.message}" for issue in self.issues if issue.level == "warning"
        )

    def errors(self) -> tuple[str, ...]:
        return tuple(
            f"[{issue.model_id}] {issue.message}" for issue in self.issues if issue.level == "error"
        )


class PlaceholderRegimeTrainingAdapter:
    """Stable placeholder kept for explicit dev/test mode only."""

    def fit_and_backtest(self, request: RegimeTrainingRequest, run_dir: Path) -> RegimeTrainingAdapterOutput:
        issues = (
            AdapterIssue(
                level="warning",
                model_id="placeholder",
                message="Running placeholder adapter output; use for explicit dev/test mode only.",
            ),
        )
        return RegimeTrainingAdapterOutput(
            per_leg_artifacts={},
            per_leg_oos_metrics={},
            portfolio_oos_metrics={"leg_count": float(len(request.legs))},
            issues=issues,
            adapter_name="placeholder",
        )


class RegistryBackedRegimeTrainingAdapter:
    """Production adapter that dispatches each leg by model id from model registry/catalog."""

    def fit_and_backtest(self, request: RegimeTrainingRequest, run_dir: Path) -> RegimeTrainingAdapterOutput:
        artifacts_dir = run_dir / "artifacts"
        artifacts_dir.mkdir(parents=True, exist_ok=True)
        per_leg_artifacts: dict[str, TrainedArtifactLocations] = {}
        per_leg_oos_metrics: dict[str, LegOutOfSampleMetrics] = {}
        issues: list[AdapterIssue] = []
        candidate_leaderboards: dict[str, list[dict[str, Any]]] = {}
        champion_by_leg: dict[str, str] = {}
        governance_by_leg: dict[str, dict[str, Any]] = {}
        training_data_bundle = self._load_training_data_bundle(request)

        if not bool(training_data_bundle.metadata.get("pass", True)):
            ratio = float(training_data_bundle.metadata.get("universe_pass_ratio", 0.0))
            threshold = float(training_data_bundle.metadata.get("required_universe_pass_ratio", 1.0))
            raise ValueError(
                f"Training data universe pass ratio {ratio:.3f} below required threshold {threshold:.3f}"
            )

        for idx, leg in enumerate(request.legs):
            leg_key = f"{idx:02d}_{leg.name.lower().replace(' ', '_')}"
            try:
                candidate_ids = self._candidate_model_ids(request, leg)
                for candidate_id in candidate_ids:
                    validate_model_leg_pairing(leg.model_type, candidate_id)

                candidate_payloads: dict[str, tuple[dict[str, np.ndarray], np.ndarray]] = {}
                for candidate_id in candidate_ids:
                    candidate_model = create_model(candidate_id)
                    candidate_payloads[candidate_id] = self._build_dataset(
                        candidate_model.required_feature_names(),
                        leg.controls,
                        request=request,
                        seed=idx,
                        training_data_bundle=training_data_bundle,
                    )

                sample_size = max(64, int(next(iter(candidate_payloads.values()))[1].size))
                folds = build_walk_forward_folds(
                    total_bars=sample_size,
                    train_bars=72,
                    validation_bars=24,
                    test_bars=24,
                    step_bars=24,
                    purge_window_bars=2,
                    embargo_window_bars=2,
                    label_horizon_bars=1,
                )
                walk_forward_result = run_walk_forward_optimization(
                    folds=folds,
                    parameter_candidates=[{"model_id": model_id} for model_id in candidate_ids],
                    evaluate_segment=lambda params, start, end: self._evaluate_segment(
                        model_id=str(params["model_id"]),
                        start=start,
                        end=end,
                        candidate_payloads=candidate_payloads,
                    ),
                    score_metric="sharpe",
                )
                leaderboard = self._build_candidate_leaderboard(
                    walk_forward_folds=walk_forward_result.folds,
                    candidate_ids=candidate_ids,
                )
                champion_model_id = str(leaderboard[0]["model_id"]) if leaderboard else candidate_ids[0]
                model = create_model(champion_model_id)

                features, labels = candidate_payloads[champion_model_id]
                split_idx = max(int(len(labels) * 0.7), 12)
                train_x = {name: values[:split_idx] for name, values in features.items()}
                test_x = {name: values[split_idx:] for name, values in features.items()}
                train_y = labels[:split_idx]
                test_y = labels[split_idx:]

                model.fit(train_x, train_y)
                train_probs = model.predict_proba(train_x)
                test_probs = model.predict_proba(test_x)

                calibrator = ProbabilityCalibrator(method="platt", n_bins=10)
                calibrator.fit(train_probs, np.where(train_y > 0, 1.0, 0.0))
                calibrated_test_probs = calibrator.transform(test_probs)
                test_labels_binary = np.where(test_y > 0, 1.0, 0.0)
                report = calibrator.report(test_probs, test_labels_binary, calibrated_probabilities=calibrated_test_probs)

                preds = np.where(calibrated_test_probs >= 0.5, 1.0, 0.0)
                metrics = {
                    "accuracy": float(np.mean(preds == test_labels_binary)),
                    "brier_score": float(report.brier_score),
                    "expected_calibration_error": float(report.expected_calibration_error),
                    "avg_confidence": float(np.mean(np.abs(calibrated_test_probs - 0.5) * 2.0)),
                    "oos_sample_size": float(test_labels_binary.size),
                }

                leg_dir = artifacts_dir / leg_key
                leg_dir.mkdir(parents=True, exist_ok=True)
                weights_path = leg_dir / "model_weights.json"
                calibration_path = leg_dir / "calibration_object.json"
                diagnostics_path = leg_dir / "diagnostics.json"
                weights_path.write_text(
                    json.dumps(
                        {
                            "model_id": champion_model_id,
                            "required_features": list(model.required_feature_names()),
                            "feature_importances": {k: float(v) for k, v in model.feature_importances_.items()},
                        },
                        indent=2,
                        sort_keys=True,
                    ),
                    encoding="utf-8",
                )
                calibration_path.write_text(
                    json.dumps(
                        {
                            "model_id": champion_model_id,
                            "method": report.method,
                            "sample_size": report.sample_size,
                            "expected_calibration_error": report.expected_calibration_error,
                            "brier_score": report.brier_score,
                        },
                        indent=2,
                        sort_keys=True,
                    ),
                    encoding="utf-8",
                )
                descriptor = get_model_descriptor(leg.model_type, champion_model_id)
                capability_tags = descriptor.capability_tags if descriptor is not None else frozenset()
                architecture_payload = leg.architecture_spec if isinstance(leg.architecture_spec, dict) else None
                diagnostics_payload = {
                    "leg_name": leg.name,
                    "model_id": champion_model_id,
                    "oos_metrics": metrics,
                    "controls": {k: float(v) for k, v in leg.controls.items()},
                    "candidate_leaderboard": leaderboard,
                }
                if architecture_payload is not None and capability_tags.intersection({"supports_architecture_spec", "needs_architecture_spec"}):
                    diagnostics_payload["architecture_spec"] = architecture_payload

                diagnostics_path.write_text(
                    json.dumps(
                        diagnostics_payload,
                        indent=2,
                        sort_keys=True,
                    ),
                    encoding="utf-8",
                )

                per_leg_artifacts[leg.name] = TrainedArtifactLocations(
                    model_weights=str(weights_path),
                    calibration_object=str(calibration_path),
                    diagnostics=str(diagnostics_path),
                )
                per_leg_oos_metrics[leg.name] = LegOutOfSampleMetrics(
                    leg_name=leg.name,
                    model_id=champion_model_id,
                    metrics=metrics,
                )

                robustness_scorecard = build_robustness_scorecards(
                    models={champion_model_id: model},
                    feature_payloads={champion_model_id: test_x},
                )
                champion_row = next(
                    (row for row in robustness_scorecard.get("models", []) if row.get("model_id") == champion_model_id),
                    {},
                )
                robust_oos_pass = float(walk_forward_result.aggregate_metrics.get("sharpe_mean", 0.0)) >= 0.0
                stress_pass = bool(champion_row.get("meets_minimum_threshold", False))
                scorecard = {
                    "dimensions": {
                        "robust_oos_performance": {"pass": robust_oos_pass},
                        "stress_resilience": {"pass": stress_pass},
                    }
                }
                baseline_rows = [
                    {
                        "max_drawdown": max(-0.25, -float(metrics.get("brier_score", 0.2))),
                        "downside_deviation": min(0.03, float(metrics.get("brier_score", 0.2)) / 10.0),
                        "rolling_drawdown_worst": max(-0.12, -float(metrics.get("brier_score", 0.2))),
                        "turnover_total": float(leg.controls.get("turnover_limit", 0.0)) * 100.0,
                        "slippage_bps": float(metrics.get("expected_calibration_error", 0.0)) * 100.0,
                    }
                ]
                gate_results = evaluate_modeling_quality_gates(
                    scorecard=scorecard,
                    calibration_report={
                        "fit_error": {"mae_bps_max": float(report.brier_score) * 100.0},
                        "stability": {"impact_coefficient_std_bps": float(report.expected_calibration_error) * 100.0},
                    },
                    baseline_rows=baseline_rows,
                )
                fallback_gate = _build_synthetic_fallback_quality_gate(
                    training_data_bundle.metadata,
                    request.training_data_settings,
                )
                gate_results.append(fallback_gate)
                blocked = promotion_blocked(gate_results)
                gates = PromotionGates()
                deployment_slots = {
                    "candidate": True,
                    "challenger": not blocked,
                    "champion": (
                        not blocked
                        and float(walk_forward_result.aggregate_metrics.get("sharpe_mean", 0.0))
                        >= gates.min_risk_adjusted_return_delta
                        and float(champion_row.get("robustness_score", 0.0)) >= gates.min_robustness_score
                        and len(champion_row.get("brittle_features", [])) <= gates.max_brittle_features
                    ),
                }
                governance_payload = {
                    "scorecard": scorecard,
                    "robustness_scorecard": robustness_scorecard,
                    "gates": gate_results,
                    "pass_fail": {gate["name"]: bool(gate.get("pass", False)) for gate in gate_results},
                    "promotion_eligible": not blocked,
                    "deployment_slot_eligibility": deployment_slots,
                }

                leaderboard_path = leg_dir / "candidate_leaderboard.json"
                governance_path = leg_dir / "governance.json"
                champion_path = leg_dir / "champion.json"
                leaderboard_path.write_text(json.dumps(leaderboard, indent=2, sort_keys=True), encoding="utf-8")
                governance_path.write_text(json.dumps(governance_payload, indent=2, sort_keys=True), encoding="utf-8")
                champion_path.write_text(
                    json.dumps(
                        {
                            "model_id": champion_model_id,
                            "promotion_eligible": governance_payload["promotion_eligible"],
                            "deployment_slot_eligibility": deployment_slots,
                        },
                        indent=2,
                        sort_keys=True,
                    ),
                    encoding="utf-8",
                )
                issues.append(
                    AdapterIssue(
                        level="warning",
                        model_id=champion_model_id,
                        leg_name=leg.name,
                        message=f"walk-forward evaluated {len(candidate_ids)} candidate(s)",
                    )
                )
                per_leg_oos_metrics[leg.name].metrics["promotion_eligible"] = float(
                    1.0 if governance_payload["promotion_eligible"] else 0.0
                )
                per_leg_oos_metrics[leg.name].metrics["walk_forward_folds"] = float(len(walk_forward_result.folds))
                per_leg_oos_metrics[leg.name].metrics["candidate_count"] = float(len(candidate_ids))
                per_leg_oos_metrics[leg.name].metrics["selected_sharpe_mean"] = float(
                    walk_forward_result.aggregate_metrics.get("sharpe_mean", 0.0)
                )

                candidate_leaderboards[leg.name] = leaderboard
                champion_by_leg[leg.name] = champion_model_id
                governance_by_leg[leg.name] = governance_payload
            except Exception as exc:
                issues.append(
                    AdapterIssue(
                        level="error",
                        model_id=request.model_choice,
                        leg_name=leg.name,
                        message=f"leg '{leg.name}' failed: {exc}",
                    )
                )

        portfolio_oos_metrics = self._aggregate_portfolio_metrics(per_leg_oos_metrics)
        return RegimeTrainingAdapterOutput(
            per_leg_artifacts=per_leg_artifacts,
            per_leg_oos_metrics=per_leg_oos_metrics,
            portfolio_oos_metrics=portfolio_oos_metrics,
            issues=tuple(issues),
            adapter_name="registry_backed",
            candidate_leaderboards=candidate_leaderboards,
            champion_by_leg=champion_by_leg,
            governance_by_leg=governance_by_leg,
            training_data_audit=training_data_bundle.metadata,
        )

    @staticmethod
    def _evaluate_segment(
        *,
        model_id: str,
        start: int,
        end: int,
        candidate_payloads: dict[str, tuple[dict[str, np.ndarray], np.ndarray]],
    ) -> dict[str, Any]:
        model = create_model(model_id)
        features, labels = candidate_payloads[model_id]
        start_idx = max(0, int(start))
        end_idx = min(int(end), int(labels.shape[0]))
        if end_idx - start_idx <= 2:
            return {"metrics": {"sharpe": -1.0, "accuracy": 0.0}, "equity": []}
        seg_x = {name: values[start_idx:end_idx] for name, values in features.items()}
        seg_y = labels[start_idx:end_idx]
        model.fit(seg_x, seg_y)
        probs = np.asarray(model.predict_proba(seg_x), dtype=float)
        binary_y = np.where(seg_y > 0, 1.0, 0.0)
        preds = np.where(probs >= 0.5, 1.0, 0.0)
        pnl = np.where(preds == binary_y, 1.0, -1.0)
        sharpe = float(np.mean(pnl) / max(1e-8, np.std(pnl))) if pnl.size > 1 else 0.0
        return {
            "metrics": {
                "accuracy": float(np.mean(preds == binary_y)),
                "sharpe": sharpe,
            },
            "equity": [
                {"timestamp": int(start_idx + i), "equity": float(np.sum(pnl[: i + 1]))}
                for i in range(int(pnl.size))
            ],
        }

    @staticmethod
    def _load_training_data_bundle(request: RegimeTrainingRequest) -> "TrainingDataBundle":
        settings = dict(request.training_data_settings or {})
        symbols = [str(s).strip().upper() for s in settings.get("universe_symbols", []) if str(s).strip()]
        required_years = max(1, int(settings.get("required_history_years", 5)))
        cache_root = Path(str(settings.get("cache_root", BACKTEST_CACHE_DIR)))
        allow_synthetic_fallback = bool(settings.get("allow_synthetic_fallback", False))
        bundle = _load_returns_from_cache(
            symbols=symbols,
            cache_root=cache_root,
            min_usable_history_years=required_years,
            required_universe_pass_ratio=float(
                settings.get("required_universe_pass_ratio", settings.get("min_universe_pass_ratio", 1.0))
            ),
        )
        bundle.metadata["allow_synthetic_fallback"] = allow_synthetic_fallback
        bundle.metadata["synthetic_fallback_used"] = False
        return bundle

    @staticmethod
    def _build_candidate_leaderboard(
        *,
        walk_forward_folds: list[dict[str, Any]],
        candidate_ids: list[str],
    ) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        for candidate_id in candidate_ids:
            scores: list[float] = []
            oos_scores: list[float] = []
            for fold in walk_forward_folds:
                diagnostics = fold.get("diagnostics", [])
                for diag in diagnostics:
                    params = diag.get("params", {})
                    if str(params.get("model_id", "")) != candidate_id:
                        continue
                    scores.append(float(diag.get("validation_score", 0.0)))
                    oos_scores.append(float(fold.get("oos_metrics", {}).get("sharpe", 0.0)))
            rows.append(
                {
                    "model_id": candidate_id,
                    "validation_score_mean": float(np.mean(scores)) if scores else float("-inf"),
                    "validation_score_std": float(np.std(scores, ddof=0)) if scores else 0.0,
                    "oos_sharpe_mean": float(np.mean(oos_scores)) if oos_scores else 0.0,
                    "folds_evaluated": len(scores),
                }
            )
        rows.sort(
            key=lambda row: (
                -float(row["validation_score_mean"]),
                -float(row["oos_sharpe_mean"]),
                str(row["model_id"]),
            )
        )
        return rows

    @staticmethod
    def _build_dataset(
        required_features: tuple[str, ...],
        controls: dict[str, float],
        *,
        request: RegimeTrainingRequest,
        seed: int,
        training_data_bundle: "TrainingDataBundle",
        sample_size: int = 160,
    ) -> tuple[dict[str, np.ndarray], np.ndarray]:
        rng_seed = int(sum(abs(float(v)) for v in controls.values()) * 10_000) + seed
        rng = np.random.default_rng(rng_seed)
        settings = dict(request.training_data_settings or {})
        base_returns = training_data_bundle.returns
        if base_returns.size < 32:
            allow_synthetic_fallback = bool(training_data_bundle.metadata.get("allow_synthetic_fallback", False))
            diagnostics = {
                "symbols_requested": training_data_bundle.metadata.get("symbols_requested", []),
                "symbols_used": training_data_bundle.metadata.get("symbols_used", []),
                "symbols_excluded": training_data_bundle.metadata.get("symbols_excluded", []),
                "universe_pass_ratio": training_data_bundle.metadata.get("universe_pass_ratio", 0.0),
                "required_universe_pass_ratio": training_data_bundle.metadata.get("required_universe_pass_ratio", 1.0),
                "minimum_usable_history_years": training_data_bundle.metadata.get("minimum_usable_history_years", 0),
                "observed_return_samples": int(base_returns.size),
                "minimum_required_return_samples": 32,
            }
            if not allow_synthetic_fallback:
                raise TrainingDataInsufficientError(
                    "INSUFFICIENT_REAL_HISTORY: observed return samples below minimum and synthetic fallback disabled; "
                    f"coverage_diagnostics={json.dumps(diagnostics, sort_keys=True)}"
                )
            training_data_bundle.metadata["synthetic_fallback_used"] = True
            training_data_bundle.metadata["synthetic_fallback_reason"] = "insufficient_real_history"
            base_returns = rng.normal(loc=0.0, scale=0.01, size=max(sample_size, 160)).astype(float)

        scenario_rows = _build_regime_training_scenarios(settings, n_assets=1)
        if scenario_rows:
            synthetic = [base_returns]
            for payload in scenario_rows:
                adjusted = _apply_training_path_adjustments(base_returns, payload.get("path_adjustments"))
                shift = float(payload.get("spec").params.get("returns_shift", 0.0)) if payload.get("spec") else 0.0
                synthetic.append(adjusted + shift)
            data = np.concatenate(synthetic)
        else:
            data = base_returns

        data = data[-max(sample_size, 160):]
        if data.size < sample_size:
            pad = rng.normal(loc=float(np.mean(data) if data.size else 0.0), scale=max(1e-6, float(np.std(data) if data.size else 0.01)), size=sample_size - data.size)
            data = np.concatenate([pad, data])

        features: dict[str, np.ndarray] = {}
        n = data.size
        for idx, feature_name in enumerate(required_features):
            window = max(2, (idx % 20) + 2)
            rolled = np.convolve(data, np.ones(window) / window, mode="same")
            noise = rng.normal(loc=0.0, scale=0.001 + idx * 0.0001, size=n)
            features[feature_name] = (rolled + noise).astype(float)

        latent = np.zeros(n, dtype=float)
        for feature_values in features.values():
            latent += feature_values
        latent += rng.normal(loc=0.0, scale=0.25, size=n)
        labels = np.where(latent >= np.median(latent), 1.0, -1.0)
        return features, labels

    @staticmethod
    def _select_model_id(request: RegimeTrainingRequest, leg: RegimeLegTrainingConfig) -> str:
        descriptors = list_models_for_leg(leg.model_type)
        if not descriptors:
            raise ValueError(f"No catalog entries configured for leg type '{leg.model_type}'")

        mode = request.model_choice.strip().lower()
        selected_model_id = leg.resolved_model_id.lower()
        allowed = {item.model_name for item in descriptors}

        if mode in {"single_model", "auto_model_search", "", "auto"}:
            if selected_model_id and selected_model_id in allowed:
                return selected_model_id
            return descriptors[0].model_name

        if mode == "ensemble":
            for descriptor in descriptors:
                if descriptor.model_name == "meta_label_classifier":
                    return descriptor.model_name
            return descriptors[0].model_name

        if mode in allowed:
            return mode

        return descriptors[0].model_name

    @staticmethod
    def _candidate_model_ids(request: RegimeTrainingRequest, leg: RegimeLegTrainingConfig) -> list[str]:
        descriptors = list_models_for_leg(leg.model_type)
        if not descriptors:
            raise ValueError(f"No catalog entries configured for leg type '{leg.model_type}'")
        mode = request.model_choice.strip().lower()
        selected_model_id = leg.resolved_model_id.lower()
        allowed = [item.model_name for item in descriptors]
        if mode in {"auto_model_search", "auto"}:
            return allowed
        if mode in {"single_model", "", "placeholder", "dev", "test"}:
            if selected_model_id and selected_model_id in allowed:
                return [selected_model_id]
            return [allowed[0]]
        if mode == "ensemble":
            return ["meta_label_classifier"] if "meta_label_classifier" in allowed else [allowed[0]]
        if mode in allowed:
            return [mode]
        return [allowed[0]]

    @staticmethod
    def _aggregate_portfolio_metrics(
        per_leg_oos_metrics: dict[str, LegOutOfSampleMetrics],
    ) -> dict[str, float]:
        if not per_leg_oos_metrics:
            return {"legs_trained": 0.0}

        metric_keys = set()
        for leg_metrics in per_leg_oos_metrics.values():
            metric_keys.update(leg_metrics.metrics.keys())

        aggregate: dict[str, float] = {"legs_trained": float(len(per_leg_oos_metrics))}
        for metric_key in sorted(metric_keys):
            values = [float(item.metrics[metric_key]) for item in per_leg_oos_metrics.values() if metric_key in item.metrics]
            if values:
                aggregate[f"portfolio_avg_{metric_key}"] = float(np.mean(values))
        return aggregate



@dataclass(frozen=True)
class _SymbolReturns:
    symbol: str
    timestamps_ms: np.ndarray
    returns: np.ndarray
    earliest_ts_ms: int
    latest_ts_ms: int


@dataclass(frozen=True)
class TrainingDataBundle:
    returns: np.ndarray
    metadata: dict[str, Any]


class TrainingDataInsufficientError(ValueError):
    """Raised when cache-backed training history is insufficient and fallback is disabled."""


def _normalize_timestamps_ms(values: np.ndarray) -> np.ndarray:
    ts = np.asarray(values, dtype=np.int64).reshape(-1)
    if not ts.size:
        return ts
    max_abs = int(np.max(np.abs(ts)))
    # If epoch seconds are provided, convert to milliseconds.
    if max_abs < 10**11:
        ts = ts * 1000
    return ts


def _to_iso(ts_ms: int | None) -> str | None:
    if ts_ms is None:
        return None
    return datetime.fromtimestamp(ts_ms / 1000, tz=timezone.utc).isoformat()


def _resolve_effective_start_ms(*, now: datetime, earliest_cached_ms: int) -> int:
    five_year_cutoff = int((now - timedelta(days=365 * 5)).timestamp() * 1000)
    return min(five_year_cutoff, int(earliest_cached_ms))


def _load_symbol_returns_from_cache(*, symbol: str, cache_root: Path) -> _SymbolReturns | None:
    safe = _safe_ticker_name(symbol)
    symbol_dir = cache_root / safe / "1m"
    if not symbol_dir.exists():
        return None

    closes: list[np.ndarray] = []
    timestamps: list[np.ndarray] = []
    for path in sorted(symbol_dir.glob(f"{safe}_1m_*.npz")):
        try:
            with np.load(path, mmap_mode="r") as payload:
                close = np.asarray(payload.get("c"), dtype=float).reshape(-1)
                raw_ts = payload.get("t")
        except Exception:
            continue
        if raw_ts is None:
            continue
        ts = _normalize_timestamps_ms(np.asarray(raw_ts))
        length = min(int(close.size), int(ts.size))
        if length < 3:
            continue
        closes.append(close[:length])
        timestamps.append(ts[:length])

    if not closes or not timestamps:
        return None
    joined_close = np.concatenate(closes)
    joined_ts = np.concatenate(timestamps)
    order = np.argsort(joined_ts, kind="mergesort")
    joined_close = joined_close[order]
    joined_ts = joined_ts[order]

    pct_returns = np.diff(joined_close) / np.where(np.abs(joined_close[:-1]) < 1e-8, 1.0, joined_close[:-1])
    pct_returns = np.clip(pct_returns, -0.25, 0.25)
    return_ts = joined_ts[1:]
    finite_mask = np.isfinite(pct_returns) & np.isfinite(return_ts)
    if not np.any(finite_mask):
        return None
    filtered_returns = np.asarray(pct_returns[finite_mask], dtype=float)
    filtered_ts = np.asarray(return_ts[finite_mask], dtype=np.int64)
    return _SymbolReturns(
        symbol=symbol,
        timestamps_ms=filtered_ts,
        returns=filtered_returns,
        earliest_ts_ms=int(np.min(joined_ts)),
        latest_ts_ms=int(np.max(filtered_ts)),
    )


def _load_returns_from_cache(
    *,
    symbols: list[str],
    cache_root: Path,
    min_usable_history_years: int,
    required_universe_pass_ratio: float = 1.0,
    now: datetime | None = None,
) -> TrainingDataBundle:
    normalized_symbols = [str(symbol).strip().upper() for symbol in symbols if str(symbol).strip()]
    if not normalized_symbols:
        return TrainingDataBundle(
            returns=np.array([], dtype=float),
            metadata={
                "symbols_requested": [],
                "symbols_used": [],
                "symbols_excluded": [],
                "effective_training_start": None,
                "effective_training_end": None,
                "per_symbol": {},
                "universe_pass_ratio": 1.0,
                "required_universe_pass_ratio": float(required_universe_pass_ratio),
                "minimum_usable_history_years": int(min_usable_history_years),
            },
        )

    now_utc = now or datetime.now(timezone.utc)
    min_years = max(0.0, float(min_usable_history_years))
    required_pass_ratio = max(0.0, min(1.0, float(required_universe_pass_ratio)))

    per_symbol: dict[str, dict[str, Any]] = {}
    used_symbols: list[str] = []
    excluded: list[dict[str, Any]] = []
    usable_returns: list[np.ndarray] = []
    training_start_ms: int | None = None
    training_end_ms: int | None = None

    for symbol in normalized_symbols:
        loaded = _load_symbol_returns_from_cache(symbol=symbol, cache_root=cache_root)
        if loaded is None:
            excluded.append({"symbol": symbol, "reason": "missing_or_unreadable_cache"})
            per_symbol[symbol] = {"bars_used": 0, "years_used": 0.0, "excluded": True}
            continue

        effective_start_ms = _resolve_effective_start_ms(now=now_utc, earliest_cached_ms=loaded.earliest_ts_ms)
        usable_mask = loaded.timestamps_ms >= effective_start_ms
        usable_ts = loaded.timestamps_ms[usable_mask]
        usable_rets = loaded.returns[usable_mask]

        if usable_ts.size >= 2:
            years_used = float((int(usable_ts[-1]) - int(usable_ts[0])) / (1000 * 60 * 60 * 24 * 365.25))
        else:
            years_used = 0.0
        bars_used = int(usable_rets.size)

        detail = {
            "effective_start": _to_iso(int(effective_start_ms)),
            "effective_end": _to_iso(int(usable_ts[-1])) if usable_ts.size else None,
            "bars_used": bars_used,
            "years_used": years_used,
            "excluded": False,
        }

        history_tolerance_years = 2.0 / 365.25
        if years_used + history_tolerance_years < min_years:
            detail["excluded"] = True
            detail["exclusion_reason"] = "insufficient_usable_history"
            excluded.append(
                {
                    "symbol": symbol,
                    "reason": "insufficient_usable_history",
                    "bars_used": bars_used,
                    "years_used": years_used,
                }
            )
            per_symbol[symbol] = detail
            continue

        used_symbols.append(symbol)
        per_symbol[symbol] = detail
        usable_returns.append(usable_rets)
        if usable_ts.size:
            symbol_start = int(np.min(usable_ts))
            symbol_end = int(np.max(usable_ts))
            training_start_ms = symbol_start if training_start_ms is None else min(training_start_ms, symbol_start)
            training_end_ms = symbol_end if training_end_ms is None else max(training_end_ms, symbol_end)

    total = len(normalized_symbols)
    pass_ratio = float(len(used_symbols) / total) if total else 1.0
    merged_returns = np.concatenate(usable_returns) if usable_returns else np.array([], dtype=float)

    metadata = {
        "symbols_requested": normalized_symbols,
        "symbols_used": used_symbols,
        "symbols_excluded": excluded,
        "effective_training_start": _to_iso(training_start_ms),
        "effective_training_end": _to_iso(training_end_ms),
        "per_symbol": per_symbol,
        "universe_pass_ratio": pass_ratio,
        "required_universe_pass_ratio": required_pass_ratio,
        "minimum_usable_history_years": min_years,
        "pass": pass_ratio >= required_pass_ratio,
    }
    return TrainingDataBundle(returns=merged_returns[np.isfinite(merged_returns)], metadata=metadata)


def _build_regime_training_scenarios(settings: dict[str, Any], *, n_assets: int) -> list[dict[str, Any]]:
    scenario_settings = settings.get("scenario_settings", [])
    if not isinstance(scenario_settings, list) or not scenario_settings:
        return []
    specs: list[ScenarioSpec] = []
    for row in scenario_settings:
        if not isinstance(row, dict):
            continue
        specs.append(
            ScenarioSpec(
                name=str(row.get("name", "scenario")),
                scenario_type=str(row.get("scenario_type", "vol_shock")),
                params=dict(row.get("params", {})),
            )
        )
    if not specs:
        return []
    return build_custom_scenarios(specs=specs, n_assets=n_assets)


def _apply_training_path_adjustments(values: np.ndarray, path_adjustments: dict[str, Any] | None) -> np.ndarray:
    adjusted = np.asarray(values, dtype=float).copy()
    if not path_adjustments or adjusted.size == 0:
        return adjusted
    n_periods = adjusted.shape[0]
    crash_len = min(n_periods, int(path_adjustments.get("crash_len", 0)))
    rebound_len = min(max(0, n_periods - crash_len), int(path_adjustments.get("rebound_len", 0)))
    crash_shift = float(path_adjustments.get("crash_shift", 0.0))
    rebound_shift = float(path_adjustments.get("rebound_shift", 0.0))
    if crash_len > 0:
        adjusted[:crash_len] += crash_shift
    if rebound_len > 0:
        adjusted[crash_len : crash_len + rebound_len] += rebound_shift
    return adjusted


def _build_synthetic_fallback_quality_gate(
    training_data_audit: dict[str, Any],
    training_data_settings: dict[str, Any] | None,
) -> dict[str, Any]:
    settings = dict(training_data_settings or {})
    fallback_used = bool(training_data_audit.get("synthetic_fallback_used", False))
    fallback_allowed = bool(settings.get("allow_synthetic_fallback", False))
    fallback_intentional = (not fallback_used) or fallback_allowed
    reason = (
        "synthetic fallback not used"
        if not fallback_used
        else ("synthetic fallback used with explicit allow flag" if fallback_allowed else "synthetic fallback used without explicit allow flag")
    )
    return {
        "name": "synthetic_fallback_intentional",
        "pass": fallback_intentional,
        "reason": reason,
        "value": 1.0 if fallback_intentional else 0.0,
        "threshold": 1.0,
        "block_promotion": not fallback_intentional,
    }


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _deterministic_run_id(request: RegimeTrainingRequest) -> str:
    payload = asdict(request)
    payload["legs"] = [asdict(leg) for leg in request.legs]
    digest = hashlib.sha256(json.dumps(payload, sort_keys=True).encode("utf-8")).hexdigest()
    return digest[:12]


def validate_regime_spec(request: RegimeTrainingRequest) -> list[str]:
    errors: list[str] = []
    if int(request.schema_version) < 2:
        errors.append("schema_version must be >= 2")
    if not request.regime_id.strip():
        errors.append("regime_id is required")
    if not request.regime_name.strip():
        errors.append("regime_name is required")
    if not request.model_choice.strip():
        errors.append("model_choice is required")
    if not request.legs:
        errors.append("at least one leg is required")

    for idx, leg in enumerate(request.legs):
        if not leg.name.strip():
            errors.append(f"legs[{idx}].name is required")
        if not leg.model_type.strip():
            errors.append(f"legs[{idx}].model_type is required")
        if not isinstance(leg.hyperparameters, dict):
            errors.append(f"legs[{idx}].hyperparameters must be an object")

    retrain_days = request.training_window.get("retrain_frequency_days")
    if retrain_days is not None and int(retrain_days) <= 0:
        errors.append("training_window.retrain_frequency_days must be > 0")

    for key, value in request.risk_limits.items():
        if float(value) < 0:
            errors.append(f"risk_limits.{key} must be >= 0")

    errors.extend(validate_model_specific_specs(request))
    return errors


def validate_model_specific_specs(request: RegimeTrainingRequest) -> list[str]:
    errors: list[str] = []
    for idx, leg in enumerate(request.legs):
        model_id = leg.resolved_model_id.lower()
        descriptor = get_model_descriptor(leg.model_type, model_id) if model_id else None
        capability_tags = descriptor.capability_tags if descriptor is not None else frozenset()

        errors.extend(
            _validate_optional_spec_object(leg.architecture_spec, field_path=f"legs[{idx}].architecture_spec")
        )
        errors.extend(_validate_optional_spec_object(leg.calibration_spec, field_path=f"legs[{idx}].calibration_spec"))
        errors.extend(
            _validate_optional_spec_object(leg.event_process_spec, field_path=f"legs[{idx}].event_process_spec")
        )

        if isinstance(leg.architecture_spec, dict):
            errors.extend(
                _validate_architecture_spec(leg.architecture_spec, field_path=f"legs[{idx}].architecture_spec")
            )
        elif "needs_architecture_spec" in capability_tags:
            errors.extend(
                _validate_architecture_spec(leg.architecture_spec, field_path=f"legs[{idx}].architecture_spec")
            )
        if "needs_calibration_spec" in capability_tags:
            errors.extend(
                _validate_calibration_spec(leg.calibration_spec, field_path=f"legs[{idx}].calibration_spec")
            )
        if "needs_event_process_spec" in capability_tags:
            errors.extend(
                _validate_event_process_spec(leg.event_process_spec, field_path=f"legs[{idx}].event_process_spec")
            )
    return errors




def _coerce_int(value: Any, default: int = 0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _coerce_float(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default

def _validate_optional_spec_object(payload: dict[str, Any] | None, *, field_path: str) -> list[str]:
    if payload is None:
        return []
    if isinstance(payload, dict):
        return []
    return [f"{field_path} must be an object when provided"]


def _validate_architecture_spec(payload: dict[str, Any] | None, *, field_path: str) -> list[str]:
    if not isinstance(payload, dict):
        return [f"{field_path} is required for ANN/neural model legs"]

    errors: list[str] = []
    schema_version = payload.get("schema_version")
    if schema_version is not None and _coerce_int(schema_version, 0) < 1:
        errors.append(f"{field_path}.schema_version must be >= 1")

    layers = payload.get("layers")
    if not isinstance(layers, list) or not layers:
        errors.append(f"{field_path}.layers must be a non-empty list")
    else:
        allowed_layer_types = {"Dense", "Dropout", "Norm", "Activation"}
        for idx, layer in enumerate(layers):
            if not isinstance(layer, dict):
                errors.append(f"{field_path}.layers[{idx}] must be an object")
                continue
            layer_type = str(layer.get("type", "")).strip()
            if layer_type not in allowed_layer_types:
                errors.append(
                    f"{field_path}.layers[{idx}].type must be one of {sorted(allowed_layer_types)}"
                )
                continue
            if layer_type == "Dense":
                units = layer.get("units")
                if _coerce_int(units, 0) <= 0:
                    errors.append(f"{field_path}.layers[{idx}].units must be > 0")
                if not str(layer.get("activation", "")).strip():
                    errors.append(f"{field_path}.layers[{idx}].activation is required")
            if layer_type == "Dropout":
                rate = _coerce_float(layer.get("rate", -1.0), -1.0)
                if rate < 0.0 or rate >= 1.0:
                    errors.append(f"{field_path}.layers[{idx}].rate must be within [0, 1)")
            if layer_type == "Norm" and not str(layer.get("norm", "")).strip():
                errors.append(f"{field_path}.layers[{idx}].norm is required")
            if layer_type == "Activation" and not str(layer.get("name", "")).strip():
                errors.append(f"{field_path}.layers[{idx}].name is required")

    optimizer = payload.get("optimizer")
    if not isinstance(optimizer, dict):
        errors.append(f"{field_path}.optimizer must be an object")
    else:
        if not str(optimizer.get("name", "")).strip():
            errors.append(f"{field_path}.optimizer.name is required")
        if _coerce_float(optimizer.get("learning_rate", 0.0), 0.0) <= 0.0:
            errors.append(f"{field_path}.optimizer.learning_rate must be > 0")

    loss = payload.get("loss")
    if not isinstance(loss, dict) or not str(loss.get("name", "")).strip():
        errors.append(f"{field_path}.loss.name is required")

    scheduler = payload.get("scheduler")
    if not isinstance(scheduler, dict) or not str(scheduler.get("name", "")).strip():
        errors.append(f"{field_path}.scheduler.name is required")

    training = payload.get("training")
    if not isinstance(training, dict):
        errors.append(f"{field_path}.training must be an object")
    else:
        if _coerce_int(training.get("batch_size", 0), 0) <= 0:
            errors.append(f"{field_path}.training.batch_size must be > 0")
        if _coerce_int(training.get("epochs", 0), 0) <= 0:
            errors.append(f"{field_path}.training.epochs must be > 0")
        early = training.get("early_stopping")
        if not isinstance(early, dict):
            errors.append(f"{field_path}.training.early_stopping must be an object")
        else:
            if not isinstance(early.get("enabled"), bool):
                errors.append(f"{field_path}.training.early_stopping.enabled must be boolean")
            if _coerce_int(early.get("patience", 0), 0) <= 0:
                errors.append(f"{field_path}.training.early_stopping.patience must be > 0")

    return errors


def _validate_calibration_spec(payload: dict[str, Any] | None, *, field_path: str) -> list[str]:
    if not isinstance(payload, dict):
        return [f"{field_path} is required for volatility/surface calibration models"]
    errors: list[str] = []
    if not str(payload.get("model", "")).strip():
        errors.append(f"{field_path}.model is required")
    parameters = payload.get("parameters")
    if not isinstance(parameters, dict):
        errors.append(f"{field_path}.parameters must be an object")
    return errors


def _validate_event_process_spec(payload: dict[str, Any] | None, *, field_path: str) -> list[str]:
    if not isinstance(payload, dict):
        return [f"{field_path} is required for Hawkes/jump-intensity models"]
    errors: list[str] = []
    if not str(payload.get("process_type", "")).strip():
        errors.append(f"{field_path}.process_type is required")
    parameters = payload.get("parameters")
    if not isinstance(parameters, dict):
        errors.append(f"{field_path}.parameters must be an object")
    return errors


def save_regime_spec(request: RegimeTrainingRequest, run_dir: Path) -> Path:
    spec_path = run_dir / "regime_spec_snapshot.json"
    spec_path.write_text(json.dumps(asdict(request), indent=2, sort_keys=True), encoding="utf-8")
    return spec_path


def compute_summary_metrics(
    adapter_output: RegimeTrainingAdapterOutput,
    request: RegimeTrainingRequest,
) -> tuple[dict[str, float], str]:
    metrics = {key: float(value) for key, value in adapter_output.portfolio_oos_metrics.items()}
    summary = (
        f"Trained {len(request.legs)} leg(s) for regime '{request.regime_name}' "
        f"using model choice '{request.model_choice}'."
    )
    return metrics, summary


def _flatten_artifact_locations(
    output: RegimeTrainingAdapterOutput,
) -> dict[str, str]:
    paths: dict[str, str] = {}
    for leg_name, artifacts in output.per_leg_artifacts.items():
        safe_leg = leg_name.lower().replace(" ", "_")
        paths[f"{safe_leg}_model_weights"] = artifacts.model_weights
        paths[f"{safe_leg}_calibration_object"] = artifacts.calibration_object
        paths[f"{safe_leg}_diagnostics"] = artifacts.diagnostics
    return paths




def _persist_cache_audit_report(request: RegimeTrainingRequest, run_dir: Path) -> str | None:
    settings = dict(request.training_data_settings or {})
    source_raw = str(settings.get("cache_audit_report", "")).strip()
    if not source_raw:
        return None
    source = Path(source_raw)
    if not source.exists() or not source.is_file():
        return None
    target = run_dir / "cache_audit_report.json"
    shutil.copyfile(source, target)
    return str(target)

def write_regime_training_manifest(
    *,
    run_dir: Path,
    request: RegimeTrainingRequest,
    result: RegimeTrainingResult,
) -> Path:
    manifest_path = run_dir / "manifest.json"
    manifest = {
        "run_id": result.run_id,
        "status": result.status,
        "request": asdict(request),
        "metrics": result.metrics,
        "timestamps": result.timestamps,
        "summary": result.summary,
        "warnings": list(result.warnings),
        "errors": list(result.errors),
        "error_payload": result.error_payload,
        "artifact_paths": result.artifact_paths,
        "metadata": result.metadata,
        "logs": list(result.logs),
    }
    manifest_path.write_text(json.dumps(manifest, indent=2, sort_keys=True), encoding="utf-8")
    return manifest_path


def run_regime_training(
    request: RegimeTrainingRequest,
    output_dir: str | Path | None = None,
    adapter: RegimeTrainingAdapter | None = None,
) -> RegimeTrainingResult:
    started_at = _utc_now_iso()
    run_id = _deterministic_run_id(request)
    resolved_output_dir = request.output_dir or output_dir
    output_root = (
        Path(resolved_output_dir) if resolved_output_dir is not None else DEFAULT_REGIME_TRAINING_OUTPUT_DIR
    )
    run_dir = output_root / f"run_{run_id}"
    run_dir.mkdir(parents=True, exist_ok=True)

    spec_path = save_regime_spec(request, run_dir)
    cache_audit_artifact = _persist_cache_audit_report(request, run_dir)
    validation_errors = validate_regime_spec(request)
    if validation_errors:
        completed_at = _utc_now_iso()
        error_payload = {
            "code": "INVALID_REGIME_SPEC",
            "stage": "validate_regime_spec",
            "errors": validation_errors,
        }
        result = RegimeTrainingResult(
            run_id=run_id,
            status="failed",
            metrics={},
            artifact_paths={"spec": str(spec_path), **({"cache_audit_report": cache_audit_artifact} if cache_audit_artifact else {})},
            timestamps={"started_at": started_at, "completed_at": completed_at},
            warnings=(),
            errors=tuple(validation_errors),
            error_payload=error_payload,
            summary="Regime training request validation failed.",
            metadata={"regime_id": request.regime_id, "regime_name": request.regime_name},
            logs=("saved config snapshot", "validation failed"),
        )
        manifest_path = write_regime_training_manifest(run_dir=run_dir, request=request, result=result)
        return replace(result, artifact_paths={**result.artifact_paths, "manifest": str(manifest_path)})

    runner = adapter or RegistryBackedRegimeTrainingAdapter()
    if request.model_choice.strip().lower() in {"placeholder", "dev", "test"}:
        runner = adapter or PlaceholderRegimeTrainingAdapter()
    try:
        adapter_output = runner.fit_and_backtest(request, run_dir)
        metrics, summary = compute_summary_metrics(adapter_output, request)
        completed_at = _utc_now_iso()
        warnings = adapter_output.warnings()
        adapter_errors = adapter_output.errors()
        artifacts = {"spec": str(spec_path), **_flatten_artifact_locations(adapter_output), **({"cache_audit_report": cache_audit_artifact} if cache_audit_artifact else {})}
        oos_metrics_payload = {
            leg_name: {
                "model_id": leg_metrics.model_id,
                "metrics": {k: float(v) for k, v in leg_metrics.metrics.items()},
            }
            for leg_name, leg_metrics in adapter_output.per_leg_oos_metrics.items()
        }
        synthetic_fallback_used = bool(adapter_output.training_data_audit.get("synthetic_fallback_used", False))
        logs_list = [
            "saved config snapshot",
            f"adapter={adapter_output.adapter_name or runner.__class__.__name__}",
            "computed summary metrics",
        ]
        warnings_list = list(warnings)
        if synthetic_fallback_used:
            logs_list.append("synthetic_fallback_used=true")
            warnings_list.append("training_data used synthetic fallback due to insufficient real history")
        logs = tuple(logs_list)
        warnings = tuple(warnings_list)
        status = "failed" if adapter_errors else "success"
        result = RegimeTrainingResult(
            run_id=run_id,
            status=status,
            metrics=metrics,
            artifact_paths=artifacts,
            timestamps={"started_at": started_at, "completed_at": completed_at},
            warnings=warnings,
            errors=adapter_errors,
            error_payload=(
                {
                    "code": "PARTIAL_TRAINING_FAILURE",
                    "stage": "fit_and_backtest",
                    "errors": list(adapter_errors),
                }
                if adapter_errors
                else None
            ),
            summary=summary,
            metadata={
                "regime_id": request.regime_id,
                "regime_name": request.regime_name,
                "model_choice": request.model_choice,
                "oos_metrics": oos_metrics_payload,
                "candidate_leaderboard": adapter_output.candidate_leaderboards,
                "champion_by_leg": adapter_output.champion_by_leg,
                "governance_by_leg": adapter_output.governance_by_leg,
                "training_data_audit": adapter_output.training_data_audit,
                "synthetic_fallback_used": synthetic_fallback_used,
            },
            logs=logs,
        )
    except Exception as exc:
        completed_at = _utc_now_iso()
        error_payload = {
            "code": "TRAINING_EXECUTION_FAILED",
            "stage": "fit_and_backtest",
            "message": str(exc),
            "exception_type": type(exc).__name__,
        }
        result = RegimeTrainingResult(
            run_id=run_id,
            status="failed",
            metrics={},
            artifact_paths={"spec": str(spec_path), **({"cache_audit_report": cache_audit_artifact} if cache_audit_artifact else {})},
            timestamps={"started_at": started_at, "completed_at": completed_at},
            warnings=(),
            errors=(str(exc),),
            error_payload=error_payload,
            summary="Regime training failed during fit/backtest stage.",
            metadata={"regime_id": request.regime_id, "regime_name": request.regime_name},
            logs=("saved config snapshot", "fit_and_backtest failed"),
        )

    manifest_path = write_regime_training_manifest(run_dir=run_dir, request=request, result=result)
    return replace(result, artifact_paths={**result.artifact_paths, "manifest": str(manifest_path)})


def execute_regime_training_pipeline(
    request: RegimeTrainingRequest,
    *,
    adapter: RegimeTrainingAdapter | None = None,
    output_dir: str | Path | None = None,
) -> RegimeTrainingResult:
    """UI seam point for Create Regime and Research Lab orchestration."""
    return run_regime_training(request=request, output_dir=output_dir, adapter=adapter)

from __future__ import annotations

import hashlib
import json
import os
from dataclasses import asdict, dataclass, field, replace
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Literal, Protocol

import numpy as np

from modeling_nextgen.calibration.probability import ProbabilityCalibrator
from modeling_nextgen.validation.quality_gates import evaluate_modeling_quality_gates, promotion_blocked
from models.deployment import PromotionGates
from models.regime_catalog import list_models_for_leg, validate_model_leg_pairing
from models.registry import create_model
from models.robustness import build_robustness_scorecards
from backtesting.walk_forward import build_walk_forward_folds, run_walk_forward_optimization
from config import API_KEY_PATH, DATA_DIR
from data_access.api_client import MassiveApiClient

DEFAULT_REGIME_TRAINING_OUTPUT_DIR = Path("data/regime_training_runs")
DEFAULT_REGIME_UNIVERSE_CACHE_DIR = DATA_DIR / "regime_universe_cache"
DEFAULT_REGIME_CACHE_POLICY = {"min_years": 5, "bar_size": "1d"}
DEFAULT_REGIME_SCENARIO_SETTINGS = {
    "panic_crash": {"enabled": True, "returns_multiplier": 1.8, "returns_shift": -0.015},
    "bull_low_vol": {"enabled": True, "returns_multiplier": 0.55, "returns_shift": 0.0025},
    "broad_market": {"enabled": True, "returns_multiplier": 1.0, "returns_shift": 0.0},
    "intraday_whipsaw": {"enabled": True, "returns_multiplier": 1.45, "returns_shift": -0.0006},
    "few_minute_shock": {"enabled": True, "returns_multiplier": 2.2, "returns_shift": -0.002},
}


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
    universe_tickers: tuple[str, ...] = ()
    cache_policy: dict[str, Any] = field(default_factory=dict)
    scenario_settings: dict[str, Any] = field(default_factory=dict)
    output_dir: str | None = None

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
    scenario_diagnostics: dict[str, Any] = field(default_factory=dict)
    data_readiness: dict[str, Any] = field(default_factory=dict)

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
        market_context = _build_market_context(request)

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
                        seed=idx,
                        market_returns=market_context["returns"],
                    )

                folds = build_walk_forward_folds(
                    total_bars=160,
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
                diagnostics_path.write_text(
                    json.dumps(
                        {
                            "leg_name": leg.name,
                            "model_id": champion_model_id,
                            "oos_metrics": metrics,
                            "controls": {k: float(v) for k, v in leg.controls.items()},
                            "candidate_leaderboard": leaderboard,
                        },
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
                    "scenario_diagnostics": market_context["scenario_diagnostics"],
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
            scenario_diagnostics=market_context["scenario_diagnostics"],
            data_readiness=market_context["data_readiness"],
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
        seed: int,
        market_returns: np.ndarray | None = None,
        sample_size: int = 160,
    ) -> tuple[dict[str, np.ndarray], np.ndarray]:
        if market_returns is not None and market_returns.size:
            aligned = np.asarray(market_returns, dtype=float)
            if aligned.ndim == 2 and aligned.shape[0] >= 24:
                return _build_market_conditioned_dataset(required_features, aligned, seed=seed)
        rng_seed = int(sum(abs(float(v)) for v in controls.values()) * 10_000) + seed
        rng = np.random.default_rng(rng_seed)
        features: dict[str, np.ndarray] = {}
        for feature_name in required_features:
            features[feature_name] = rng.normal(loc=0.0, scale=1.0, size=sample_size).astype(float)

        latent = np.zeros(sample_size, dtype=float)
        for feature_values in features.values():
            latent += feature_values
        latent += rng.normal(loc=0.0, scale=0.5, size=sample_size)
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


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _load_api_key() -> str:
    env_key = os.getenv("MASSIVE_API_KEY", "").strip()
    if env_key:
        return env_key
    if API_KEY_PATH.exists():
        return API_KEY_PATH.read_text(encoding="utf-8").strip()
    return ""


def _safe_ticker_name(ticker: str) -> str:
    return "".join(char if char.isalnum() else "_" for char in ticker.upper())


def _build_market_context(request: RegimeTrainingRequest) -> dict[str, Any]:
    returns, readiness = _load_or_fetch_universe_returns(request)
    scenario_settings = {**DEFAULT_REGIME_SCENARIO_SETTINGS, **request.scenario_settings}
    scenario_diag = _evaluate_market_scenarios(returns, settings=scenario_settings)
    return {"returns": returns, "data_readiness": readiness, "scenario_diagnostics": scenario_diag}


def _load_or_fetch_universe_returns(request: RegimeTrainingRequest) -> tuple[np.ndarray, dict[str, Any]]:
    tickers = [str(t).strip().upper() for t in request.universe_tickers if str(t).strip()]
    if not tickers:
        tickers = [request.regime_id.upper()]
    min_years = int((request.cache_policy or {}).get("min_years", DEFAULT_REGIME_CACHE_POLICY["min_years"]))
    target_days = max(252, min_years * 252)
    cache_dir = DEFAULT_REGIME_UNIVERSE_CACHE_DIR
    cache_dir.mkdir(parents=True, exist_ok=True)
    api_key = _load_api_key()
    api_client = MassiveApiClient(api_key) if api_key else None
    now = date.today()
    start_cutoff = now - timedelta(days=max(365, min_years * 365))
    closes_by_ticker: dict[str, np.ndarray] = {}
    fetches: dict[str, str] = {}

    for ticker in tickers:
        cache_path = cache_dir / f"{_safe_ticker_name(ticker)}.json"
        existing = _load_cached_closes(cache_path)
        should_fetch = existing.size == 0 or _coverage_years(existing) < float(min_years)
        if should_fetch and api_client is not None:
            try:
                bars = api_client.fetch_daily_aggregates(ticker, days_back=min_years * 365 + 10)
                fetched = _coerce_daily_closes(bars, start_cutoff)
                if fetched.size:
                    existing = fetched
                    _save_cached_closes(cache_path, ticker=ticker, closes=fetched)
                    fetches[ticker] = "refreshed"
                else:
                    fetches[ticker] = "fetch_empty"
            except Exception:
                fetches[ticker] = "fetch_failed"
        elif should_fetch:
            fetches[ticker] = "missing_api_key"
        else:
            fetches[ticker] = "cache_ok"

        if existing.size == 0:
            existing = _synthetic_closes(target_days, seed=len(closes_by_ticker) + 11)
            fetches[ticker] = f"{fetches.get(ticker, 'synthetic')}_synthetic"
        closes_by_ticker[ticker] = existing

    min_len = min((series.size for series in closes_by_ticker.values()), default=0)
    if min_len <= 2:
        return np.zeros((0, 0), dtype=float), {"tickers": tickers, "fetches": fetches, "years_target": min_years}
    matrix = np.column_stack([series[-min_len:] for series in closes_by_ticker.values()])
    returns = np.zeros_like(matrix, dtype=float)
    returns[1:] = matrix[1:] / np.where(matrix[:-1] == 0.0, 1.0, matrix[:-1]) - 1.0
    return returns[1:], {
        "tickers": tickers,
        "fetches": fetches,
        "years_target": min_years,
        "bars": int(max(0, returns.shape[0])),
    }


def _load_cached_closes(path: Path) -> np.ndarray:
    if not path.exists():
        return np.zeros(0, dtype=float)
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return np.zeros(0, dtype=float)
    bars = payload.get("bars", []) if isinstance(payload, dict) else []
    values = [float(row.get("c", 0.0)) for row in bars if isinstance(row, dict) and row.get("c") is not None]
    return np.asarray(values, dtype=float)


def _save_cached_closes(path: Path, *, ticker: str, closes: np.ndarray) -> None:
    bars = [{"idx": int(i), "c": float(value)} for i, value in enumerate(np.asarray(closes, dtype=float))]
    payload = {"ticker": ticker, "cached_at": _utc_now_iso(), "bars": bars}
    path.write_text(json.dumps(payload, indent=2, sort_keys=True), encoding="utf-8")


def _coerce_daily_closes(rows: list[dict[str, Any]], start_cutoff: date) -> np.ndarray:
    closes: list[float] = []
    for row in rows:
        ts = row.get("t")
        close = row.get("c")
        if ts is None or close is None:
            continue
        dt = datetime.fromtimestamp(float(ts) / 1000.0, tz=timezone.utc).date()
        if dt < start_cutoff:
            continue
        closes.append(float(close))
    return np.asarray(closes, dtype=float)


def _coverage_years(series: np.ndarray) -> float:
    return float(np.asarray(series).size) / 252.0


def _synthetic_closes(length: int, *, seed: int) -> np.ndarray:
    rng = np.random.default_rng(seed)
    rets = rng.normal(loc=0.0002, scale=0.012, size=max(3, int(length)))
    return 100.0 * np.cumprod(1.0 + rets)


def _build_market_conditioned_dataset(
    required_features: tuple[str, ...],
    market_returns: np.ndarray,
    *,
    seed: int,
) -> tuple[dict[str, np.ndarray], np.ndarray]:
    arr = np.asarray(market_returns, dtype=float)
    market_mean = np.mean(arr, axis=1)
    market_dispersion = np.std(arr, axis=1)
    sample_size = market_mean.shape[0]
    rng = np.random.default_rng(seed + 1000)
    features: dict[str, np.ndarray] = {}
    for idx, feature_name in enumerate(required_features):
        scale = 0.5 + (idx % 5) * 0.2
        base = market_mean * scale + market_dispersion * (1.0 - min(scale, 0.9))
        noise = rng.normal(loc=0.0, scale=0.03, size=sample_size)
        features[feature_name] = (base + noise).astype(float)
    target = np.roll(market_mean, -1)
    target[-1] = target[-2] if sample_size > 1 else 0.0
    labels = np.where(target >= np.median(target), 1.0, -1.0)
    return features, labels


def _evaluate_market_scenarios(returns: np.ndarray, *, settings: dict[str, Any]) -> dict[str, dict[str, float]]:
    base = np.asarray(returns, dtype=float)
    if base.ndim != 2 or base.size == 0:
        return {}
    baseline = np.mean(base, axis=1)
    out: dict[str, dict[str, float]] = {}
    for name, cfg_raw in settings.items():
        cfg = cfg_raw if isinstance(cfg_raw, dict) else {}
        if not bool(cfg.get("enabled", True)):
            continue
        mult = float(cfg.get("returns_multiplier", 1.0))
        shift = float(cfg.get("returns_shift", 0.0))
        stressed = baseline * mult + shift
        equity = np.cumprod(1.0 + stressed)
        peak = np.maximum.accumulate(equity)
        drawdown = np.where(peak > 0.0, equity / peak - 1.0, 0.0)
        out[name] = {
            "mean_return": float(np.mean(stressed)),
            "volatility": float(np.std(stressed)),
            "max_drawdown": float(np.min(drawdown)),
        }
    return out


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

        errors.extend(
            _validate_optional_spec_object(leg.architecture_spec, field_path=f"legs[{idx}].architecture_spec")
        )
        errors.extend(_validate_optional_spec_object(leg.calibration_spec, field_path=f"legs[{idx}].calibration_spec"))
        errors.extend(
            _validate_optional_spec_object(leg.event_process_spec, field_path=f"legs[{idx}].event_process_spec")
        )

        if _requires_architecture_spec(model_id):
            errors.extend(
                _validate_architecture_spec(leg.architecture_spec, field_path=f"legs[{idx}].architecture_spec")
            )
        if _requires_calibration_spec(model_id):
            errors.extend(
                _validate_calibration_spec(leg.calibration_spec, field_path=f"legs[{idx}].calibration_spec")
            )
        if _requires_event_process_spec(model_id):
            errors.extend(
                _validate_event_process_spec(leg.event_process_spec, field_path=f"legs[{idx}].event_process_spec")
            )
    return errors


def _validate_optional_spec_object(payload: dict[str, Any] | None, *, field_path: str) -> list[str]:
    if payload is None:
        return []
    if isinstance(payload, dict):
        return []
    return [f"{field_path} must be an object when provided"]


def _requires_architecture_spec(model_id: str) -> bool:
    return any(token in model_id for token in ("ann", "neural", "transformer", "mlp"))


def _requires_calibration_spec(model_id: str) -> bool:
    return any(token in model_id for token in ("local_vol", "heston", "sabr", "vol_surface", "black_scholes"))


def _requires_event_process_spec(model_id: str) -> bool:
    return any(token in model_id for token in ("hawkes", "jump", "intensity"))


def _validate_architecture_spec(payload: dict[str, Any] | None, *, field_path: str) -> list[str]:
    if not isinstance(payload, dict):
        return [f"{field_path} is required for ANN/neural model legs"]
    layers = payload.get("layers")
    if not isinstance(layers, list) or not layers:
        return [f"{field_path}.layers must be a non-empty list"]
    return []


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
            artifact_paths={"spec": str(spec_path)},
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
        artifacts = {"spec": str(spec_path), **_flatten_artifact_locations(adapter_output)}
        oos_metrics_payload = {
            leg_name: {
                "model_id": leg_metrics.model_id,
                "metrics": {k: float(v) for k, v in leg_metrics.metrics.items()},
            }
            for leg_name, leg_metrics in adapter_output.per_leg_oos_metrics.items()
        }
        logs = (
            "saved config snapshot",
            f"adapter={adapter_output.adapter_name or runner.__class__.__name__}",
            "computed summary metrics",
        )
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
                "scenario_diagnostics": adapter_output.scenario_diagnostics,
                "data_readiness": adapter_output.data_readiness,
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
            artifact_paths={"spec": str(spec_path)},
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

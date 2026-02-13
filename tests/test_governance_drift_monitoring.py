from backtesting.cache_runner import _build_governance_metadata, _evaluate_governance_gate_checks
from backtesting.monitoring import evaluate_drift_monitoring


def test_evaluate_drift_monitoring_detects_breach() -> None:
    payload = evaluate_drift_monitoring(
        expected={"signal_agreement": 0.95, "fill_slippage_bps": 2.0, "pnl_attribution": 0.9},
        observed={"signal_agreement": 0.70, "fill_slippage_bps": 12.0, "pnl_attribution": 0.4},
        thresholds={"max_signal_agreement_drift": 0.10, "max_fill_slippage_drift_bps": 5.0, "max_pnl_attribution_divergence": 0.15},
    )
    assert payload["within_tolerance"] is False
    assert payload["checks"]["signal_agreement"] is False
    assert payload["checks"]["fill_slippage"] is False
    assert payload["checks"]["pnl_attribution"] is False
    assert len(payload["alert_summaries"]) == 3


def test_evaluate_drift_monitoring_feature_label_residual_and_policy() -> None:
    payload = evaluate_drift_monitoring(
        expected={
            "features": {
                "ret_1d": {"mean": 0.0, "std": 1.0, "psi": 0.01},
                "vol_20d": {"mean": 0.2, "std": 0.1, "psi": 0.03},
            },
            "label": {"error_rate": 0.10, "psi": 0.04, "kld": 0.02},
            "residual": {"mean": 0.0, "std": 1.0, "autocorr_lag1": 0.05},
        },
        observed={
            "features": {
                "ret_1d": {"mean": 0.9, "std": 1.9, "psi": 0.5},
                "vol_20d": {"mean": 0.21, "std": 0.11, "psi": 0.04},
            },
            "label": {"error_rate": 0.20, "psi": 0.30, "kld": 0.20},
            "residual": {"mean": 0.5, "std": 1.5, "autocorr_lag1": 0.4},
            "performance_tracking": {
                "pre_retrain": {"accuracy": 0.55, "rmse": 0.32, "sharpe": 0.6},
                "post_retrain": {"accuracy": 0.72, "rmse": 0.25, "sharpe": 0.9},
                "alerts": {"triggered": 20, "false_alarms": 2},
            },
        },
        thresholds={"max_false_alarm_rate": 0.25},
    )

    assert payload["detectors"]["feature"]["features_evaluated"] == 2
    assert payload["detectors"]["label"]["within_tolerance"] is False
    assert payload["detectors"]["residual"]["within_tolerance"] is False
    assert payload["policy_engine"]["recommended_action"] in {"retrain", "fallback"}
    assert payload["performance_tracking"]["delta"]["accuracy"] > 0.0
    assert payload["performance_tracking"]["alerts"]["within_tolerance"] is True
    assert "route_to" in payload["dashboards"]


def test_governance_metadata_includes_drift_monitoring_and_gate_check() -> None:
    governance = _build_governance_metadata(
        {
            "promotion_state": "paper",
            "dataset_snapshot_lock": "snapshot:v1",
            "approval_status": "pending",
            "expected_outcomes": {"signal_agreement": 0.98, "fill_slippage_bps": 1.0, "pnl_attribution": 0.95},
            "observed_outcomes": {"signal_agreement": 0.97, "fill_slippage_bps": 4.0, "pnl_attribution": 0.91},
        }
    )
    assert "drift_monitoring" in governance
    assert governance["gate_checks"]["drift_monitoring"] is True
    assert "friction_adjusted_edge" in governance["promotion_required_checks"]
    assert "max_feature_mean_shift" in governance["gate_thresholds"]

    checks = _evaluate_governance_gate_checks(
        metrics={"signal_diagnostics_ready": True, "sharpe": 1.2, "rolling_sharpe_mean": 1.0, "turnover_total": 1.0, "friction_adjusted_edge": 1.2},
        fold_rows=[{"fold": 1}, {"fold": 2}, {"fold": 3}],
        governance=governance,
    )
    assert checks["drift_monitoring"] is True
    assert checks["friction_adjusted_edge"] is True



def test_governance_metadata_requires_causal_robustness_for_paper() -> None:
    governance = _build_governance_metadata(
        {
            "promotion_state": "paper",
            "dataset_snapshot_lock": "snapshot:v2",
            "causal_validation": {
                "methods": [
                    {
                        "method": "difference_in_differences",
                        "effect_tstat": 2.6,
                        "pretrend_pvalue": 0.31,
                        "placebo_pvalue": 0.22,
                        "relative_attenuation": 0.21,
                    },
                    {
                        "method": "synthetic_control",
                        "effect_tstat": 2.1,
                        "pretrend_pvalue": 0.28,
                        "placebo_pvalue": 0.17,
                        "relative_attenuation": 0.25,
                    },
                    {
                        "method": "propensity_score_matching",
                        "effect_tstat": 2.2,
                        "pretrend_pvalue": 0.19,
                        "placebo_pvalue": 0.09,
                        "relative_attenuation": 0.30,
                    },
                ]
            },
        }
    )
    assert "causal_robustness" in governance["promotion_required_checks"]
    assert governance["causal_robustness"]["pass"] is True
    checks = _evaluate_governance_gate_checks(
        metrics={"signal_diagnostics_ready": True, "sharpe": 0.8, "rolling_sharpe_mean": 0.7, "turnover_total": 1.2, "friction_adjusted_edge": 0.7},
        fold_rows=[{"fold": 1}, {"fold": 2}, {"fold": 3}],
        governance=governance,
    )
    assert checks["causal_robustness"] is True



def test_governance_causal_robustness_fails_on_missing_method_and_placebo() -> None:
    governance = _build_governance_metadata(
        {
            "promotion_state": "paper",
            "dataset_snapshot_lock": "snapshot:v3",
            "causal_validation": {
                "methods": [
                    {
                        "method": "difference_in_differences",
                        "effect_tstat": 2.4,
                        "pretrend_pvalue": 0.25,
                        "placebo_pvalue": 0.03,
                        "relative_attenuation": 0.20,
                    },
                    {
                        "method": "synthetic_control",
                        "effect_tstat": 2.0,
                        "pretrend_pvalue": 0.14,
                        "placebo_pvalue": 0.08,
                        "relative_attenuation": 0.19,
                    },
                ]
            },
        }
    )
    assert governance["causal_robustness"]["pass"] is False
    assert "propensity_score_matching" in governance["causal_robustness"]["missing_methods"]
    method_rows = governance["causal_robustness"]["method_results"]
    did = [row for row in method_rows if row["method"] == "difference_in_differences"][0]
    assert did["pass"] is False


def test_governance_metadata_requires_experiment_id_for_shadow_and_production() -> None:
    shadow = _build_governance_metadata({"promotion_state": "shadow", "dataset_snapshot_lock": "snap"})
    production = _build_governance_metadata({"promotion_state": "production", "dataset_snapshot_lock": "snap", "experiment_id": "EXP-123"})
    assert "experiment_id" in shadow["promotion_required_checks"]
    assert shadow["gate_checks"]["experiment_id"] is False
    assert production["gate_checks"]["experiment_id"] is True

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

    checks = _evaluate_governance_gate_checks(
        metrics={"signal_diagnostics_ready": True, "sharpe": 1.2, "rolling_sharpe_mean": 1.0, "turnover_total": 1.0},
        fold_rows=[{"fold": 1}, {"fold": 2}, {"fold": 3}],
        governance=governance,
    )
    assert checks["drift_monitoring"] is True

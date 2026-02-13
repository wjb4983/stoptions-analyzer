from __future__ import annotations

from typing import Any


SEVERITY_ORDER = {"low": 0, "medium": 1, "high": 2}


def _to_float(value: Any, default: float) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return float(default)


def _normalize_name(value: Any) -> str:
    text = str(value or "").strip().lower().replace(" ", "_")
    return text or "metric"


def _safe_ratio(numerator: float, denominator: float) -> float:
    if denominator <= 0:
        return 0.0
    return float(numerator) / float(denominator)


def _severity_from_ratio(ratio: float) -> str:
    if ratio >= 1.5:
        return "high"
    if ratio >= 1.0:
        return "medium"
    return "low"


def _pick_higher_severity(left: str, right: str) -> str:
    return left if SEVERITY_ORDER.get(left, 0) >= SEVERITY_ORDER.get(right, 0) else right


def _evaluate_feature_detectors(
    *,
    expected: dict[str, Any],
    observed: dict[str, Any],
    thresholds: dict[str, Any],
) -> dict[str, Any]:
    expected_features = expected.get("features") if isinstance(expected.get("features"), dict) else {}
    observed_features = observed.get("features") if isinstance(observed.get("features"), dict) else {}
    feature_threshold_overrides = thresholds.get("feature_thresholds") if isinstance(thresholds.get("feature_thresholds"), dict) else {}

    default_mean_shift = _to_float(thresholds.get("max_feature_mean_shift", 0.15), 0.15)
    default_std_ratio_shift = _to_float(thresholds.get("max_feature_std_ratio_shift", 0.35), 0.35)
    default_psi = _to_float(thresholds.get("max_feature_psi", 0.20), 0.20)

    all_features = sorted(set(expected_features.keys()) | set(observed_features.keys()))
    rows: list[dict[str, Any]] = []
    overall_severity = "low"

    for name in all_features:
        e_row = expected_features.get(name) if isinstance(expected_features.get(name), dict) else {}
        o_row = observed_features.get(name) if isinstance(observed_features.get(name), dict) else {}
        threshold_row = feature_threshold_overrides.get(name) if isinstance(feature_threshold_overrides.get(name), dict) else {}

        e_mean = _to_float(e_row.get("mean", 0.0), 0.0)
        o_mean = _to_float(o_row.get("mean", e_mean), e_mean)
        e_std = max(1e-9, abs(_to_float(e_row.get("std", 1.0), 1.0)))
        o_std = max(1e-9, abs(_to_float(o_row.get("std", e_std), e_std)))

        mean_shift = abs(o_mean - e_mean) / e_std
        std_ratio_shift = abs((o_std / e_std) - 1.0)
        psi = _to_float(o_row.get("psi", e_row.get("psi", 0.0)), 0.0)

        max_mean_shift = _to_float(threshold_row.get("max_mean_shift", default_mean_shift), default_mean_shift)
        max_std_ratio_shift = _to_float(threshold_row.get("max_std_ratio_shift", default_std_ratio_shift), default_std_ratio_shift)
        max_psi = _to_float(threshold_row.get("max_psi", default_psi), default_psi)

        mean_ratio = _safe_ratio(mean_shift, max(max_mean_shift, 1e-9))
        std_ratio = _safe_ratio(std_ratio_shift, max(max_std_ratio_shift, 1e-9))
        psi_ratio = _safe_ratio(psi, max(max_psi, 1e-9))
        breach_score = max(mean_ratio, std_ratio, psi_ratio)
        severity = _severity_from_ratio(breach_score)
        overall_severity = _pick_higher_severity(overall_severity, severity)

        row = {
            "feature": str(name),
            "mean_shift_z": mean_shift,
            "std_ratio_shift": std_ratio_shift,
            "psi": psi,
            "thresholds": {
                "max_mean_shift": max_mean_shift,
                "max_std_ratio_shift": max_std_ratio_shift,
                "max_psi": max_psi,
            },
            "breach_score": breach_score,
            "severity": severity,
            "within_tolerance": severity == "low",
        }
        rows.append(row)

    return {
        "rows": rows,
        "features_evaluated": len(rows),
        "severity": overall_severity,
        "within_tolerance": all(bool(row.get("within_tolerance", False)) for row in rows),
    }


def _evaluate_label_and_residual_detectors(*, expected: dict[str, Any], observed: dict[str, Any], thresholds: dict[str, Any]) -> dict[str, Any]:
    label_expected = expected.get("label") if isinstance(expected.get("label"), dict) else {}
    label_observed = observed.get("label") if isinstance(observed.get("label"), dict) else {}
    residual_expected = expected.get("residual") if isinstance(expected.get("residual"), dict) else {}
    residual_observed = observed.get("residual") if isinstance(observed.get("residual"), dict) else {}

    label_psi = abs(_to_float(label_observed.get("psi", label_expected.get("psi", 0.0)), 0.0))
    label_kld = abs(_to_float(label_observed.get("kld", label_expected.get("kld", 0.0)), 0.0))
    label_error_rate_drift = abs(
        _to_float(label_observed.get("error_rate", label_expected.get("error_rate", 0.0)), 0.0)
        - _to_float(label_expected.get("error_rate", 0.0), 0.0)
    )

    residual_mean_shift = abs(
        _to_float(residual_observed.get("mean", residual_expected.get("mean", 0.0)), 0.0)
        - _to_float(residual_expected.get("mean", 0.0), 0.0)
    )
    residual_std_ratio = abs(
        _to_float(residual_observed.get("std", residual_expected.get("std", 1.0)), 1.0)
        / max(1e-9, abs(_to_float(residual_expected.get("std", 1.0), 1.0)))
        - 1.0
    )
    residual_autocorr = abs(_to_float(residual_observed.get("autocorr_lag1", residual_expected.get("autocorr_lag1", 0.0)), 0.0))

    max_label_psi = _to_float(thresholds.get("max_label_psi", 0.20), 0.20)
    max_label_kld = _to_float(thresholds.get("max_label_kld", 0.12), 0.12)
    max_label_error_rate_drift = _to_float(thresholds.get("max_label_error_rate_drift", 0.03), 0.03)
    max_residual_mean_shift = _to_float(thresholds.get("max_residual_mean_shift", 0.10), 0.10)
    max_residual_std_ratio_shift = _to_float(thresholds.get("max_residual_std_ratio_shift", 0.25), 0.25)
    max_residual_autocorr = _to_float(thresholds.get("max_residual_autocorr", 0.20), 0.20)

    label_score = max(
        _safe_ratio(label_psi, max(max_label_psi, 1e-9)),
        _safe_ratio(label_kld, max(max_label_kld, 1e-9)),
        _safe_ratio(label_error_rate_drift, max(max_label_error_rate_drift, 1e-9)),
    )
    residual_score = max(
        _safe_ratio(residual_mean_shift, max(max_residual_mean_shift, 1e-9)),
        _safe_ratio(residual_std_ratio, max(max_residual_std_ratio_shift, 1e-9)),
        _safe_ratio(residual_autocorr, max(max_residual_autocorr, 1e-9)),
    )

    label_severity = _severity_from_ratio(label_score)
    residual_severity = _severity_from_ratio(residual_score)

    return {
        "label": {
            "psi": label_psi,
            "kld": label_kld,
            "error_rate_drift": label_error_rate_drift,
            "breach_score": label_score,
            "severity": label_severity,
            "within_tolerance": label_severity == "low",
            "thresholds": {
                "max_label_psi": max_label_psi,
                "max_label_kld": max_label_kld,
                "max_label_error_rate_drift": max_label_error_rate_drift,
            },
        },
        "residual": {
            "mean_shift": residual_mean_shift,
            "std_ratio_shift": residual_std_ratio,
            "autocorr_lag1": residual_autocorr,
            "breach_score": residual_score,
            "severity": residual_severity,
            "within_tolerance": residual_severity == "low",
            "thresholds": {
                "max_residual_mean_shift": max_residual_mean_shift,
                "max_residual_std_ratio_shift": max_residual_std_ratio_shift,
                "max_residual_autocorr": max_residual_autocorr,
            },
        },
    }


def _build_policy_engine(*, severity: str, detectors: dict[str, Any], thresholds: dict[str, Any]) -> dict[str, Any]:
    retrain_on = {_normalize_name(name) for name in (thresholds.get("retrain_on") or ["high"])}
    recalibrate_on = {_normalize_name(name) for name in (thresholds.get("recalibrate_on") or ["medium"])}

    if _normalize_name(severity) in retrain_on:
        action = "retrain"
    elif _normalize_name(severity) in recalibrate_on:
        action = "recalibrate"
    else:
        action = "fallback" if severity == "high" else "continue"

    if action == "continue" and (not detectors.get("label", {}).get("within_tolerance", True) or not detectors.get("residual", {}).get("within_tolerance", True)):
        action = "recalibrate"

    return {
        "severity": severity,
        "recommended_action": action,
        "playbook": {
            "retrain": ["freeze promotion", "launch retraining pipeline", "re-validate before deploy"],
            "recalibrate": ["recalibrate thresholds/probabilities", "monitor residual drift hourly"],
            "fallback": ["route to fallback model", "reduce exposure limits", "escalate to on-call"],
            "continue": ["no intervention required"],
        },
    }


def _build_performance_tracking(*, expected: dict[str, Any], observed: dict[str, Any], thresholds: dict[str, Any]) -> dict[str, Any]:
    perf = observed.get("performance_tracking") if isinstance(observed.get("performance_tracking"), dict) else {}
    pre = perf.get("pre_retrain") if isinstance(perf.get("pre_retrain"), dict) else {}
    post = perf.get("post_retrain") if isinstance(perf.get("post_retrain"), dict) else {}
    alerts = perf.get("alerts") if isinstance(perf.get("alerts"), dict) else {}

    pre_acc = _to_float(pre.get("accuracy", expected.get("accuracy", 0.0)), 0.0)
    post_acc = _to_float(post.get("accuracy", pre_acc), pre_acc)
    pre_rmse = _to_float(pre.get("rmse", expected.get("rmse", 0.0)), 0.0)
    post_rmse = _to_float(post.get("rmse", pre_rmse), pre_rmse)

    pre_sharpe = _to_float(pre.get("sharpe", expected.get("sharpe", 0.0)), 0.0)
    post_sharpe = _to_float(post.get("sharpe", pre_sharpe), pre_sharpe)

    triggered_alerts = _to_float(alerts.get("triggered", 0.0), 0.0)
    false_alarms = _to_float(alerts.get("false_alarms", 0.0), 0.0)

    max_false_alarm_rate = _to_float(thresholds.get("max_false_alarm_rate", 0.25), 0.25)
    false_alarm_rate = _safe_ratio(false_alarms, max(1.0, triggered_alerts))

    return {
        "pre_retrain": {
            "accuracy": pre_acc,
            "rmse": pre_rmse,
            "sharpe": pre_sharpe,
        },
        "post_retrain": {
            "accuracy": post_acc,
            "rmse": post_rmse,
            "sharpe": post_sharpe,
        },
        "delta": {
            "accuracy": post_acc - pre_acc,
            "rmse": post_rmse - pre_rmse,
            "sharpe": post_sharpe - pre_sharpe,
        },
        "alerts": {
            "triggered": triggered_alerts,
            "false_alarms": false_alarms,
            "false_alarm_rate": false_alarm_rate,
            "within_tolerance": false_alarm_rate <= max_false_alarm_rate,
            "max_false_alarm_rate": max_false_alarm_rate,
        },
    }


def _build_drift_dashboards(*, alerts: list[dict[str, str]], policy: dict[str, Any], thresholds: dict[str, Any]) -> dict[str, Any]:
    routing = thresholds.get("alert_routing") if isinstance(thresholds.get("alert_routing"), dict) else {}
    channels = {
        "low": routing.get("low", ["dashboard"]),
        "medium": routing.get("medium", ["dashboard", "slack:#ml-monitoring"]),
        "high": routing.get("high", ["dashboard", "slack:#incident-ml", "pagerduty:ml-oncall"]),
    }
    severity = str(policy.get("severity", "low"))
    return {
        "widgets": [
            "feature_drift_heatmap",
            "label_concept_drift_timeseries",
            "residual_control_chart",
            "pre_post_retrain_performance",
            "false_alarm_rate_gauge",
        ],
        "alert_count": len(alerts),
        "route_to": channels.get(severity, channels.get("low", ["dashboard"])),
        "channels": channels,
    }


def evaluate_drift_monitoring(
    *,
    expected: dict[str, Any] | None,
    observed: dict[str, Any] | None,
    thresholds: dict[str, Any] | None,
) -> dict[str, Any]:
    expected = expected if isinstance(expected, dict) else {}
    observed = observed if isinstance(observed, dict) else {}
    thresholds = thresholds if isinstance(thresholds, dict) else {}

    expected_signal_agreement = _to_float(expected.get("signal_agreement", 1.0), 1.0)
    observed_signal_agreement = _to_float(observed.get("signal_agreement", expected_signal_agreement), expected_signal_agreement)
    signal_agreement_drift = abs(observed_signal_agreement - expected_signal_agreement)

    expected_fill_slippage_bps = _to_float(expected.get("fill_slippage_bps", 0.0), 0.0)
    observed_fill_slippage_bps = _to_float(observed.get("fill_slippage_bps", expected_fill_slippage_bps), expected_fill_slippage_bps)
    fill_slippage_drift_bps = observed_fill_slippage_bps - expected_fill_slippage_bps

    expected_pnl_attribution = _to_float(expected.get("pnl_attribution", 1.0), 1.0)
    observed_pnl_attribution = _to_float(observed.get("pnl_attribution", expected_pnl_attribution), expected_pnl_attribution)
    pnl_attribution_divergence = abs(observed_pnl_attribution - expected_pnl_attribution)

    max_signal_agreement_drift = _to_float(thresholds.get("max_signal_agreement_drift", 0.10), 0.10)
    max_fill_slippage_drift_bps = _to_float(thresholds.get("max_fill_slippage_drift_bps", 5.0), 5.0)
    max_pnl_attribution_divergence = _to_float(thresholds.get("max_pnl_attribution_divergence", 0.15), 0.15)

    checks = {
        "signal_agreement": signal_agreement_drift <= max_signal_agreement_drift,
        "fill_slippage": fill_slippage_drift_bps <= max_fill_slippage_drift_bps,
        "pnl_attribution": pnl_attribution_divergence <= max_pnl_attribution_divergence,
    }

    alerts: list[dict[str, str]] = []
    if not checks["signal_agreement"]:
        alerts.append(
            {
                "metric": "signal_agreement",
                "severity": "high",
                "summary": f"Signal agreement drift {signal_agreement_drift:.3f} exceeds tolerance {max_signal_agreement_drift:.3f}.",
            }
        )
    if not checks["fill_slippage"]:
        alerts.append(
            {
                "metric": "fill_slippage",
                "severity": "high",
                "summary": f"Fill slippage drift {fill_slippage_drift_bps:.2f}bps exceeds tolerance {max_fill_slippage_drift_bps:.2f}bps.",
            }
        )
    if not checks["pnl_attribution"]:
        alerts.append(
            {
                "metric": "pnl_attribution",
                "severity": "high",
                "summary": f"PnL attribution divergence {pnl_attribution_divergence:.3f} exceeds tolerance {max_pnl_attribution_divergence:.3f}.",
            }
        )

    feature_detectors = _evaluate_feature_detectors(expected=expected, observed=observed, thresholds=thresholds)
    label_residual_detectors = _evaluate_label_and_residual_detectors(expected=expected, observed=observed, thresholds=thresholds)

    for row in feature_detectors.get("rows", []):
        if isinstance(row, dict) and not bool(row.get("within_tolerance", True)):
            alerts.append(
                {
                    "metric": f"feature:{row.get('feature', 'unknown')}",
                    "severity": str(row.get("severity", "medium")),
                    "summary": f"Feature drift {row.get('feature', 'unknown')} breach score {float(row.get('breach_score', 0.0)):.2f}.",
                }
            )

    for name in ("label", "residual"):
        row = label_residual_detectors.get(name, {}) if isinstance(label_residual_detectors.get(name), dict) else {}
        if row and not bool(row.get("within_tolerance", True)):
            alerts.append(
                {
                    "metric": name,
                    "severity": str(row.get("severity", "medium")),
                    "summary": f"{name.title()} drift breach score {float(row.get('breach_score', 0.0)):.2f}.",
                }
            )

    overall_severity = "low"
    if alerts:
        overall_severity = max((str(item.get("severity", "low")) for item in alerts), key=lambda value: SEVERITY_ORDER.get(value, 0))

    policy_engine = _build_policy_engine(
        severity=overall_severity,
        detectors={"feature": feature_detectors, **label_residual_detectors},
        thresholds=thresholds,
    )
    performance_tracking = _build_performance_tracking(expected=expected, observed=observed, thresholds=thresholds)
    dashboards = _build_drift_dashboards(alerts=alerts, policy=policy_engine, thresholds=thresholds)

    within_tolerance = all(checks.values()) and bool(feature_detectors.get("within_tolerance", True))
    within_tolerance = within_tolerance and bool(label_residual_detectors.get("label", {}).get("within_tolerance", True))
    within_tolerance = within_tolerance and bool(label_residual_detectors.get("residual", {}).get("within_tolerance", True))
    within_tolerance = within_tolerance and bool(performance_tracking.get("alerts", {}).get("within_tolerance", True))

    return {
        "expected": {
            "signal_agreement": expected_signal_agreement,
            "fill_slippage_bps": expected_fill_slippage_bps,
            "pnl_attribution": expected_pnl_attribution,
        },
        "observed": {
            "signal_agreement": observed_signal_agreement,
            "fill_slippage_bps": observed_fill_slippage_bps,
            "pnl_attribution": observed_pnl_attribution,
        },
        "thresholds": {
            "max_signal_agreement_drift": max_signal_agreement_drift,
            "max_fill_slippage_drift_bps": max_fill_slippage_drift_bps,
            "max_pnl_attribution_divergence": max_pnl_attribution_divergence,
        },
        "drift": {
            "signal_agreement": signal_agreement_drift,
            "fill_slippage_bps": fill_slippage_drift_bps,
            "pnl_attribution": pnl_attribution_divergence,
        },
        "checks": checks,
        "detectors": {
            "feature": feature_detectors,
            **label_residual_detectors,
        },
        "policy_engine": policy_engine,
        "performance_tracking": performance_tracking,
        "dashboards": dashboards,
        "alert_summaries": alerts,
        "within_tolerance": within_tolerance,
    }

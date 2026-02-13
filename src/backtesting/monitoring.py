from __future__ import annotations

from typing import Any


def _to_float(value: Any, default: float) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return float(default)


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
        "alert_summaries": alerts,
        "within_tolerance": all(checks.values()),
    }

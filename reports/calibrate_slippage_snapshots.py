from __future__ import annotations

import argparse
import json
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import numpy as np


def _month_key(ts: str) -> str:
    if not ts:
        return "1970-01"
    return str(ts)[:7]


def _bucket_id(row: dict[str, object]) -> str:
    p = float(row.get("participation_rate", 0.0))
    v = float(row.get("volatility", 0.0))
    s = float(row.get("spread_bps", 0.0))
    r = str(row.get("regime", "unknown"))
    p_bin = "p_low" if p < 0.1 else ("p_mid" if p < 0.3 else "p_high")
    v_bin = "v_low" if v < 0.01 else ("v_mid" if v < 0.03 else "v_high")
    s_bin = "s_tight" if s < 5.0 else ("s_mid" if s < 15.0 else "s_wide")
    return f"{p_bin}|{v_bin}|{s_bin}|{r}"


def _load_rows(path: Path) -> list[dict[str, object]]:
    payload = json.loads(path.read_text())
    if isinstance(payload, list):
        return [row for row in payload if isinstance(row, dict)]
    if isinstance(payload, dict) and isinstance(payload.get("fills"), list):
        return [row for row in payload["fills"] if isinstance(row, dict)]
    return []


def calibrate(fills_path: Path, out_snapshots: Path, out_report: Path) -> None:
    rows = _load_rows(fills_path)
    grouped: dict[str, list[dict[str, object]]] = defaultdict(list)
    for row in rows:
        grouped[_month_key(str(row.get("event_timestamp") or row.get("timestamp") or ""))].append(row)

    snapshots: list[dict[str, object]] = []
    fit_errors: list[float] = []
    coeffs: list[float] = []

    for month in sorted(grouped.keys()):
        month_rows = grouped[month]
        parts = np.asarray([max(float(r.get("participation_rate", 0.0)), 0.0) for r in month_rows], dtype=float)
        slips = np.asarray([max(float(r.get("slippage_bps", 0.0)), 0.0) for r in month_rows], dtype=float)
        if parts.size == 0:
            continue
        denom = float(np.sum(parts * parts))
        coeff = float(np.sum(parts * slips) / denom) if denom > 0 else 0.0
        pred = parts * coeff
        mae = float(np.mean(np.abs(slips - pred))) if slips.size else 0.0
        fit_errors.append(mae)
        coeffs.append(coeff)

        bucket_metrics: dict[str, dict[str, float]] = {}
        by_bucket: dict[str, list[float]] = defaultdict(list)
        for row in month_rows:
            by_bucket[_bucket_id(row)].append(float(row.get("slippage_bps", 0.0)))
        for key, vals in by_bucket.items():
            arr = np.asarray(vals, dtype=float)
            bucket_metrics[key] = {
                "count": int(arr.size),
                "mean_slippage_bps": float(np.mean(arr)) if arr.size else 0.0,
            }

        snapshots.append(
            {
                "effective_date": f"{month}-01",
                "stable": bool(len(coeffs) < 2 or abs(coeffs[-1] - coeffs[-2]) <= 10.0),
                "params": {
                    "base_bps": 0.0,
                    "impact_coefficient_bps": coeff,
                    "participation_exponent": 1.0,
                    "max_participation": 1.0,
                },
                "fit_error_mae_bps": mae,
                "bucket_metrics": bucket_metrics,
            }
        )

    payload = {
        "generated_at": datetime.utcnow().isoformat() + "Z",
        "default_params": {
            "base_bps": 0.0,
            "impact_coefficient_bps": 20.0,
            "participation_exponent": 1.0,
            "max_participation": 1.0,
        },
        "snapshots": snapshots,
    }
    out_snapshots.write_text(json.dumps(payload, indent=2))

    stability = float(np.std(np.asarray(coeffs, dtype=float))) if coeffs else 0.0
    report = {
        "generated_at": payload["generated_at"],
        "fills_path": str(fills_path),
        "snapshot_count": len(snapshots),
        "fit_error": {
            "mae_bps_mean": float(np.mean(np.asarray(fit_errors, dtype=float))) if fit_errors else 0.0,
            "mae_bps_max": float(np.max(np.asarray(fit_errors, dtype=float))) if fit_errors else 0.0,
        },
        "stability": {
            "impact_coefficient_std_bps": stability,
            "impact_coefficient_range_bps": float(max(coeffs) - min(coeffs)) if coeffs else 0.0,
        },
    }
    out_report.write_text(json.dumps(report, indent=2))


def main() -> None:
    parser = argparse.ArgumentParser(description="Calibrate slippage snapshots from historical fills")
    parser.add_argument("--fills", default="reports/historical_fills.json")
    parser.add_argument("--snapshots-out", default="reports/slippage_calibration_snapshots.json")
    parser.add_argument("--report-out", default="reports/calibration_report.json")
    args = parser.parse_args()
    calibrate(Path(args.fills), Path(args.snapshots_out), Path(args.report_out))


if __name__ == "__main__":
    main()

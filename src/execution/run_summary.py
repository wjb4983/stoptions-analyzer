from __future__ import annotations

import csv
import json
from pathlib import Path
from typing import Any


SUMMARY_DIR_NAME = "summary"


def _read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


def _write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, sort_keys=True), encoding="utf-8")


def _write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow({name: row.get(name, "") for name in fieldnames})


def _load_metrics(run_dir: Path) -> dict[str, float]:
    metrics_path = run_dir / "metrics.json"
    parsed = _read_json(metrics_path)
    if isinstance(parsed, list):
        out: dict[str, float] = {}
        for row in parsed:
            if not isinstance(row, dict):
                continue
            metric = str(row.get("metric", "")).strip()
            value = row.get("value")
            if metric and isinstance(value, (int, float)):
                out[metric] = float(value)
        if out:
            return out
    aggregate = _read_json(run_dir / "aggregate_metrics.json")
    if isinstance(aggregate, dict):
        return {str(key): float(value) for key, value in aggregate.items() if isinstance(value, (int, float))}
    return {}


def _build_leaderboard_summary(run_dir: Path) -> dict[str, Any]:
    leaderboard = _read_json(run_dir / "leaderboard.json")
    rows = [row for row in leaderboard if isinstance(row, dict)] if isinstance(leaderboard, list) else []
    metrics = _load_metrics(run_dir)
    top = rows[0] if rows else {}
    payload = {
        "row_count": len(rows),
        "top_row": top,
        "key_metrics": {key: metrics.get(key) for key in ("sharpe", "cagr", "max_drawdown", "total_return") if key in metrics},
    }
    return payload


def _build_risk_summary(run_dir: Path) -> dict[str, Any]:
    risk_rows = _read_json(run_dir / "risk_diagnostics.json")
    rows = [row for row in risk_rows if isinstance(row, dict)] if isinstance(risk_rows, list) else []
    metrics = _load_metrics(run_dir)
    return {
        "row_count": len(rows),
        "key_metrics": {
            key: metrics.get(key)
            for key in (
                "volatility",
                "max_drawdown",
                "turnover_total",
                "var_95",
                "cost_total",
            )
            if key in metrics
        },
    }


def _build_governance_summary(run_dir: Path) -> dict[str, Any]:
    manifest = _read_json(run_dir / "manifest.json")
    governance = manifest.get("governance_metadata") if isinstance(manifest, dict) and isinstance(manifest.get("governance_metadata"), dict) else {}
    checks = governance.get("gate_checks", {}) if isinstance(governance.get("gate_checks"), dict) else {}
    failed_checks = sorted([name for name, passed in checks.items() if not bool(passed)])
    return {
        "promotion_state": governance.get("promotion_state"),
        "approval_status": governance.get("approval_status"),
        "is_promotion_ready": bool(governance.get("is_promotion_ready", False)),
        "failed_gate_checks": failed_checks,
        "missing_required_checks": list(governance.get("missing_required_checks", [])) if isinstance(governance.get("missing_required_checks"), list) else [],
    }


def _build_top_diagnostics_summary(run_dir: Path) -> dict[str, Any]:
    explain = _read_json(run_dir / "trade_explainability.json")
    explain_rows = [row for row in explain if isinstance(row, dict)] if isinstance(explain, list) else []
    flagged = 0
    by_regime: dict[str, int] = {}
    for row in explain_rows:
        red_flags = row.get("red_flags")
        if isinstance(red_flags, list) and red_flags:
            flagged += 1
            regime = str(row.get("regime", row.get("regime_label", "unknown"))).strip() or "unknown"
            by_regime[regime] = by_regime.get(regime, 0) + 1
    ranked_regimes = sorted(by_regime.items(), key=lambda item: item[1], reverse=True)[:5]
    return {
        "trade_count": len(explain_rows),
        "flagged_trade_count": flagged,
        "flagged_by_regime": [{"regime": regime, "count": count} for regime, count in ranked_regimes],
    }


def build_workflow_summary(*, run_dir: str | Path, workflow: str) -> list[str]:
    root = Path(run_dir).expanduser().resolve()
    if not root.exists() or not root.is_dir():
        return []

    summary_dir = root / SUMMARY_DIR_NAME
    created: list[str] = []

    leaderboard = _build_leaderboard_summary(root)
    risk = _build_risk_summary(root)
    governance = _build_governance_summary(root)
    diagnostics = _build_top_diagnostics_summary(root)

    leaderboard_path = summary_dir / "leaderboard_stats.json"
    risk_path = summary_dir / "risk_metrics.json"
    governance_path = summary_dir / "governance_checks.json"
    diagnostics_path = summary_dir / "top_diagnostics.json"

    _write_json(leaderboard_path, {"workflow": workflow, **leaderboard})
    _write_json(risk_path, {"workflow": workflow, **risk})
    _write_json(governance_path, {"workflow": workflow, **governance})
    _write_json(diagnostics_path, {"workflow": workflow, **diagnostics})
    created.extend([
        leaderboard_path.relative_to(root).as_posix(),
        risk_path.relative_to(root).as_posix(),
        governance_path.relative_to(root).as_posix(),
        diagnostics_path.relative_to(root).as_posix(),
    ])

    csv_rows = [
        {"summary_type": "leaderboard_stats", "row_count": leaderboard.get("row_count", 0), "key": "top_row", "value": json.dumps(leaderboard.get("top_row", {}), sort_keys=True)},
        {"summary_type": "risk_metrics", "row_count": risk.get("row_count", 0), "key": "key_metrics", "value": json.dumps(risk.get("key_metrics", {}), sort_keys=True)},
        {"summary_type": "governance_checks", "row_count": 0, "key": "failed_gate_checks", "value": json.dumps(governance.get("failed_gate_checks", []), sort_keys=True)},
        {"summary_type": "top_diagnostics", "row_count": diagnostics.get("trade_count", 0), "key": "flagged_trade_count", "value": str(diagnostics.get("flagged_trade_count", 0))},
    ]
    csv_path = summary_dir / "summary_metrics.csv"
    _write_csv(csv_path, csv_rows, ["summary_type", "row_count", "key", "value"])
    created.append(csv_path.relative_to(root).as_posix())
    return created

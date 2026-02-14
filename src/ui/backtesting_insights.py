from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
import statistics
from typing import Any

from backtesting.experiment_registry import append_governance_event, read_registry

SUPPORTED_MANIFEST_SCHEMA_RANGE = ("1.0", "2.0")
SUPPORTED_METRIC_SCHEMA_RANGE = ("1.0", "1.0")


def _parse_version(version: str) -> tuple[int, ...]:
    parts: list[int] = []
    for token in str(version).split('.'):
        try:
            parts.append(int(token))
        except ValueError:
            parts.append(0)
    return tuple(parts)


def _version_in_range(version: str, minimum: str, maximum: str) -> bool:
    parsed = _parse_version(version)
    return _parse_version(minimum) <= parsed <= _parse_version(maximum)


def _read_metric_manifest(run_dir: Path) -> dict[str, Any]:
    path = run_dir / "metric_tables_manifest.json"
    if not path.exists():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return payload if isinstance(payload, dict) else {}


def _validate_metric_manifest(payload: dict[str, Any]) -> list[str]:
    warnings: list[str] = []
    tables = payload.get("tables", []) if isinstance(payload, dict) else []
    for table in tables if isinstance(tables, list) else []:
        if not isinstance(table, dict):
            continue
        schema = str(table.get("schema_version", ""))
        compat = table.get("compatibility", {}) if isinstance(table.get("compatibility"), dict) else {}
        min_schema = str(compat.get("minimum_reader_schema", SUPPORTED_METRIC_SCHEMA_RANGE[0]))
        max_schema = str(compat.get("maximum_reader_schema", SUPPORTED_METRIC_SCHEMA_RANGE[1]))
        if not _version_in_range(SUPPORTED_METRIC_SCHEMA_RANGE[0], min_schema, max_schema):
            warnings.append(f"reader schema unsupported for table {table.get('table')}: requires [{min_schema}, {max_schema}]")
        if schema and not _version_in_range(schema, min_schema, max_schema):
            warnings.append(f"table schema out of compatibility range for {table.get('table')}: {schema} not in [{min_schema}, {max_schema}]")
    return warnings


def read_experiment_index(output_dir: Path) -> list[dict[str, Any]]:
    return list(reversed(read_registry(output_dir)))


def apply_governance_decision(
    output_dir: Path,
    *,
    run_id: str,
    action: str,
    reason: str,
    actor: str = "ui",
) -> bool:
    action_l = action.strip().lower()
    if action_l not in {"promote", "reject", "waive"}:
        raise ValueError(f"Unsupported governance action: {action}")
    if not reason.strip():
        raise ValueError("A reason is required for governance actions")

    rows = read_registry(output_dir)
    target: dict[str, Any] | None = None
    for row in rows:
        if str(row.get("run_id", "")) == run_id:
            target = row
            break
    if target is None:
        return False

    governance = target.get("governance", {}) if isinstance(target.get("governance"), dict) else {}
    if action_l == "promote":
        drift_monitoring = governance.get("drift_monitoring", {}) if isinstance(governance.get("drift_monitoring"), dict) else {}
        if not bool(drift_monitoring.get("within_tolerance", False)):
            return False
        if not str(governance.get("experiment_id", "")).strip():
            return False
        current = str(governance.get("promotion_state", "research")).strip() or "research"
        order = ["research", "paper", "shadow", "production"]
        if current in order and current != order[-1]:
            governance["promotion_state"] = order[order.index(current) + 1]
        governance["approval_status"] = "approved"
    elif action_l == "reject":
        governance["approval_status"] = "rejected"
    else:
        governance["approval_status"] = "waived"

    event = {
        "timestamp": datetime.now().isoformat(),
        "event": f"workflow_{action_l}",
        "reason": reason.strip(),
        "actor": actor,
    }
    trail = governance.get("audit_trail") if isinstance(governance.get("audit_trail"), list) else []
    trail.append(event)
    governance["audit_trail"] = trail
    target["governance"] = governance

    append_governance_event(
        output_dir,
        run_id=run_id,
        action=action_l,
        reason=reason.strip(),
        actor=actor,
        resulting_promotion_state=str(governance.get("promotion_state", "")),
        resulting_approval_status=str(governance.get("approval_status", "")),
    )

    from backtesting.experiment_registry import append_experiment_entry

    append_experiment_entry(output_dir, target)
    return True


def parse_tags(manifest: dict[str, Any] | None) -> list[str]:
    if not isinstance(manifest, dict):
        return []
    tags: list[str] = []
    note_text = str(manifest.get("notes", ""))
    for token in note_text.replace("\n", " ").split(" "):
        token = token.strip()
        if token.startswith("#") and len(token) > 1:
            tags.append(token[1:])
    param_tags = manifest.get("tags")
    if isinstance(param_tags, list):
        for value in param_tags:
            if isinstance(value, str) and value.strip():
                tags.append(value.strip())
    unique: list[str] = []
    seen: set[str] = set()
    for tag in tags:
        key = tag.lower()
        if key in seen:
            continue
        seen.add(key)
        unique.append(tag)
    return unique


def metric_deltas(base: dict[str, float], other: dict[str, float]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for metric in sorted(set(base.keys()) & set(other.keys())):
        base_value = float(base[metric])
        other_value = float(other[metric])
        delta = other_value - base_value
        pct = (delta / abs(base_value)) if abs(base_value) > 1e-12 else None
        rows.append(
            {
                "metric": metric,
                "base": base_value,
                "compare": other_value,
                "delta": delta,
                "delta_pct": pct,
            }
        )
    rows.sort(key=lambda row: abs(float(row["delta"])), reverse=True)
    return rows


def parameter_diffs(base: dict[str, Any], other: dict[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    keys = sorted(set(base.keys()) | set(other.keys()))
    for key in keys:
        a = base.get(key)
        b = other.get(key)
        if a == b:
            continue
        rows.append({"parameter": key, "base": a, "compare": b})
    return rows


def build_guardrails(
    metrics: dict[str, float],
    *,
    fold_rows: list[dict[str, Any]] | None = None,
    trade_count: int | None = None,
    robustness: dict[str, Any] | None = None,
    evidence_links: dict[str, str] | None = None,
) -> list[dict[str, str]]:
    badges: list[dict[str, str]] = []
    sharpe = float(metrics.get("sharpe", 0.0)) if "sharpe" in metrics else None
    d_sharpe = float(metrics.get("deflated_sharpe_ratio", 0.0)) if "deflated_sharpe_ratio" in metrics else None
    turnover = float(metrics.get("turnover_total", 0.0)) if "turnover_total" in metrics else None

    if sharpe is not None and d_sharpe is not None and sharpe > 0.0 and d_sharpe < 0.35:
        badges.append({"label": "Overfit Risk", "severity": "high", "reason": "High Sharpe with low deflated Sharpe."})

    if trade_count is not None and trade_count < 40:
        badges.append({"label": "Low Sample", "severity": "high", "reason": f"Only {trade_count} trades."})

    if turnover is not None and turnover > 4.0:
        badges.append({"label": "High Turnover", "severity": "medium", "reason": f"Turnover total {turnover:.2f}."})

    if fold_rows:
        unique = len({json.dumps(row.get("selected_params", {}), sort_keys=True) for row in fold_rows})
        if unique > max(3, int(len(fold_rows) * 0.6)):
            badges.append({"label": "Unstable Params", "severity": "medium", "reason": "Many unique selected parameters across folds."})

    if robustness and isinstance(robustness, dict):
        white = robustness.get("white_reality_check", {})
        spa = robustness.get("spa", {})
        white_p = float(white.get("p_value", 1.0)) if isinstance(white, dict) else 1.0
        spa_p = float(spa.get("p_value", 1.0)) if isinstance(spa, dict) else 1.0
        if white_p > 0.1:
            badges.append({"label": "Weak RC", "severity": "medium", "reason": f"White RC p-value {white_p:.3f}."})
        if spa_p > 0.1:
            badges.append({"label": "Weak SPA", "severity": "medium", "reason": f"SPA p-value {spa_p:.3f}."})

        mt = robustness.get("multiple_testing", {})
        if isinstance(mt, dict):
            min_nominal = float(mt.get("min_raw_pvalue", 1.0))
            min_adjusted = float(mt.get("min_bh_adjusted_pvalue", 1.0))
            if min_nominal <= 0.05 and min_adjusted > 0.05:
                badges.append({
                    "label": "Alpha Not Robust",
                    "severity": "high",
                    "reason": f"Nominal p={min_nominal:.3f} but BH-adjusted p={min_adjusted:.3f}.",
                })

    if not badges:
        badges.append({"label": "Guardrails OK", "severity": "low", "reason": "No obvious stability alerts."})
    if evidence_links:
        for badge in badges:
            label = str(badge.get("label", "")).lower().replace(" ", "_")
            link = evidence_links.get(label) or evidence_links.get("default")
            if link:
                badge["artifact"] = link
    return badges


def fold_variance_rows(base_rows: list[dict[str, Any]], other_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    def _collect(rows: list[dict[str, Any]]) -> dict[str, list[float]]:
        collected: dict[str, list[float]] = {}
        for row in rows:
            if not isinstance(row, dict):
                continue
            for key, value in row.items():
                if isinstance(value, (int, float)):
                    collected.setdefault(str(key), []).append(float(value))
            diag = row.get("diagnostics")
            if isinstance(diag, dict):
                for key, value in diag.items():
                    if isinstance(value, (int, float)):
                        collected.setdefault(f"diagnostics.{key}", []).append(float(value))
        return collected

    left = _collect(base_rows)
    right = _collect(other_rows)
    rows: list[dict[str, Any]] = []
    for metric in sorted(set(left) & set(right)):
        lvals = left.get(metric, [])
        rvals = right.get(metric, [])
        if len(lvals) < 2 or len(rvals) < 2:
            continue
        lstd = statistics.pstdev(lvals)
        rstd = statistics.pstdev(rvals)
        rows.append(
            {
                "metric": metric,
                "base_mean": statistics.mean(lvals),
                "compare_mean": statistics.mean(rvals),
                "delta_mean": statistics.mean(rvals) - statistics.mean(lvals),
                "base_std": lstd,
                "compare_std": rstd,
                "delta_std": rstd - lstd,
                "base_n": len(lvals),
                "compare_n": len(rvals),
            }
        )
    rows.sort(key=lambda row: abs(float(row["delta_std"])), reverse=True)
    return rows


def insights_table_schema(rows: list[dict[str, Any]]) -> list[str]:
    schema: list[str] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        for key in row.keys():
            name = str(key)
            if name not in schema:
                schema.append(name)
    return schema


def build_scenario_comparison(base_payload: dict[str, Any], other_payload: dict[str, Any]) -> list[dict[str, Any]]:
    def _rows(payload: dict[str, Any]) -> dict[str, dict[str, Any]]:
        rows = payload.get("scenario_attribution", []) if isinstance(payload, dict) else []
        out: dict[str, dict[str, Any]] = {}
        if isinstance(rows, list):
            for row in rows:
                if isinstance(row, dict):
                    out[str(row.get("scenario", ""))] = row
        return out

    base_rows = _rows(base_payload)
    other_rows = _rows(other_payload)
    merged: list[dict[str, Any]] = []
    for name in sorted(set(base_rows) | set(other_rows)):
        a = base_rows.get(name, {})
        b = other_rows.get(name, {})
        a_pnl = float(a.get("pnl_total", 0.0))
        b_pnl = float(b.get("pnl_total", 0.0))
        merged.append({
            "scenario": name,
            "base_pnl": a_pnl,
            "compare_pnl": b_pnl,
            "delta_pnl": b_pnl - a_pnl,
            "base_delta_sharpe": float(a.get("delta_sharpe", 0.0)),
            "compare_delta_sharpe": float(b.get("delta_sharpe", 0.0)),
        })
    merged.sort(key=lambda row: abs(float(row["delta_pnl"])), reverse=True)
    return merged


def build_robustness_frontier_view(frontier_payload: dict[str, Any]) -> list[dict[str, Any]]:
    rows = frontier_payload.get("frontier", []) if isinstance(frontier_payload, dict) else []
    if not isinstance(rows, list):
        return []
    normalized: list[dict[str, Any]] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        normalized.append(
            {
                "aum_scale": float(row.get("aum_scale", 0.0)),
                "expected_alpha_net_cost_bps": float(row.get("expected_alpha_net_cost_bps", 0.0)),
                "projected_post_cost_sharpe": float(row.get("projected_post_cost_sharpe", 0.0)),
                "participation_rate": float(row.get("participation_rate", 0.0)),
                "robustness_score": float(row.get("robustness_score", 0.0)),
            }
        )
    normalized.sort(key=lambda row: float(row["aum_scale"]))
    return normalized


def compare_robustness_frontiers(base_payload: dict[str, Any], other_payload: dict[str, Any]) -> list[dict[str, Any]]:
    base_rows = {float(row["aum_scale"]): row for row in build_robustness_frontier_view(base_payload)}
    other_rows = {float(row["aum_scale"]): row for row in build_robustness_frontier_view(other_payload)}
    merged: list[dict[str, Any]] = []
    for scale in sorted(set(base_rows) | set(other_rows)):
        a = base_rows.get(scale, {})
        b = other_rows.get(scale, {})
        base_alpha = float(a.get("expected_alpha_net_cost_bps", 0.0))
        other_alpha = float(b.get("expected_alpha_net_cost_bps", 0.0))
        base_sharpe = float(a.get("projected_post_cost_sharpe", 0.0))
        other_sharpe = float(b.get("projected_post_cost_sharpe", 0.0))
        base_score = float(a.get("robustness_score", 0.0))
        other_score = float(b.get("robustness_score", 0.0))
        merged.append(
            {
                "aum_scale": scale,
                "base_alpha_bps": base_alpha,
                "compare_alpha_bps": other_alpha,
                "delta_alpha_bps": other_alpha - base_alpha,
                "base_sharpe": base_sharpe,
                "compare_sharpe": other_sharpe,
                "delta_sharpe": other_sharpe - base_sharpe,
                "base_robustness": base_score,
                "compare_robustness": other_score,
                "delta_robustness": other_score - base_score,
            }
        )
    merged.sort(key=lambda row: abs(float(row["delta_alpha_bps"])), reverse=True)
    return merged


def read_stress_scenarios(run_dir: Path) -> dict[str, Any]:
    path = run_dir / "stress_scenarios.json"
    if not path.exists():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}
    if isinstance(payload, dict):
        out = dict(payload)
        if "schema_version" not in out:
            out["schema_version"] = "legacy"
        return out
    if isinstance(payload, list):
        return {"schema_version": "legacy", "scenario_attribution": payload}
    return {}


def load_manifest(manifest_path: Path) -> dict[str, Any]:
    try:
        payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(payload, dict):
        return {}

    out = dict(payload)
    manifest_schema = str(out.get("manifest_schema_version", "1.0"))
    warnings: list[str] = []
    if not _version_in_range(manifest_schema, SUPPORTED_MANIFEST_SCHEMA_RANGE[0], SUPPORTED_MANIFEST_SCHEMA_RANGE[1]):
        warnings.append(f"unsupported manifest schema {manifest_schema}")

    metric_manifest = _read_metric_manifest(manifest_path.parent)
    warnings.extend(_validate_metric_manifest(metric_manifest))
    if warnings:
        out["compatibility_warnings"] = warnings
    return out


def index_manifests(output_dir: Path) -> list[dict[str, Any]]:
    indexed: list[dict[str, Any]] = []
    for row in read_experiment_index(output_dir):
        manifest_path_raw = row.get("manifest_path")
        if not manifest_path_raw:
            continue
        manifest_path = Path(str(manifest_path_raw))
        manifest = load_manifest(manifest_path)
        indexed.append({"index": row, "manifest": manifest, "manifest_path": str(manifest_path)})
    return indexed


def search_manifests(indexed_rows: list[dict[str, Any]], *, query: str = "", run_type: str | None = None, tag: str | None = None) -> list[dict[str, Any]]:
    query_l = query.strip().lower()
    tag_l = (tag or "").strip().lower()
    selected: list[dict[str, Any]] = []
    for row in indexed_rows:
        manifest = row.get("manifest", {}) if isinstance(row.get("manifest"), dict) else {}
        if run_type and str(manifest.get("run_type", "")) != run_type:
            continue
        tags = [t.lower() for t in parse_tags(manifest)]
        if tag_l and tag_l not in tags:
            continue
        haystack = json.dumps({"index": row.get("index", {}), "manifest": manifest}, sort_keys=True).lower()
        if query_l and query_l not in haystack:
            continue
        selected.append(row)
    return selected


def compare_manifests(base: dict[str, Any], other: dict[str, Any]) -> dict[str, Any]:
    base_params = base.get("parameters", {}) if isinstance(base.get("parameters"), dict) else {}
    other_params = other.get("parameters", {}) if isinstance(other.get("parameters"), dict) else {}
    base_metrics = base.get("metric_tables", {}) if isinstance(base.get("metric_tables"), dict) else {}
    other_metrics = other.get("metric_tables", {}) if isinstance(other.get("metric_tables"), dict) else {}
    base_deps = base.get("dependency_versions", {}) if isinstance(base.get("dependency_versions"), dict) else {}
    other_deps = other.get("dependency_versions", {}) if isinstance(other.get("dependency_versions"), dict) else {}
    return {
        "parameter_diffs": parameter_diffs(base_params, other_params),
        "metric_table_diffs": parameter_diffs(base_metrics, other_metrics),
        "dependency_diffs": parameter_diffs(base_deps, other_deps),
        "config_hash_changed": str(base.get("config_hash", "")) != str(other.get("config_hash", "")),
        "reproducibility_fingerprint_changed": str(base.get("reproducibility_fingerprint", "")) != str(other.get("reproducibility_fingerprint", "")),
    }

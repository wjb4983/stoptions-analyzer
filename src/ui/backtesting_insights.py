from __future__ import annotations

import json
from pathlib import Path
from typing import Any


def read_experiment_index(output_dir: Path) -> list[dict[str, Any]]:
    jsonl_path = output_dir / "experiment_index.jsonl"
    rows: list[dict[str, Any]] = []
    if jsonl_path.exists():
        for line in jsonl_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line:
                continue
            try:
                parsed = json.loads(line)
            except json.JSONDecodeError:
                continue
            if isinstance(parsed, dict):
                rows.append(parsed)
    return rows


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

    if not badges:
        badges.append({"label": "Guardrails OK", "severity": "low", "reason": "No obvious stability alerts."})
    return badges


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


def read_stress_scenarios(run_dir: Path) -> dict[str, Any]:
    path = run_dir / "stress_scenarios.json"
    if not path.exists():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}
    return payload if isinstance(payload, dict) else {}


def load_manifest(manifest_path: Path) -> dict[str, Any]:
    try:
        payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return payload if isinstance(payload, dict) else {}


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

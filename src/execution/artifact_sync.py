from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
import json
from pathlib import Path
import shutil
from typing import Any

from config import ANALYSIS_OUTPUT_DIR, BACKTEST_OUTPUT_DIR
from backtesting.regime_builder import DEFAULT_REGIME_OUTPUT_DIR

ARTIFACT_MANIFEST_NAME = "artifact_manifest.json"
DEFAULT_REMOTE_NAMESPACE_PREFIX = "remote_sync"

_SUMMARY_PRIORITY_FILES = (
    "manifest.json",
    "artifact_manifest.json",
    "metric_tables_manifest.json",
    "metrics.json",
    "aggregate_metrics.json",
    "leaderboard.json",
    "summary.json",
    "status.json",
    "artifacts.json",
)


@dataclass(frozen=True)
class ArtifactSyncResult:
    remote_job_id: str
    local_synced_run_dir: Path
    mode: str
    copied_files: list[str]


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _relative(path: Path, root: Path) -> str:
    return path.relative_to(root).as_posix()


def _iter_files(root: Path) -> list[Path]:
    return sorted((p for p in root.rglob("*") if p.is_file()), key=lambda item: item.as_posix())


def _build_entry(path: Path, root: Path) -> dict[str, Any]:
    stat = path.stat()
    return {
        "path": _relative(path, root),
        "size_bytes": int(stat.st_size),
        "modified_at": datetime.fromtimestamp(stat.st_mtime, tz=timezone.utc).isoformat(),
    }


def write_artifact_manifest(
    *,
    run_dir: str | Path,
    workflow: str,
    run_id: str | None = None,
    metadata: dict[str, Any] | None = None,
    manifest_name: str = ARTIFACT_MANIFEST_NAME,
) -> Path:
    root = Path(run_dir).expanduser().resolve()
    root.mkdir(parents=True, exist_ok=True)
    entries = [_build_entry(path, root) for path in _iter_files(root) if path.name != manifest_name]
    payload: dict[str, Any] = {
        "manifest_schema_version": "1.0",
        "manifest_type": "artifact_inventory",
        "workflow": str(workflow),
        "run_id": str(run_id or root.name),
        "generated_at": _utc_now(),
        "run_dir": ".",
        "artifact_count": len(entries),
        "artifacts": entries,
        "metadata": dict(metadata or {}),
    }
    output_path = root / manifest_name
    output_path.write_text(json.dumps(payload, indent=2, sort_keys=True), encoding="utf-8")
    return output_path


def _summary_manifest_files(source_root: Path) -> list[str]:
    selected: list[str] = []
    for name in _SUMMARY_PRIORITY_FILES:
        candidate = source_root / name
        if candidate.exists() and candidate.is_file():
            selected.append(name)
    return selected


def sync_run_artifacts(
    *,
    remote_job_id: str,
    remote_run_dir: str | Path,
    local_output_root: str | Path,
    mode: str = "summary",
    namespace_prefix: str = DEFAULT_REMOTE_NAMESPACE_PREFIX,
    include_files: list[str] | None = None,
) -> ArtifactSyncResult:
    source_root = Path(remote_run_dir).expanduser().resolve()
    if not source_root.exists() or not source_root.is_dir():
        raise FileNotFoundError(f"Remote run directory not found: {source_root}")

    mode_normalized = str(mode).strip().lower()
    if mode_normalized not in {"summary", "full"}:
        raise ValueError("mode must be either 'summary' or 'full'")

    output_root = Path(local_output_root).expanduser().resolve()
    output_root.mkdir(parents=True, exist_ok=True)
    target_root = output_root / f"{namespace_prefix}__{remote_job_id}"
    target_root.mkdir(parents=True, exist_ok=True)

    if mode_normalized == "full":
        selected_files = [_relative(path, source_root) for path in _iter_files(source_root)]
    elif include_files:
        selected_files = [str(path).strip().replace("\\", "/") for path in include_files if str(path).strip()]
    else:
        selected_files = _summary_manifest_files(source_root)

    copied: list[str] = []
    for rel_path in selected_files:
        src = source_root / rel_path
        if not src.exists() or not src.is_file():
            continue
        dest = target_root / rel_path
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dest)
        copied.append(rel_path)

    write_artifact_manifest(
        run_dir=target_root,
        workflow="remote_sync",
        run_id=remote_job_id,
        metadata={
            "mode": mode_normalized,
            "source_run_dir": str(source_root),
            "remote_job_id": remote_job_id,
            "namespace_prefix": namespace_prefix,
            "copied_files": copied,
        },
    )
    return ArtifactSyncResult(
        remote_job_id=remote_job_id,
        local_synced_run_dir=target_root,
        mode=mode_normalized,
        copied_files=copied,
    )


def sync_backtest_artifacts(*, remote_job_id: str, remote_run_dir: str | Path, mode: str = "summary") -> ArtifactSyncResult:
    return sync_run_artifacts(
        remote_job_id=remote_job_id,
        remote_run_dir=remote_run_dir,
        local_output_root=BACKTEST_OUTPUT_DIR,
        mode=mode,
    )


def sync_analysis_artifacts(*, remote_job_id: str, remote_run_dir: str | Path, mode: str = "summary") -> ArtifactSyncResult:
    return sync_run_artifacts(
        remote_job_id=remote_job_id,
        remote_run_dir=remote_run_dir,
        local_output_root=ANALYSIS_OUTPUT_DIR,
        mode=mode,
    )


def sync_regime_training_artifacts(*, remote_job_id: str, remote_run_dir: str | Path, mode: str = "summary") -> ArtifactSyncResult:
    return sync_run_artifacts(
        remote_job_id=remote_job_id,
        remote_run_dir=remote_run_dir,
        local_output_root=DEFAULT_REGIME_OUTPUT_DIR,
        mode=mode,
    )


def update_remote_sync_mapping(
    *,
    existing_mapping: dict[str, str] | None,
    remote_job_id: str,
    local_synced_run_dir: str | Path,
) -> dict[str, str]:
    merged = dict(existing_mapping or {})
    merged[str(remote_job_id)] = str(Path(local_synced_run_dir).expanduser().resolve())
    return merged

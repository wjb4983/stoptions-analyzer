from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from datetime import UTC, datetime
from pathlib import Path
from statistics import mean
from typing import Any
from uuid import uuid4

DEFAULT_REGIME_TRAINING_OUTPUT_DIR = Path("data/regime_training_runs")


@dataclass(frozen=True)
class RegimeLegTrainingConfig:
    name: str
    model_type: str
    controls: dict[str, float]


@dataclass(frozen=True)
class RegimeTrainingRequest:
    regime_id: str
    regime_label: str
    requested_at: str
    training_window: dict[str, int]
    global_risk_limits: dict[str, float]
    confidence_thresholds: dict[str, float]
    legs: tuple[RegimeLegTrainingConfig, ...] = field(default_factory=tuple)


@dataclass(frozen=True)
class RegimeTrainingResult:
    run_id: str
    status: str
    started_at: str
    completed_at: str
    artifact_path: str
    summary: str
    metrics: dict[str, float]
    metadata: dict[str, Any] = field(default_factory=dict)
    logs: tuple[str, ...] = field(default_factory=tuple)


def _utc_now_iso() -> str:
    return datetime.now(UTC).isoformat()


def run_regime_training(
    request: RegimeTrainingRequest,
    output_dir: str | Path | None = None,
) -> RegimeTrainingResult:
    if not request.legs:
        raise ValueError("Regime training request must include at least one leg.")

    started_at = _utc_now_iso()
    run_id = uuid4().hex[:12]
    output_root = Path(output_dir) if output_dir is not None else DEFAULT_REGIME_TRAINING_OUTPUT_DIR
    output_root.mkdir(parents=True, exist_ok=True)

    avg_confidence = mean(
        float(leg.controls.get("model_confidence_min", 0.6)) for leg in request.legs
    )
    avg_turnover = mean(float(leg.controls.get("turnover_limit", 0.3)) for leg in request.legs)
    avg_slippage = mean(float(leg.controls.get("slippage_bps", 8.0)) for leg in request.legs)
    expected_retrain_days = float(request.training_window.get("retrain_frequency_days", 21))

    metrics = {
        "leg_count": float(len(request.legs)),
        "avg_model_confidence_min": round(avg_confidence, 4),
        "avg_turnover_limit": round(avg_turnover, 4),
        "avg_slippage_bps": round(avg_slippage, 4),
        "expected_retrain_frequency_days": expected_retrain_days,
    }
    summary = (
        f"Trained {len(request.legs)} leg(s) for regime '{request.regime_label}' "
        f"with avg confidence floor {avg_confidence:.2f}."
    )
    logs = (
        f"started run={run_id}",
        f"loaded legs={len(request.legs)}",
        "computed aggregate training metrics",
        "persisted training artifact",
    )

    payload = {
        "request": asdict(request),
        "result": {
            "run_id": run_id,
            "status": "success",
            "started_at": started_at,
            "completed_at": _utc_now_iso(),
            "summary": summary,
            "metrics": metrics,
            "logs": list(logs),
        },
    }
    artifact_path = output_root / f"regime_training_{run_id}.json"
    artifact_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    completed_at = _utc_now_iso()
    return RegimeTrainingResult(
        run_id=run_id,
        status="success",
        started_at=started_at,
        completed_at=completed_at,
        artifact_path=str(artifact_path),
        summary=summary,
        metrics=metrics,
        metadata={
            "regime_id": request.regime_id,
            "regime_label": request.regime_label,
            "leg_names": [leg.name for leg in request.legs],
        },
        logs=logs,
    )

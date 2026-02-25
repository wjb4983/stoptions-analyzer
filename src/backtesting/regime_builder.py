from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

from models.regime_catalog import validate_model_leg_pairing

DEFAULT_REGIME_OUTPUT_DIR = Path("data/regimes")
SUPPORTED_LEG_FAMILIES = (
    "timeseries_momentum",
    "volatility_risk_premium_selling",
    "cheap_vol_buying",
    "regime_change_detection",
    "volatility_clustering",
    "iv_ev_spread_term_structure",
    "self_exciting_event_intensity",
    "vol_surface_calibration",
)

_REQUIRED_KNOBS_BY_LEG_FAMILY: dict[str, tuple[str, ...]] = {
    "timeseries_momentum": (
        "lookback_days",
        "vol_filter_max",
        "sizing_cap",
        "stop_loss_pct",
    ),
    "volatility_risk_premium_selling": (
        "lookback_days",
        "carry_threshold",
        "vol_filter_min",
        "sizing_cap",
        "stop_loss_pct",
    ),
    "cheap_vol_buying": (
        "lookback_days",
        "carry_threshold",
        "vol_filter_max",
        "sizing_cap",
        "stop_loss_pct",
    ),
    "regime_change_detection": (
        "lookback_days",
        "detection_threshold",
        "sizing_cap",
        "stop_loss_pct",
    ),
    "volatility_clustering": (
        "lookback_days",
        "vol_filter_min",
        "vol_filter_max",
        "sizing_cap",
        "stop_loss_pct",
    ),
    "iv_ev_spread_term_structure": (
        "lookback_days",
        "carry_threshold",
        "vol_filter_min",
        "sizing_cap",
        "stop_loss_pct",
    ),
    "self_exciting_event_intensity": (
        "lookback_days",
        "detection_threshold",
        "vol_filter_min",
        "sizing_cap",
        "stop_loss_pct",
    ),
    "vol_surface_calibration": (
        "lookback_days",
        "detection_threshold",
        "vol_filter_max",
        "sizing_cap",
        "stop_loss_pct",
    ),
}


@dataclass(frozen=True)
class TrainingSpec:
    train_start: str
    train_end: str
    validation_window_days: int
    retrain_frequency_days: int


@dataclass(frozen=True)
class RiskSpec:
    max_gross_exposure: float
    max_net_exposure: float
    max_drawdown_pct: float
    default_sizing_cap: float
    default_stop_loss_pct: float


@dataclass(frozen=True)
class ModelChoiceSpec:
    model_name: str
    objective: str
    hyperparameters: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class RegimeLegSpec:
    name: str
    leg_family: str
    knobs: dict[str, Any] = field(default_factory=dict)
    enabled: bool = True


@dataclass(frozen=True)
class RegimeSpec:
    regime_name: str
    training: TrainingSpec
    risk: RiskSpec
    model_choice: ModelChoiceSpec
    legs: tuple[RegimeLegSpec, ...] = field(default_factory=tuple)
    schema_version: int = 1

    def to_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        payload["legs"] = [asdict(leg) for leg in self.legs]
        return payload

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "RegimeSpec":
        legs = tuple(RegimeLegSpec(**leg_payload) for leg_payload in payload.get("legs", []))
        return cls(
            regime_name=str(payload["regime_name"]),
            schema_version=int(payload.get("schema_version", 1)),
            training=TrainingSpec(**payload["training"]),
            risk=RiskSpec(**payload["risk"]),
            model_choice=ModelChoiceSpec(**payload["model_choice"]),
            legs=legs,
        )


def _require_numeric(knobs: dict[str, Any], key: str, *, lower_exclusive: float | None = None, lower_inclusive: float | None = None, upper_inclusive: float | None = None) -> None:
    value = knobs[key]
    if not isinstance(value, (int, float)):
        raise ValueError(f"{key} must be numeric")
    cast = float(value)
    if lower_exclusive is not None and cast <= lower_exclusive:
        raise ValueError(f"{key} must be > {lower_exclusive}")
    if lower_inclusive is not None and cast < lower_inclusive:
        raise ValueError(f"{key} must be >= {lower_inclusive}")
    if upper_inclusive is not None and cast > upper_inclusive:
        raise ValueError(f"{key} must be <= {upper_inclusive}")


def validate_leg_spec(leg: RegimeLegSpec) -> None:
    if leg.leg_family not in SUPPORTED_LEG_FAMILIES:
        raise ValueError(f"Unsupported leg_family: {leg.leg_family}")

    required_knobs = _REQUIRED_KNOBS_BY_LEG_FAMILY[leg.leg_family]
    missing = [key for key in required_knobs if key not in leg.knobs]
    if missing:
        raise ValueError(f"Missing required knobs for {leg.leg_family}: {', '.join(missing)}")

    if "lookback_days" in leg.knobs:
        _require_numeric(leg.knobs, "lookback_days", lower_exclusive=0)
    if "carry_threshold" in leg.knobs:
        _require_numeric(leg.knobs, "carry_threshold", lower_inclusive=-1.0, upper_inclusive=1.0)
    if "vol_filter_min" in leg.knobs:
        _require_numeric(leg.knobs, "vol_filter_min", lower_inclusive=0.0)
    if "vol_filter_max" in leg.knobs:
        _require_numeric(leg.knobs, "vol_filter_max", lower_inclusive=0.0)
    if "vol_filter_min" in leg.knobs and "vol_filter_max" in leg.knobs:
        if float(leg.knobs["vol_filter_min"]) > float(leg.knobs["vol_filter_max"]):
            raise ValueError("vol_filter_min must be <= vol_filter_max")
    if "detection_threshold" in leg.knobs:
        _require_numeric(leg.knobs, "detection_threshold", lower_exclusive=0.0)

    _require_numeric(leg.knobs, "sizing_cap", lower_exclusive=0.0, upper_inclusive=1.0)
    _require_numeric(leg.knobs, "stop_loss_pct", lower_exclusive=0.0, upper_inclusive=1.0)


def validate_regime_spec(spec: RegimeSpec) -> None:
    if not spec.regime_name.strip():
        raise ValueError("regime_name must be non-empty")
    if not spec.legs:
        raise ValueError("at least one leg must be configured")

    for leg in spec.legs:
        validate_leg_spec(leg)
        validate_model_leg_pairing(leg.leg_family, spec.model_choice.model_name)

    if spec.risk.max_gross_exposure <= 0:
        raise ValueError("max_gross_exposure must be > 0")
    if spec.risk.default_sizing_cap <= 0 or spec.risk.default_sizing_cap > 1:
        raise ValueError("default_sizing_cap must be in (0, 1]")
    if spec.risk.default_stop_loss_pct <= 0 or spec.risk.default_stop_loss_pct > 1:
        raise ValueError("default_stop_loss_pct must be in (0, 1]")



def _sanitize_name(value: str) -> str:
    sanitized = re.sub(r"[^a-zA-Z0-9_-]+", "_", value.strip()).strip("_")
    return sanitized.lower() or "regime"


def save_regime_spec(spec: RegimeSpec, output_dir: str | Path | None = None) -> Path:
    validate_regime_spec(spec)
    base_dir = Path(output_dir) if output_dir is not None else DEFAULT_REGIME_OUTPUT_DIR
    base_dir.mkdir(parents=True, exist_ok=True)

    file_path = base_dir / f"{_sanitize_name(spec.regime_name)}.json"
    file_path.write_text(json.dumps(spec.to_dict(), indent=2, sort_keys=True), encoding="utf-8")
    return file_path


def load_regime_spec(regime_name_or_path: str | Path, output_dir: str | Path | None = None) -> RegimeSpec:
    candidate = Path(regime_name_or_path)
    if candidate.suffix.lower() != ".json":
        base_dir = Path(output_dir) if output_dir is not None else DEFAULT_REGIME_OUTPUT_DIR
        candidate = base_dir / f"{_sanitize_name(str(regime_name_or_path))}.json"

    payload = json.loads(candidate.read_text(encoding="utf-8"))
    spec = RegimeSpec.from_dict(payload)
    validate_regime_spec(spec)
    return spec

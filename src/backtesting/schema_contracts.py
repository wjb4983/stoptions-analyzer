from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class SchemaContract:
    name: str
    current_version: str
    minimum_compatible_version: str

    def is_compatible(self, version: str) -> bool:
        candidate = _parse_semver(version)
        minimum = _parse_semver(self.minimum_compatible_version)
        current = _parse_semver(self.current_version)
        if candidate[0] != current[0]:
            return False
        return minimum <= candidate <= current


REGIME_TRAINING_MANIFEST_CONTRACT = SchemaContract(
    name="regime_training_manifest",
    current_version="2.1.0",
    minimum_compatible_version="2.0.0",
)

EXPORT_BUNDLE_MANIFEST_CONTRACT = SchemaContract(
    name="export_bundle_manifest",
    current_version="1.1.0",
    minimum_compatible_version="1.0.0",
)

BACKTEST_HYDRATION_PAYLOAD_CONTRACT = SchemaContract(
    name="backtest_hydration_payload",
    current_version="1.1.0",
    minimum_compatible_version="1.0.0",
)


def _parse_semver(version: str) -> tuple[int, int, int]:
    raw = str(version).strip()
    parts = raw.split(".")
    if len(parts) < 2:
        raise ValueError(f"Invalid semantic version '{version}'.")
    while len(parts) < 3:
        parts.append("0")
    try:
        major, minor, patch = (int(parts[0]), int(parts[1]), int(parts[2]))
    except ValueError as exc:
        raise ValueError(f"Invalid semantic version '{version}'.") from exc
    if major < 0 or minor < 0 or patch < 0:
        raise ValueError(f"Invalid semantic version '{version}'.")
    return major, minor, patch

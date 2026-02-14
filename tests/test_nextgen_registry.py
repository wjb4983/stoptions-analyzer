from __future__ import annotations

from modeling_nextgen.adapters.registry_adapter import LegacyRegistryAdapter
from modeling_nextgen.core.config import NextGenModelingConfig
from modeling_nextgen.core.registry import NextGenRegistry
from models.registry import ModelMetadataRegistry


def test_nextgen_registry_exposes_model_metadata() -> None:
    registry = NextGenRegistry(backend=ModelMetadataRegistry())
    registry.register_model(
        "model_v3",
        lineage=("model_v1", "model_v2"),
        hyperparams={"learning_rate": 0.01, "max_depth": 4},
        calibration_version="cal-v2026.02",
        robustness_passed=True,
        stress_passed=False,
        deployment_slot="candidate",
    )

    assert registry.model_lineage("model_v3") == ("model_v1", "model_v2")
    assert registry.model_hyperparams("model_v3") == {"learning_rate": 0.01, "max_depth": 4}
    assert registry.calibration_version("model_v3") == "cal-v2026.02"
    assert registry.robustness_and_stress_status("model_v3") == (True, False)


def test_lookup_api_for_slot_promotion() -> None:
    backend = ModelMetadataRegistry()
    adapter = LegacyRegistryAdapter(config=NextGenModelingConfig(enabled=True), metadata_registry=NextGenRegistry(backend=backend))
    adapter.register_model_metadata(
        "model_champ",
        lineage=("model_base",),
        hyperparams={"alpha": 0.2},
        calibration_version="cal-v1",
        robustness_passed=True,
        stress_passed=True,
        deployment_slot="champion",
    )
    adapter.register_model_metadata(
        "model_chal",
        lineage=("model_champ",),
        hyperparams={"alpha": 0.25},
        calibration_version="cal-v2",
        robustness_passed=True,
        stress_passed=True,
        deployment_slot="challenger",
    )

    assert adapter.lookup_for_slot_promotion() == "model_chal"

    backend.register(
        "model_bad",
        lineage=("model_champ",),
        hyperparams={"alpha": 0.3},
        calibration_version=None,
        robustness_passed=True,
        stress_passed=True,
        deployment_slot="challenger",
    )
    assert adapter.lookup_for_slot_promotion() is None

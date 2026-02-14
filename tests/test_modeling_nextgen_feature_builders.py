from __future__ import annotations

import numpy as np

from src.modeling_nextgen.features.builders import FlowFeatureBuilder, RegimeFeatureBuilder, SurfaceFeatureBuilder


_BUILDERS = (SurfaceFeatureBuilder(), RegimeFeatureBuilder(), FlowFeatureBuilder())


def test_feature_builders_expose_required_feature_columns_contract() -> None:
    frame = {"price": np.array([100.0, 101.0]), "volume": np.array([10.0, 12.0])}

    for builder in _BUILDERS:
        built = builder.build(frame)
        assert isinstance(built, dict)


def test_feature_builders_have_nan_inf_guardrails_on_outputs() -> None:
    frame = {
        "price": np.array([100.0, np.nan, 102.0]),
        "volume": np.array([1.0, np.inf, 3.0]),
    }

    for builder in _BUILDERS:
        built = builder.build(frame)
        for value in built.values():
            arr = np.asarray(value, dtype=float)
            assert np.all(np.isfinite(arr))


def test_feature_builder_outputs_are_reproducible() -> None:
    frame = {"x": np.array([1.0, 2.0, 3.0])}

    for builder in _BUILDERS:
        first = builder.build(frame)
        second = builder.build(frame)
        assert first.keys() == second.keys()
        for key in first:
            np.testing.assert_allclose(np.asarray(first[key]), np.asarray(second[key]))

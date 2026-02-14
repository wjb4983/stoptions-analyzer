from __future__ import annotations

import numpy as np
import pytest

from src.modeling_nextgen.models.deep.cross_asset_graph import CrossAssetGraphModel


def _graph_inputs(n_samples: int = 20, n_nodes: int = 4) -> tuple[dict[str, np.ndarray], np.ndarray]:
    rng = np.random.default_rng(7)
    features = {
        "returns": rng.normal(size=(n_samples, n_nodes)),
        "sector_ids": np.array([0, 1, 1, 2])[:n_nodes],
        "macro_exposures": rng.normal(size=(n_nodes, 3)),
    }
    labels = rng.normal(size=(n_samples, n_nodes))
    return features, labels


def test_dynamic_graph_construction_invariants() -> None:
    features, _ = _graph_inputs()
    model = CrossAssetGraphModel(use_deep_stack=False, sparsity_top_k=2)

    adjacency = model._build_dynamic_adjacency(features)

    assert adjacency.shape == (4, 4)
    assert np.all(np.isfinite(adjacency))
    np.testing.assert_allclose(np.diag(adjacency), np.zeros(4), atol=1e-12)

    row_sums = adjacency.sum(axis=1)
    np.testing.assert_allclose(row_sums, np.ones(4), atol=1e-8)


def test_cross_asset_graph_handles_disconnected_graph_without_nan() -> None:
    features = {
        "returns": np.zeros((8, 3), dtype=float),
        "sector_ids": np.array([0, 1, 2]),
        "macro_exposures": np.zeros((3, 2), dtype=float),
    }
    labels = np.zeros((8, 3), dtype=float)

    model = CrossAssetGraphModel(use_deep_stack=False, epochs=5)
    model.fit(features, labels)
    result = model.predict(features)

    assert result.predictions.shape == (8, 3)
    assert np.all(np.isfinite(result.predictions))
    assert result.uncertainty is not None
    assert np.all(np.isfinite(result.uncertainty))


def test_cross_asset_graph_rejects_invalid_inputs() -> None:
    model = CrossAssetGraphModel(use_deep_stack=False)

    with pytest.raises(ValueError, match="Missing required inputs"):
        model.fit({}, np.zeros(4))

    with pytest.raises(ValueError, match="sector_ids must have shape"):
        model.fit(
            {
                "returns": np.ones((10, 3)),
                "sector_ids": np.ones((10, 3)),
            },
            np.ones((10, 3)),
        )

    with pytest.raises(ValueError, match="labels must have shape"):
        model.fit({"returns": np.ones((10, 3))}, np.ones((9, 3)))

from __future__ import annotations

import numpy as np
import pytest

from src.modeling_nextgen.models.ml.panel_baselines import (
    ElasticNetBaseline,
    PanelWalkForwardSplitter,
)


def test_panel_splitter_enforces_temporal_gap_and_disjoint_indices() -> None:
    dates = np.array(
        [
            "2024-01-01",
            "2024-01-01",
            "2024-01-02",
            "2024-01-02",
            "2024-01-03",
            "2024-01-03",
            "2024-01-04",
            "2024-01-04",
            "2024-01-05",
            "2024-01-05",
        ],
        dtype="datetime64[D]",
    )
    asset_ids = np.array(["A", "B"] * 5)
    splitter = PanelWalkForwardSplitter(train_dates=2, test_dates=1, purge_dates=1, embargo_dates=0, label_horizon_dates=1)

    folds = list(splitter.split(asset_ids=asset_ids, dates=dates))
    assert len(folds) == 2

    for fold in folds:
        assert all(fold.leakage_checks.values())
        assert np.intersect1d(fold.train_idx, fold.test_idx).size == 0
        assert fold.train_end_date < fold.test_start_date


def test_baseline_requires_fit_before_predict() -> None:
    model = ElasticNetBaseline(alpha=0.1)
    with pytest.raises(RuntimeError, match="not fitted"):
        model.predict(np.zeros((2, 3), dtype=float))

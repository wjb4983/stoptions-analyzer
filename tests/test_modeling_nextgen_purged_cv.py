from __future__ import annotations

import numpy as np

from src.backtesting.validation import (
    generate_combinatorial_purged_cv_splits as bt_generate_combinatorial_purged_cv_splits,
)
from src.backtesting.validation import generate_purged_kfold_splits as bt_generate_purged_kfold_splits
from src.modeling_nextgen.validation.purged_cv import (
    generate_combinatorial_purged_cv_splits,
    generate_grouped_combinatorial_purged_cv_splits,
    generate_grouped_purged_kfold_splits,
    generate_purged_kfold_splits,
)


def test_nextgen_purged_kfold_compatible_with_backtesting() -> None:
    got = generate_purged_kfold_splits(
        n_samples=31,
        n_splits=4,
        purge_window_bars=2,
        embargo_window_bars=1,
        label_horizon_bars=3,
    )
    expected = bt_generate_purged_kfold_splits(
        n_samples=31,
        n_splits=4,
        purge_window_bars=2,
        embargo_window_bars=1,
        label_horizon_bars=3,
    )

    assert [s.train_indices.tolist() for s in got] == [s.train_indices.tolist() for s in expected]
    assert [s.test_indices.tolist() for s in got] == [s.test_indices.tolist() for s in expected]
    assert [s.metadata for s in got] == [s.metadata for s in expected]


def test_nextgen_cpcv_compatible_with_backtesting() -> None:
    got = generate_combinatorial_purged_cv_splits(
        n_samples=30,
        n_groups=5,
        n_test_groups=2,
        purge_window_bars=2,
        embargo_window_bars=2,
        seed=42,
    )
    expected = bt_generate_combinatorial_purged_cv_splits(
        n_samples=30,
        n_groups=5,
        n_test_groups=2,
        purge_window_bars=2,
        embargo_window_bars=2,
        seed=42,
    )

    assert [s.train_indices.tolist() for s in got] == [s.train_indices.tolist() for s in expected]
    assert [s.test_indices.tolist() for s in got] == [s.test_indices.tolist() for s in expected]
    assert [s.metadata for s in got] == [s.metadata for s in expected]


def test_grouped_purged_kfold_respects_time_boundaries_across_assets() -> None:
    asset_ids = np.array(["A", "B", "A", "B", "A", "B", "A", "B"])
    timestamps = np.array([1, 1, 2, 2, 3, 3, 4, 4])

    splits = generate_grouped_purged_kfold_splits(
        asset_ids=asset_ids,
        timestamps=timestamps,
        n_splits=4,
        purge_window_bars=1,
        embargo_window_bars=1,
    )

    first = splits[0]
    assert first.test_indices.tolist() == [0, 1]
    assert first.train_indices.tolist() == [4, 5, 6, 7]


def test_grouped_cpcv_holds_out_timestamp_groups_not_single_rows() -> None:
    asset_ids = np.array(["A", "B", "C"] * 4)
    timestamps = np.array([1, 1, 1, 2, 2, 2, 3, 3, 3, 4, 4, 4])

    splits = generate_grouped_combinatorial_purged_cv_splits(
        asset_ids=asset_ids,
        timestamps=timestamps,
        n_groups=4,
        n_test_groups=2,
        purge_window_bars=0,
        embargo_window_bars=0,
        label_horizon_bars=1,
        seed=0,
    )

    split = next(s for s in splits if s.metadata["held_out_groups"] == [1, 3])
    assert split.test_indices.tolist() == [3, 4, 5, 9, 10, 11]
    assert set(split.train_indices.tolist()).isdisjoint(split.test_indices.tolist())

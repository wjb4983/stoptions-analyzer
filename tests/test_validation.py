from __future__ import annotations

from src.backtesting.validation import generate_combinatorial_purged_cv_splits, generate_purged_kfold_splits


def test_purged_kfold_has_no_train_test_overlap() -> None:
    splits = generate_purged_kfold_splits(
        n_samples=100,
        n_splits=5,
        purge_window_bars=2,
        embargo_window_bars=3,
        label_horizon_bars=1,
    )
    assert splits
    for split in splits:
        train = set(split.train_indices.tolist())
        test = set(split.test_indices.tolist())
        assert train.isdisjoint(test)


def test_purged_kfold_embargo_window_is_respected() -> None:
    splits = generate_purged_kfold_splits(
        n_samples=50,
        n_splits=5,
        purge_window_bars=1,
        embargo_window_bars=4,
        label_horizon_bars=1,
    )
    first = splits[0]
    test_end = int(first.metadata["test_boundary"][1])
    embargo_end = int(first.metadata["embargo_boundary"][1])
    embargo_slice = set(range(test_end, embargo_end))
    train = set(first.train_indices.tolist())
    assert embargo_slice.isdisjoint(train)


def test_cpcv_split_order_is_deterministic_for_seed() -> None:
    left = generate_combinatorial_purged_cv_splits(
        n_samples=120,
        n_groups=6,
        n_test_groups=2,
        purge_window_bars=1,
        embargo_window_bars=1,
        seed=123,
    )
    right = generate_combinatorial_purged_cv_splits(
        n_samples=120,
        n_groups=6,
        n_test_groups=2,
        purge_window_bars=1,
        embargo_window_bars=1,
        seed=123,
    )
    assert [split.metadata["held_out_groups"] for split in left] == [split.metadata["held_out_groups"] for split in right]


def test_cpcv_test_indices_include_only_held_out_groups_for_adjacent_groups() -> None:
    splits = generate_combinatorial_purged_cv_splits(
        n_samples=12,
        n_groups=4,
        n_test_groups=2,
        purge_window_bars=3,
        embargo_window_bars=3,
        seed=5,
    )

    split = next(s for s in splits if s.metadata["held_out_groups"] == [1, 2])
    assert split.test_indices.tolist() == [3, 4, 5, 6, 7, 8]
    assert set(split.train_indices.tolist()).isdisjoint(set(split.test_indices.tolist()))


def test_cpcv_test_indices_include_only_held_out_groups_for_non_adjacent_groups() -> None:
    splits = generate_combinatorial_purged_cv_splits(
        n_samples=12,
        n_groups=4,
        n_test_groups=2,
        purge_window_bars=3,
        embargo_window_bars=3,
        seed=11,
    )

    split = next(s for s in splits if s.metadata["held_out_groups"] == [0, 2])
    assert split.test_indices.tolist() == [0, 1, 2, 6, 7, 8]
    assert set(split.train_indices.tolist()).isdisjoint(set(split.test_indices.tolist()))


def test_cpcv_large_purge_embargo_do_not_expand_test_set() -> None:
    splits = generate_combinatorial_purged_cv_splits(
        n_samples=15,
        n_groups=5,
        n_test_groups=1,
        purge_window_bars=10,
        embargo_window_bars=10,
        seed=0,
    )

    split = next(s for s in splits if s.metadata["held_out_groups"] == [2])
    assert split.test_indices.tolist() == [6, 7, 8]
    assert split.train_indices.size == 0
    assert split.metadata["purge_ranges"] == [[0, 6]]
    assert split.metadata["embargo_ranges"] == [[9, 15]]

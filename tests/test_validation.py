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

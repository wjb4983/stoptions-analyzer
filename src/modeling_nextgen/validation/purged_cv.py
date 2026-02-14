from __future__ import annotations

import itertools
from typing import Sequence

import numpy as np

from backtesting.validation import (
    PurgedSplit,
    generate_combinatorial_purged_cv_splits as _generate_combinatorial_purged_cv_splits,
    generate_purged_kfold_splits as _generate_purged_kfold_splits,
)


def generate_purged_kfold_splits(
    *,
    n_samples: int,
    n_splits: int,
    purge_window_bars: int = 0,
    embargo_window_bars: int = 0,
    label_horizon_bars: int = 1,
) -> list[PurgedSplit]:
    """Compatibility shim for backtesting purged k-fold semantics."""

    return _generate_purged_kfold_splits(
        n_samples=n_samples,
        n_splits=n_splits,
        purge_window_bars=purge_window_bars,
        embargo_window_bars=embargo_window_bars,
        label_horizon_bars=label_horizon_bars,
    )


def generate_combinatorial_purged_cv_splits(
    *,
    n_samples: int,
    n_groups: int,
    n_test_groups: int,
    purge_window_bars: int = 0,
    embargo_window_bars: int = 0,
    label_horizon_bars: int = 1,
    seed: int | None = None,
) -> list[PurgedSplit]:
    """Compatibility shim for backtesting CPCV semantics."""

    return _generate_combinatorial_purged_cv_splits(
        n_samples=n_samples,
        n_groups=n_groups,
        n_test_groups=n_test_groups,
        purge_window_bars=purge_window_bars,
        embargo_window_bars=embargo_window_bars,
        label_horizon_bars=label_horizon_bars,
        seed=seed,
    )


def generate_grouped_purged_kfold_splits(
    *,
    asset_ids: Sequence[object],
    timestamps: Sequence[object],
    n_splits: int,
    purge_window_bars: int = 0,
    embargo_window_bars: int = 0,
    label_horizon_bars: int = 1,
) -> list[PurgedSplit]:
    """Purged K-Fold where test boundaries are defined on unique timestamps across assets."""

    return _generate_grouped_splits(
        asset_ids=asset_ids,
        timestamps=timestamps,
        n_groups=n_splits,
        n_test_groups=1,
        purge_window_bars=purge_window_bars,
        embargo_window_bars=embargo_window_bars,
        label_horizon_bars=label_horizon_bars,
        seed=None,
        scheme="grouped_purged_kfold",
    )


def generate_grouped_combinatorial_purged_cv_splits(
    *,
    asset_ids: Sequence[object],
    timestamps: Sequence[object],
    n_groups: int,
    n_test_groups: int,
    purge_window_bars: int = 0,
    embargo_window_bars: int = 0,
    label_horizon_bars: int = 1,
    seed: int | None = None,
) -> list[PurgedSplit]:
    """CPCV where test boundaries are defined on unique timestamps across assets."""

    return _generate_grouped_splits(
        asset_ids=asset_ids,
        timestamps=timestamps,
        n_groups=n_groups,
        n_test_groups=n_test_groups,
        purge_window_bars=purge_window_bars,
        embargo_window_bars=embargo_window_bars,
        label_horizon_bars=label_horizon_bars,
        seed=seed,
        scheme="grouped_cpcv",
    )


def _generate_grouped_splits(
    *,
    asset_ids: Sequence[object],
    timestamps: Sequence[object],
    n_groups: int,
    n_test_groups: int,
    purge_window_bars: int,
    embargo_window_bars: int,
    label_horizon_bars: int,
    seed: int | None,
    scheme: str,
) -> list[PurgedSplit]:
    asset_arr = np.asarray(asset_ids)
    ts_arr = np.asarray(timestamps)
    if asset_arr.ndim != 1 or ts_arr.ndim != 1:
        raise ValueError("asset_ids and timestamps must be 1D")
    if asset_arr.shape[0] != ts_arr.shape[0]:
        raise ValueError("asset_ids and timestamps must have the same length")
    if asset_arr.size == 0:
        raise ValueError("asset_ids and timestamps cannot be empty")
    if n_groups < 2:
        raise ValueError("n_groups must be >= 2")
    if n_groups > np.unique(ts_arr).size:
        raise ValueError("n_groups cannot exceed number of unique timestamps")
    if n_test_groups <= 0 or n_test_groups >= n_groups:
        raise ValueError("n_test_groups must be in [1, n_groups-1]")

    purge = max(int(purge_window_bars), int(label_horizon_bars))
    embargo = max(int(embargo_window_bars), int(label_horizon_bars))

    unique_ts = np.unique(ts_arr)
    unique_ts.sort()
    boundaries = np.linspace(0, unique_ts.size, n_groups + 1, dtype=int)

    ts_positions = np.searchsorted(unique_ts, ts_arr)

    group_ids = np.arange(n_groups, dtype=int)
    held_out_combinations = list(itertools.combinations(group_ids.tolist(), int(n_test_groups)))
    if seed is not None:
        rng = np.random.default_rng(int(seed))
        rng.shuffle(held_out_combinations)

    all_idx = np.arange(ts_arr.size, dtype=int)
    splits: list[PurgedSplit] = []

    for split_id, held_out_groups in enumerate(held_out_combinations):
        held_out_time_mask = np.zeros(unique_ts.size, dtype=bool)
        test_ranges: list[list[int]] = []
        purge_ranges: list[list[int]] = []
        embargo_ranges: list[list[int]] = []

        for group_id in held_out_groups:
            start = int(boundaries[group_id])
            end = int(boundaries[group_id + 1])
            held_out_time_mask[start:end] = True
            test_ranges.append([start, end])

        train_time_mask = ~held_out_time_mask
        for start, end in test_ranges:
            purge_start = max(0, start - purge)
            embargo_end = min(unique_ts.size, end + embargo)
            train_time_mask[purge_start:start] = False
            train_time_mask[end:embargo_end] = False
            purge_ranges.append([purge_start, start])
            embargo_ranges.append([end, embargo_end])

        row_test_mask = held_out_time_mask[ts_positions]
        row_train_mask = train_time_mask[ts_positions]

        splits.append(
            PurgedSplit(
                split_id=split_id,
                train_indices=all_idx[row_train_mask],
                test_indices=all_idx[row_test_mask],
                purge_window_bars=purge,
                embargo_window_bars=embargo,
                label_horizon_bars=int(label_horizon_bars),
                metadata={
                    "scheme": scheme,
                    "n_groups": int(n_groups),
                    "n_test_groups": int(n_test_groups),
                    "held_out_groups": [int(x) for x in held_out_groups],
                    "group_boundaries": [[int(boundaries[i]), int(boundaries[i + 1])] for i in range(n_groups)],
                    "test_ranges": test_ranges,
                    "purge_ranges": purge_ranges,
                    "embargo_ranges": embargo_ranges,
                },
            )
        )

    return splits


__all__ = [
    "PurgedSplit",
    "generate_purged_kfold_splits",
    "generate_combinatorial_purged_cv_splits",
    "generate_grouped_purged_kfold_splits",
    "generate_grouped_combinatorial_purged_cv_splits",
]

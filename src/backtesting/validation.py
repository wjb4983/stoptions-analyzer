from __future__ import annotations

import itertools
from dataclasses import dataclass
from typing import Any

import numpy as np


@dataclass(frozen=True)
class PurgedSplit:
    split_id: int
    train_indices: np.ndarray
    test_indices: np.ndarray
    purge_window_bars: int
    embargo_window_bars: int
    label_horizon_bars: int
    metadata: dict[str, Any]

    def to_dict(self) -> dict[str, Any]:
        return {
            "split_id": int(self.split_id),
            "train_size": int(self.train_indices.size),
            "test_size": int(self.test_indices.size),
            "train_ranges": _indices_to_ranges(self.train_indices),
            "test_ranges": _indices_to_ranges(self.test_indices),
            "purge_window_bars": int(self.purge_window_bars),
            "embargo_window_bars": int(self.embargo_window_bars),
            "label_horizon_bars": int(self.label_horizon_bars),
            "metadata": dict(self.metadata),
        }


def generate_purged_kfold_splits(
    *,
    n_samples: int,
    n_splits: int,
    purge_window_bars: int = 0,
    embargo_window_bars: int = 0,
    label_horizon_bars: int = 1,
) -> list[PurgedSplit]:
    if n_samples <= 0:
        raise ValueError("n_samples must be positive")
    if n_splits < 2:
        raise ValueError("n_splits must be >= 2")
    if n_splits > n_samples:
        raise ValueError("n_splits cannot exceed n_samples")

    purge = max(int(purge_window_bars), int(label_horizon_bars))
    embargo = max(int(embargo_window_bars), int(label_horizon_bars))
    boundaries = np.linspace(0, n_samples, n_splits + 1, dtype=int)
    all_idx = np.arange(n_samples, dtype=int)
    splits: list[PurgedSplit] = []

    for split_id in range(n_splits):
        test_start = int(boundaries[split_id])
        test_end = int(boundaries[split_id + 1])
        test_idx = all_idx[test_start:test_end]
        train_mask = np.ones(n_samples, dtype=bool)
        purge_start = max(0, test_start - purge)
        embargo_end = min(n_samples, test_end + embargo)
        train_mask[purge_start:embargo_end] = False
        train_idx = all_idx[train_mask]
        splits.append(
            PurgedSplit(
                split_id=split_id,
                train_indices=train_idx,
                test_indices=test_idx,
                purge_window_bars=purge,
                embargo_window_bars=embargo,
                label_horizon_bars=int(label_horizon_bars),
                metadata={
                    "scheme": "purged_kfold",
                    "test_boundary": [test_start, test_end],
                    "purge_boundary": [purge_start, test_start],
                    "embargo_boundary": [test_end, embargo_end],
                },
            )
        )
    return splits


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
    if n_samples <= 0:
        raise ValueError("n_samples must be positive")
    if n_groups < 3:
        raise ValueError("n_groups must be at least 3")
    if n_test_groups <= 0 or n_test_groups >= n_groups:
        raise ValueError("n_test_groups must be in [1, n_groups-1]")

    purge = max(int(purge_window_bars), int(label_horizon_bars))
    embargo = max(int(embargo_window_bars), int(label_horizon_bars))
    boundaries = np.linspace(0, n_samples, n_groups + 1, dtype=int)
    group_ids = np.arange(n_groups, dtype=int)
    combinations = list(itertools.combinations(group_ids.tolist(), int(n_test_groups)))
    if seed is not None:
        rng = np.random.default_rng(int(seed))
        rng.shuffle(combinations)

    all_idx = np.arange(n_samples, dtype=int)
    splits: list[PurgedSplit] = []
    for split_id, held_out_groups in enumerate(combinations):
        held_out_mask = np.zeros(n_samples, dtype=bool)
        test_ranges: list[list[int]] = []
        purge_ranges: list[list[int]] = []
        embargo_ranges: list[list[int]] = []
        for group_id in held_out_groups:
            test_start = int(boundaries[group_id])
            test_end = int(boundaries[group_id + 1])
            held_out_mask[test_start:test_end] = True
            test_ranges.append([test_start, test_end])

        train_mask = ~held_out_mask
        for test_start, test_end in test_ranges:
            purge_start = max(0, test_start - purge)
            embargo_end = min(n_samples, test_end + embargo)
            train_mask[purge_start:test_start] = False
            train_mask[test_end:embargo_end] = False
            purge_ranges.append([purge_start, test_start])
            embargo_ranges.append([test_end, embargo_end])

        test_idx = all_idx[held_out_mask]
        train_idx = all_idx[train_mask]
        splits.append(
            PurgedSplit(
                split_id=split_id,
                train_indices=train_idx,
                test_indices=test_idx,
                purge_window_bars=purge,
                embargo_window_bars=embargo,
                label_horizon_bars=int(label_horizon_bars),
                metadata={
                    "scheme": "cpcv",
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


def _indices_to_ranges(indices: np.ndarray) -> list[list[int]]:
    if indices.size == 0:
        return []
    idx = np.asarray(indices, dtype=int)
    starts = [int(idx[0])]
    ends: list[int] = []
    for prev, nxt in zip(idx[:-1], idx[1:]):
        if int(nxt) != int(prev) + 1:
            ends.append(int(prev) + 1)
            starts.append(int(nxt))
    ends.append(int(idx[-1]) + 1)
    return [[start, end] for start, end in zip(starts, ends)]

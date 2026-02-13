from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import numpy as np
from numpy.typing import NDArray


DEFAULT_MONEYNESS_BUCKETS: NDArray[np.float64] = np.array(
    [0.7, 0.8, 0.9, 0.95, 1.0, 1.05, 1.1, 1.2, 1.3], dtype=np.float64
)
DEFAULT_TENOR_BUCKETS: NDArray[np.float64] = np.array(
    [7 / 365, 14 / 365, 30 / 365, 60 / 365, 90 / 365, 180 / 365, 365 / 365], dtype=np.float64
)


@dataclass(frozen=True)
class InterpolationExtrapolationPolicy:
    interpolation_method: str = "inverse_distance"
    extrapolation_method: str = "nearest"
    max_neighbors: int = 4
    min_points_for_interpolation: int = 3
    distance_epsilon: float = 1e-8


@dataclass(frozen=True)
class CanonicalSurfaceMetadata:
    moneyness_buckets: NDArray[np.float64]
    tenor_buckets: NDArray[np.float64]
    policy: InterpolationExtrapolationPolicy
    mean: float
    std: float
    source_shape: tuple[int, int]


@dataclass(frozen=True)
class CanonicalVolSurface:
    raw_tensor: NDArray[np.float64]
    normalized_tensor: NDArray[np.float64]
    observed_mask: NDArray[np.bool_]
    interpolated_mask: NDArray[np.bool_]
    extrapolated_mask: NDArray[np.bool_]
    missing_mask: NDArray[np.bool_]
    metadata: CanonicalSurfaceMetadata


def build_canonical_vol_surface(
    raw_moneyness: NDArray[np.float64],
    raw_tenor: NDArray[np.float64],
    raw_implied_vol: NDArray[np.float64],
    *,
    moneyness_buckets: NDArray[np.float64] | None = None,
    tenor_buckets: NDArray[np.float64] | None = None,
    policy: InterpolationExtrapolationPolicy | None = None,
) -> CanonicalVolSurface:
    """Build a canonical options surface with fixed buckets and fill policies.

    Inputs are expected as 2D tensors of shape ``(n_dates, n_quotes)``.
    Missing quotes should be represented by NaN.
    """

    _validate_raw_inputs(raw_moneyness=raw_moneyness, raw_tenor=raw_tenor, raw_implied_vol=raw_implied_vol)

    m_buckets = np.asarray(moneyness_buckets if moneyness_buckets is not None else DEFAULT_MONEYNESS_BUCKETS)
    t_buckets = np.asarray(tenor_buckets if tenor_buckets is not None else DEFAULT_TENOR_BUCKETS)
    policy = policy or InterpolationExtrapolationPolicy()

    n_dates, _ = raw_implied_vol.shape
    grid_shape = (n_dates, m_buckets.size, t_buckets.size)

    raw_tensor = np.full(grid_shape, np.nan, dtype=np.float64)
    observed_mask = np.zeros(grid_shape, dtype=bool)
    interpolated_mask = np.zeros(grid_shape, dtype=bool)
    extrapolated_mask = np.zeros(grid_shape, dtype=bool)

    bucket_indices_m = _nearest_bucket_indices(raw_moneyness, m_buckets)
    bucket_indices_t = _nearest_bucket_indices(raw_tenor, t_buckets)

    for date_idx in range(n_dates):
        valid_quote_mask = (
            np.isfinite(raw_moneyness[date_idx])
            & np.isfinite(raw_tenor[date_idx])
            & np.isfinite(raw_implied_vol[date_idx])
        )

        if not np.any(valid_quote_mask):
            continue

        m_idx = bucket_indices_m[date_idx, valid_quote_mask]
        t_idx = bucket_indices_t[date_idx, valid_quote_mask]
        iv = raw_implied_vol[date_idx, valid_quote_mask]

        for bucket_m, bucket_t in np.unique(np.stack([m_idx, t_idx], axis=1), axis=0):
            cell_mask = (m_idx == bucket_m) & (t_idx == bucket_t)
            raw_tensor[date_idx, bucket_m, bucket_t] = float(np.nanmean(iv[cell_mask]))
            observed_mask[date_idx, bucket_m, bucket_t] = True

        _fill_surface_holes(
            date_idx=date_idx,
            raw_tensor=raw_tensor,
            observed_mask=observed_mask,
            interpolated_mask=interpolated_mask,
            extrapolated_mask=extrapolated_mask,
            moneyness_buckets=m_buckets,
            tenor_buckets=t_buckets,
            policy=policy,
        )

    missing_mask = np.isnan(raw_tensor)
    normalized_tensor, mean, std = _zscore(raw_tensor)

    metadata = CanonicalSurfaceMetadata(
        moneyness_buckets=m_buckets.astype(np.float64),
        tenor_buckets=t_buckets.astype(np.float64),
        policy=policy,
        mean=float(mean),
        std=float(std),
        source_shape=(int(raw_implied_vol.shape[0]), int(raw_implied_vol.shape[1])),
    )

    return CanonicalVolSurface(
        raw_tensor=raw_tensor,
        normalized_tensor=normalized_tensor,
        observed_mask=observed_mask,
        interpolated_mask=interpolated_mask,
        extrapolated_mask=extrapolated_mask,
        missing_mask=missing_mask,
        metadata=metadata,
    )


def _fill_surface_holes(
    *,
    date_idx: int,
    raw_tensor: NDArray[np.float64],
    observed_mask: NDArray[np.bool_],
    interpolated_mask: NDArray[np.bool_],
    extrapolated_mask: NDArray[np.bool_],
    moneyness_buckets: NDArray[np.float64],
    tenor_buckets: NDArray[np.float64],
    policy: InterpolationExtrapolationPolicy,
) -> None:
    observed_cells = observed_mask[date_idx]
    if observed_cells.sum() < 1:
        return

    m_grid, t_grid = np.meshgrid(moneyness_buckets, tenor_buckets, indexing="ij")
    all_points = np.stack([m_grid.ravel(), t_grid.ravel()], axis=1)

    obs_indices = np.argwhere(observed_cells)
    obs_points = np.stack(
        [moneyness_buckets[obs_indices[:, 0]], tenor_buckets[obs_indices[:, 1]]],
        axis=1,
    )
    obs_values = raw_tensor[date_idx, obs_indices[:, 0], obs_indices[:, 1]]

    m_min, m_max = obs_points[:, 0].min(), obs_points[:, 0].max()
    t_min, t_max = obs_points[:, 1].min(), obs_points[:, 1].max()

    for target in all_points:
        m_target, t_target = target
        m_i = int(np.argmin(np.abs(moneyness_buckets - m_target)))
        t_i = int(np.argmin(np.abs(tenor_buckets - t_target)))

        if observed_cells[m_i, t_i]:
            continue

        outside_hull = m_target < m_min or m_target > m_max or t_target < t_min or t_target > t_max
        distances = np.sqrt(np.sum((obs_points - target) ** 2, axis=1))

        if not outside_hull and obs_points.shape[0] >= policy.min_points_for_interpolation:
            k = min(policy.max_neighbors, obs_points.shape[0])
            nn_idx = np.argsort(distances)[:k]
            d = distances[nn_idx]
            w = 1.0 / np.maximum(d, policy.distance_epsilon)
            value = float(np.sum(w * obs_values[nn_idx]) / np.sum(w))
            raw_tensor[date_idx, m_i, t_i] = value
            interpolated_mask[date_idx, m_i, t_i] = True
            continue

        nn_idx = int(np.argmin(distances))
        raw_tensor[date_idx, m_i, t_i] = float(obs_values[nn_idx])
        extrapolated_mask[date_idx, m_i, t_i] = bool(outside_hull)
        interpolated_mask[date_idx, m_i, t_i] = not bool(outside_hull)


def _nearest_bucket_indices(
    values: NDArray[np.float64],
    buckets: NDArray[np.float64],
) -> NDArray[np.int64]:
    diffs = np.abs(values[..., None] - buckets[None, None, :])
    return np.argmin(diffs, axis=-1)


def _zscore(values: NDArray[np.float64]) -> tuple[NDArray[np.float64], float, float]:
    finite = np.isfinite(values)
    if not finite.any():
        return np.full_like(values, np.nan), 0.0, 1.0

    mean = float(np.nanmean(values))
    std = float(np.nanstd(values))
    if std <= 0:
        std = 1.0

    normalized = (values - mean) / std
    normalized[~finite] = np.nan
    return normalized, mean, std


def _validate_raw_inputs(
    *,
    raw_moneyness: NDArray[np.float64],
    raw_tenor: NDArray[np.float64],
    raw_implied_vol: NDArray[np.float64],
) -> None:
    shapes = {raw_moneyness.shape, raw_tenor.shape, raw_implied_vol.shape}
    if len(shapes) != 1:
        raise ValueError("raw_moneyness, raw_tenor and raw_implied_vol must share shape")

    if raw_implied_vol.ndim != 2:
        raise ValueError("raw inputs must be 2D tensors of shape (n_dates, n_quotes)")

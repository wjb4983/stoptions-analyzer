from __future__ import annotations

import numpy as np

from src.modeling_nextgen.features.vol_surface import (
    InterpolationExtrapolationPolicy,
    build_canonical_vol_surface,
)


def test_build_canonical_vol_surface_outputs_expected_shapes_and_metadata() -> None:
    raw_moneyness = np.array([[0.9, 1.0, 1.1, np.nan], [0.95, 1.0, 1.2, 1.3]], dtype=np.float64)
    raw_tenor = np.array([[30 / 365, 90 / 365, 180 / 365, np.nan], [14 / 365, 30 / 365, 90 / 365, 365 / 365]], dtype=np.float64)
    raw_iv = np.array([[0.28, 0.24, 0.22, np.nan], [0.3, 0.25, 0.23, 0.21]], dtype=np.float64)

    surface = build_canonical_vol_surface(raw_moneyness, raw_tenor, raw_iv)

    assert surface.raw_tensor.shape == (2, 9, 7)
    assert surface.normalized_tensor.shape == (2, 9, 7)
    assert surface.observed_mask.shape == (2, 9, 7)
    assert surface.missing_mask.shape == (2, 9, 7)

    assert np.isfinite(surface.metadata.mean)
    assert surface.metadata.std > 0
    assert surface.metadata.source_shape == (2, 4)


def test_masks_mark_observed_interpolated_and_extrapolated_cells() -> None:
    raw_moneyness = np.array([[0.9, 1.0, 1.1]], dtype=np.float64)
    raw_tenor = np.array([[30 / 365, 90 / 365, 180 / 365]], dtype=np.float64)
    raw_iv = np.array([[0.29, 0.24, 0.22]], dtype=np.float64)

    policy = InterpolationExtrapolationPolicy(max_neighbors=3, min_points_for_interpolation=3)
    surface = build_canonical_vol_surface(raw_moneyness, raw_tenor, raw_iv, policy=policy)

    assert surface.observed_mask.sum() == 3
    assert surface.interpolated_mask.sum() > 0
    assert surface.extrapolated_mask.sum() > 0
    assert not surface.missing_mask.any()

    finite = np.isfinite(surface.raw_tensor)
    assert finite.all()
    assert np.isfinite(surface.normalized_tensor[finite]).all()

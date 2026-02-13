"""Analysis package for cross-sectional and time-series strategies."""

from .diagnostics import compute_signal_diagnostics, validate_signal_diagnostics
from .cross_asset import (
    build_release_aware_macro_features,
    join_cross_asset_indicators,
    compare_conditioned_models,
    compute_cross_market_transmission,
)
from .options import (
    aggregate_option_exposures,
    compute_options_feature_pipeline,
    summarize_lifecycle_events,
)
from .prompt_pack import build_prompt_pack_markdown, write_prompt_pack

__all__ = [
    "aggregate_option_exposures",
    "summarize_lifecycle_events",
    "compute_options_feature_pipeline",
    "compute_signal_diagnostics",
    "validate_signal_diagnostics",
    "build_prompt_pack_markdown",
    "write_prompt_pack",
    "build_release_aware_macro_features",
    "join_cross_asset_indicators",
    "compare_conditioned_models",
    "compute_cross_market_transmission",
]

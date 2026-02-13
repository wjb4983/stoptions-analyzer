"""Analysis package for cross-sectional and time-series strategies."""

from .diagnostics import compute_signal_diagnostics, validate_signal_diagnostics
from .options import aggregate_option_exposures, summarize_lifecycle_events

__all__ = [
    "aggregate_option_exposures",
    "summarize_lifecycle_events",
    "compute_signal_diagnostics",
    "validate_signal_diagnostics",
]

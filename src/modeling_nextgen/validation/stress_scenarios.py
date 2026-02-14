from __future__ import annotations

from dataclasses import dataclass

import numpy as np


@dataclass(frozen=True)
class StressTemplateConfig:
    jump_probability: float = 0.08
    jump_scale: float = 2.5
    volatility_regime_threshold: float = 0.75
    low_vol_scale: float = 0.6
    high_vol_scale: float = 1.8
    spread_multiplier: float = 1.7
    impact_multiplier: float = 2.2


def _as_array(features: dict[str, np.ndarray]) -> dict[str, np.ndarray]:
    return {name: np.asarray(values, dtype=float).copy() for name, values in features.items()}


def build_stress_template_scenarios(
    *,
    features: dict[str, np.ndarray],
    rng: np.random.Generator,
    config: StressTemplateConfig | None = None,
) -> dict[str, dict[str, np.ndarray]]:
    """Build standardized stress template scenarios for robustness validation."""

    cfg = config or StressTemplateConfig()
    base = _as_array(features)

    jump_diffusion: dict[str, np.ndarray] = {}
    volatility_regime_jump: dict[str, np.ndarray] = {}
    spread_impact_widening: dict[str, np.ndarray] = {}

    for name, values in base.items():
        sigma = float(np.nanstd(values))
        sigma = sigma if sigma > 0.0 else 1.0
        mean = float(np.nanmean(values)) if np.isfinite(np.nanmean(values)) else 0.0

        jump_mask = rng.random(size=values.shape) < cfg.jump_probability
        jump_sizes = rng.normal(0.0, sigma * cfg.jump_scale, size=values.shape)
        jump_diffusion[name] = values + jump_mask.astype(float) * jump_sizes

        centered = values - mean
        vol_cut = float(np.quantile(np.abs(centered), cfg.volatility_regime_threshold))
        high_vol_mask = np.abs(centered) >= vol_cut
        vol_scaled = centered.copy()
        vol_scaled[~high_vol_mask] *= cfg.low_vol_scale
        vol_scaled[high_vol_mask] *= cfg.high_vol_scale
        volatility_regime_jump[name] = mean + vol_scaled

        widened = values.copy()
        lname = name.lower()
        if "spread" in lname or "bid_ask" in lname:
            widened = mean + (widened - mean) * cfg.spread_multiplier
        elif "impact" in lname or "slippage" in lname or "cost" in lname:
            widened = mean + (widened - mean) * cfg.impact_multiplier
        else:
            widened = mean + (widened - mean) * ((cfg.spread_multiplier + cfg.impact_multiplier) / 2.0)
        spread_impact_widening[name] = widened

    return {
        "jump_diffusion_shocks": jump_diffusion,
        "volatility_regime_jumps": volatility_regime_jump,
        "spread_impact_widening": spread_impact_widening,
    }

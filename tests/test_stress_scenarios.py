from __future__ import annotations

import numpy as np

from modeling_nextgen.validation.stress_scenarios import build_stress_template_scenarios


def test_build_stress_template_scenarios_returns_expected_templates() -> None:
    rng = np.random.default_rng(42)
    features = {
        "returns_1m": np.linspace(-1.0, 1.0, 32),
        "bid_ask_spread": np.linspace(0.01, 0.03, 32),
        "market_impact": np.linspace(0.1, 0.2, 32),
    }

    scenarios = build_stress_template_scenarios(features=features, rng=rng)

    assert set(scenarios.keys()) == {
        "jump_diffusion_shocks",
        "volatility_regime_jumps",
        "spread_impact_widening",
    }
    for scenario_payload in scenarios.values():
        assert set(scenario_payload.keys()) == set(features.keys())
        for key, value in scenario_payload.items():
            assert value.shape == features[key].shape

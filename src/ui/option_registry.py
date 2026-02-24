from __future__ import annotations

from typing import Iterable

# Mirrors supported values in src/backtesting/cache_runner.py.
ENTRY_SIGNALS: tuple[str, ...] = (
    "ts_momentum",
    "ma_trend",
    "breakout",
    "mean_reversion",
    "vol_carry",
    "trend_strength",
    "seasonality_event",
    "vrp_harvest",
    "cheap_vol_long",
)
EXIT_SIGNALS: tuple[str, ...] = ("none", "momentum_flip", "trailing_stop", "max_hold")
EXECUTION_MODELS: tuple[str, ...] = (
    "bps",
    "spread",
    "participation",
    "square_root",
    "latency_drift",
    "modular",
    "volatility_scaled",
)
BENCHMARK_NAMES: tuple[str, ...] = ("buy_hold", "equal_weight_momentum", "volatility_parity")
OPTIMIZER_SAMPLERS: tuple[str, ...] = ("tpe", "random", "bayesian", "cma-es", "grid")


_OPTION_ALIASES: dict[str, dict[str, str]] = {
    "optimizer sampler": {
        "cma_es": "cma-es",
        "cma": "cma-es",
    },
}


def _normalize(value: str) -> str:
    return str(value).strip().lower().replace(" ", "_")


def normalize_supported_option(value: str, supported: tuple[str, ...], *, field_name: str | None = None) -> str | None:
    key = _normalize(value)
    supported_by_key = {_normalize(item): item for item in supported}
    if key in supported_by_key:
        return supported_by_key[key]
    aliases = _OPTION_ALIASES.get((field_name or "").strip().lower(), {})
    mapped = aliases.get(key)
    if mapped in supported:
        return mapped
    return None


def validate_option_values(
    values: Iterable[str],
    *,
    supported: tuple[str, ...],
    field_name: str,
) -> tuple[list[str], list[str], dict[str, str]]:
    valid: list[str] = []
    stale: list[str] = []
    migrations: dict[str, str] = {}
    for raw in values:
        original = str(raw).strip()
        if not original:
            continue
        normalized = normalize_supported_option(original, supported, field_name=field_name)
        if normalized is None:
            stale.append(original)
            continue
        if _normalize(original) != _normalize(normalized):
            migrations[original] = normalized
        if normalized not in valid:
            valid.append(normalized)
    return valid, stale, migrations


def migration_hint_text(*, stale: list[str], migrations: dict[str, str], supported: tuple[str, ...], field_name: str) -> str:
    parts: list[str] = []
    if stale:
        parts.append(f"Unsupported {field_name}: {', '.join(stale)}")
    if migrations:
        parts.append(
            "Migration suggestions: "
            + ", ".join(f"{old} → {new}" for old, new in migrations.items())
        )
    if stale or migrations:
        parts.append(f"Supported values: {', '.join(supported)}")
    return " | ".join(parts)

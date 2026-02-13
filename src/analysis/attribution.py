from __future__ import annotations

from dataclasses import dataclass

import numpy as np


@dataclass(frozen=True)
class AttributionPayload:
    time_series: list[dict[str, object]]
    summary: list[dict[str, object]]


def build_attribution_payload(
    *,
    timestamps: np.ndarray,
    prices: np.ndarray,
    positions: np.ndarray,
    slippage_drag: np.ndarray,
    fee_drag: np.ndarray,
    borrow_drag: np.ndarray,
) -> AttributionPayload:
    """Build Brinson-style and factor-model attribution variants.

    Produces variants for both cross-sectional and time-series strategy styles.
    """

    close = np.asarray(prices, dtype=float)
    pos = np.asarray(positions, dtype=float)
    if close.ndim != 2 or pos.ndim != 2 or close.shape != pos.shape:
        raise ValueError("prices and positions must be 2D arrays with matching shapes")

    n_periods, n_assets = close.shape
    weights = np.full((n_periods, n_assets), 1.0 / max(1, n_assets), dtype=float)

    asset_returns = np.zeros_like(close)
    asset_returns[1:] = np.divide(close[1:], close[:-1], out=np.ones_like(close[1:]), where=close[:-1] != 0.0) - 1.0
    gross_alpha = np.sum(pos * asset_returns * weights, axis=1)

    slippage = _to_len(slippage_drag, n_periods)
    fees = _to_len(fee_drag, n_periods)
    borrow = _to_len(borrow_drag, n_periods)
    cost_drag = slippage + fees + borrow
    net_alpha = gross_alpha - cost_drag

    ts_rows: list[dict[str, object]] = []
    ts_rows.extend(
        _build_brinson_rows(
            variant="brinson_cross_sectional",
            strategy_style="cross_sectional",
            timestamps=timestamps,
            asset_returns=asset_returns,
            effective_weights=pos * weights,
            benchmark_weights=np.full_like(weights, 1.0 / max(1, n_assets), dtype=float),
            gross_alpha=gross_alpha,
            net_alpha=net_alpha,
            cost_drag=cost_drag,
            slippage=slippage,
            fees=fees,
            borrow=borrow,
        )
    )
    ts_rows.extend(
        _build_brinson_rows(
            variant="brinson_time_series",
            strategy_style="time_series",
            timestamps=timestamps,
            asset_returns=asset_returns,
            effective_weights=pos * weights,
            benchmark_weights=np.zeros_like(weights),
            gross_alpha=gross_alpha,
            net_alpha=net_alpha,
            cost_drag=cost_drag,
            slippage=slippage,
            fees=fees,
            borrow=borrow,
        )
    )
    ts_rows.extend(
        _build_factor_rows(
            variant="factor_cross_sectional",
            strategy_style="cross_sectional",
            timestamps=timestamps,
            asset_returns=asset_returns,
            effective_weights=pos * weights,
            benchmark_weights=np.full_like(weights, 1.0 / max(1, n_assets), dtype=float),
            gross_alpha=gross_alpha,
            net_alpha=net_alpha,
            cost_drag=cost_drag,
            slippage=slippage,
            fees=fees,
            borrow=borrow,
        )
    )
    ts_rows.extend(
        _build_factor_rows(
            variant="factor_time_series",
            strategy_style="time_series",
            timestamps=timestamps,
            asset_returns=asset_returns,
            effective_weights=pos * weights,
            benchmark_weights=np.zeros_like(weights),
            gross_alpha=gross_alpha,
            net_alpha=net_alpha,
            cost_drag=cost_drag,
            slippage=slippage,
            fees=fees,
            borrow=borrow,
        )
    )

    summary = _build_summary(ts_rows)
    return AttributionPayload(time_series=ts_rows, summary=summary)


def _to_len(values: np.ndarray, n: int) -> np.ndarray:
    arr = np.asarray(values, dtype=float).reshape(-1)
    if arr.size == n:
        return arr
    out = np.zeros(n, dtype=float)
    out[: min(n, arr.size)] = arr[: min(n, arr.size)]
    return out


def _build_brinson_rows(
    *,
    variant: str,
    strategy_style: str,
    timestamps: np.ndarray,
    asset_returns: np.ndarray,
    effective_weights: np.ndarray,
    benchmark_weights: np.ndarray,
    gross_alpha: np.ndarray,
    net_alpha: np.ndarray,
    cost_drag: np.ndarray,
    slippage: np.ndarray,
    fees: np.ndarray,
    borrow: np.ndarray,
) -> list[dict[str, object]]:
    active = effective_weights - benchmark_weights
    cross_mean = np.mean(asset_returns, axis=1, keepdims=True)
    allocation = np.sum(active * cross_mean, axis=1)
    selection = np.sum(active * (asset_returns - cross_mean), axis=1)
    explained = allocation + selection
    residual = gross_alpha - explained
    return _rows(
        variant=variant,
        strategy_style=strategy_style,
        timestamps=timestamps,
        gross_alpha=gross_alpha,
        explained_component=explained,
        residual=residual,
        net_alpha=net_alpha,
        cost_drag=cost_drag,
        slippage=slippage,
        fees=fees,
        borrow=borrow,
    )


def _build_factor_rows(
    *,
    variant: str,
    strategy_style: str,
    timestamps: np.ndarray,
    asset_returns: np.ndarray,
    effective_weights: np.ndarray,
    benchmark_weights: np.ndarray,
    gross_alpha: np.ndarray,
    net_alpha: np.ndarray,
    cost_drag: np.ndarray,
    slippage: np.ndarray,
    fees: np.ndarray,
    borrow: np.ndarray,
) -> list[dict[str, object]]:
    market_factor = np.mean(asset_returns, axis=1, keepdims=True)
    factor_implied = np.repeat(market_factor, asset_returns.shape[1], axis=1)
    active = effective_weights - benchmark_weights
    explained = np.sum(active * factor_implied, axis=1)
    residual = gross_alpha - explained
    return _rows(
        variant=variant,
        strategy_style=strategy_style,
        timestamps=timestamps,
        gross_alpha=gross_alpha,
        explained_component=explained,
        residual=residual,
        net_alpha=net_alpha,
        cost_drag=cost_drag,
        slippage=slippage,
        fees=fees,
        borrow=borrow,
    )


def _rows(
    *,
    variant: str,
    strategy_style: str,
    timestamps: np.ndarray,
    gross_alpha: np.ndarray,
    explained_component: np.ndarray,
    residual: np.ndarray,
    net_alpha: np.ndarray,
    cost_drag: np.ndarray,
    slippage: np.ndarray,
    fees: np.ndarray,
    borrow: np.ndarray,
) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for idx in range(gross_alpha.size):
        rows.append(
            {
                "timestamp": str(timestamps[idx]),
                "variant": variant,
                "strategy_style": strategy_style,
                "gross_alpha": float(gross_alpha[idx]),
                "explained_component": float(explained_component[idx]),
                "residual_unexplained": float(residual[idx]),
                "cost_drag": float(cost_drag[idx]),
                "slippage_drag": float(slippage[idx]),
                "fee_drag": float(fees[idx]),
                "borrow_drag": float(borrow[idx]),
                "net_alpha": float(net_alpha[idx]),
            }
        )
    return rows


def _build_summary(rows: list[dict[str, object]]) -> list[dict[str, object]]:
    buckets: dict[str, list[dict[str, object]]] = {}
    for row in rows:
        buckets.setdefault(str(row["variant"]), []).append(row)
    out: list[dict[str, object]] = []
    for variant, variant_rows in buckets.items():
        count = max(1, len(variant_rows))
        out.append(
            {
                "variant": variant,
                "strategy_style": variant_rows[0]["strategy_style"],
                "bars": len(variant_rows),
                "gross_alpha_total": float(sum(float(r["gross_alpha"]) for r in variant_rows)),
                "cost_drag_total": float(sum(float(r["cost_drag"]) for r in variant_rows)),
                "slippage_drag_total": float(sum(float(r["slippage_drag"]) for r in variant_rows)),
                "borrow_drag_total": float(sum(float(r["borrow_drag"]) for r in variant_rows)),
                "residual_unexplained_total": float(sum(float(r["residual_unexplained"]) for r in variant_rows)),
                "net_alpha_total": float(sum(float(r["net_alpha"]) for r in variant_rows)),
                "gross_alpha_mean": float(sum(float(r["gross_alpha"]) for r in variant_rows) / count),
            }
        )
    return sorted(out, key=lambda row: str(row["variant"]))


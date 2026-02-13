# Backtesting MVP Design

## 1) Universe input format

The strategy universe is provided as a tabular dataset with one row per tradable symbol.

**Required fields**

- `symbol` *(string)*: instrument identifier (e.g., `AAPL`, `SPY_2026C450`).
- `asset_type` *(string)*: `equity`, `etf`, `option`, `future`, etc.
- `exchange` *(string, optional)*: listing venue.
- `currency` *(string, default `USD`)*: quote currency for PnL aggregation.
- `multiplier` *(float, default `1.0`)*: contract/share multiplier for sizing and PnL.
- `active_from` *(timestamp, optional)* and `active_to` *(timestamp, optional)*.

**Example (JSON Lines)**

```json
{"symbol":"AAPL","asset_type":"equity","currency":"USD","multiplier":1.0}
{"symbol":"MSFT","asset_type":"equity","currency":"USD","multiplier":1.0}
```

## 2) Price input schema (OHLCV + timestamps)

Bar data is long-form, one row per `(timestamp, symbol)`.

**Required fields**

- `timestamp` *(RFC3339/ISO-8601 string or epoch)*: bar end time in UTC.
- `symbol` *(string)*.
- `open` *(float)*.
- `high` *(float)*.
- `low` *(float)*.
- `close` *(float)*.
- `volume` *(float/int)*.

**Optional fields**

- `vwap` *(float)*.
- `open_interest` *(float/int, derivatives)*.
- `corporate_action_factor` *(float)*.

**Invariants**

- Bars must be time-ordered per symbol.
- No duplicate `(timestamp, symbol)`.
- `low <= min(open, close) <= max(open, close) <= high`.

## 3) Signal output schema

Signal output is normalized to a target position/exposure for each `(timestamp, symbol)` generated **at bar close**.

**Fields**

- `timestamp` *(bar close timestamp)*.
- `symbol` *(string)*.
- `signal` *(float)*: target position in units or normalized exposure (implementation-specific).
- `confidence` *(float, optional, 0..1)*.
- `metadata` *(object, optional)*.

**Execution convention**

Signals generated at close of bar `t` are executed at **open of bar `t+1`**.

## 4) Position sizing rules

MVP supports deterministic sizing rules:

1. **Fixed units**: buy/sell `N` units per signal change.
2. **Target exposure**: map `signal` to target weight, then convert to units using portfolio equity and instrument price.
3. **Risk cap**: enforce max gross exposure and max per-symbol exposure.
4. **Rounding**: round to instrument lot size (`1` share by default).
5. **No leverage by default** unless explicitly enabled.

## 5) Trade execution assumptions (MVP vs implemented)

The original MVP spec was deliberately simple. The current codebase now supports a richer execution stack.

### 5.1 MVP assumptions (design target)

- Primary fill rule: **next-open fill**.
- Fill price starts from bar `open` and is adjusted by a simple slippage model.
- Fees are charged per fill and deducted from cash/PnL.
- Partial fills are out of scope (assume full fill).
- Market impact is represented only through pluggable slippage (no full order book simulation).

### 5.2 Currently implemented behavior

- **Next-open semantics are still the baseline** in both the reference event loop and execution contracts.
- **Partial fills are supported** via per-bar participation caps with residual quantity carried forward bar-to-bar.
- **Latency inputs are supported** as both `latency_bars` and `latency_ms` execution context fields.
- **Queue-rank effects are modeled** through `queue_rank_proxy` (0..1), which increases spread/impact/drift penalties in multiple slippage components.
- **Lifecycle-aware event-driven bridging is available** through the vectorized adapter that maps requested trades to `submit`/`amend`/`cancel`/`fill`/`expire` events.

### 5.3 MVP assumptions vs implementation status

| Topic | MVP assumption | Current implementation |
|---|---|---|
| Fill timing | Next bar open | Still next-open baseline in execution/event loop. |
| Fill completeness | Full fill on first eligible bar | `PartialFillModel` can cap participation and carry residuals. |
| Latency | Ignored | `ExecutionContext` includes `latency_bars` and `latency_ms`; drift models can penalize delay. |
| Queue position | Ignored | `queue_rank_proxy` adjusts spread, impact, and latency-drift terms. |
| Event-driven integration | Not explicitly modeled | `VectorizedExecutionAdapter` emits deterministic lifecycle events and replays them. |

## 6) Model risk / assumptions map

The table below documents where model risk enters and the concrete implementation touchpoints.

| Assumption / risk | Why it matters | Execution class / adapter path |
|---|---|---|
| Next-open execution | Can understate intrabar gap/impact in fast markets. | `src/backtesting/event_driven.py` (`EventDrivenBacktester` and `EventDrivenRunner` contract). |
| Participation-capped partial fills | Realized exposure may lag target; residual inventory path-dependent. | `src/backtesting/execution.py` (`PartialFillModel`, `FillEvent`). |
| Queue rank proxy instead of full LOB state | Proxy may mis-rank true queue priority and miss venue microstructure effects. | `src/backtesting/execution.py` (`ExecutionContext.queue_rank_proxy`, `SpreadSlippage`, `ParticipationImpactSlippage`, `LatencyQueueDriftSlippage`, `VolatilityScaledSlippage`). |
| Latency as bars/ms scalars | Fixed-latency approximation omits network jitter and exchange-specific matching delay tails. | `src/backtesting/execution.py` (`ExecutionContext.latency_bars`, `ExecutionContext.latency_ms`, `LatencyQueueDriftSlippage`). |
| Deterministic lifecycle adapter | Reproducibility is high, but stochastic queue/outage behavior is not native unless injected in inputs. | `src/backtesting/event_driven.py` (`VectorizedExecutionAdapter`, `OrderLifecycleBook`, `replay_lifecycle`). |

## 7) Required MVP metrics

- **CAGR**: annualized growth rate from equity curve.
- **Sharpe ratio**: annualized mean return divided by annualized volatility.
- **Max drawdown**: minimum peak-to-trough drawdown.
- **Turnover**: sum of absolute traded notional (or units) normalized by average equity.
- **Win rate**: fraction of closed trades with positive PnL.

These are computed from backtest outputs and reported in a single metrics object.

## 8) Execution settings examples (deterministic vs stochastic)

Both examples preserve next-open semantics but differ in how execution inputs are generated.

### Deterministic configuration (reproducible baseline)

Use fixed settings for all bars/assets:

- `latency_bars = 1`
- `latency_ms = 0`
- `queue_rank_proxy = 0.25`
- `max_participation_per_bar = 0.10`
- Slippage stack example: `CompositeSlippage([SpreadSlippage(2.0), ParticipationImpactSlippage(base_bps=0.5, impact_coefficient_bps=12.0), LatencyQueueDriftSlippage(drift_bps_per_bar=1.0, queue_drift_bps=1.5)])`

This yields deterministic fills/lifecycle output for fixed inputs.

### Stochastic configuration (scenario/Monte Carlo style)

Use seeded random draws for execution-state fields while keeping strategy logic fixed:

- `latency_ms ~ LogNormal(mean=6.5, sigma=0.35)` clipped to operational bounds.
- `queue_rank_proxy ~ Beta(2, 3)` per bar/asset.
- `max_participation_per_bar` sampled by liquidity regime (e.g., low/medium/high volume buckets).

Typical workflow:

1. Set a random seed per scenario for reproducibility.
2. Generate `queue_rank_proxy`, `latency_ms`, and participation-cap matrices.
3. Pass those matrices through `VectorizedExecutionAdapter.execute(...)` and slippage models consuming `ExecutionContext`.
4. Aggregate distributional outcomes (PnL, drawdown, turnover, residual inventory days).

## 9) Next after MVP backlog

1. **Walk-forward optimization**
   - Add rolling train/validation/test windows for momentum and cost hyperparameters.
   - Store each fold's parameter choice and out-of-sample metrics for auditability.

2. **Regime filters**
   - Gate exposure with simple market-state filters (trend/volatility/liquidity regimes).
   - Track filter hit-rate and incremental alpha vs. always-on baseline.

3. **Multi-asset risk parity**
   - Extend sizing to equal-risk-contribution allocations across symbols.
   - Add leverage and concentration controls at portfolio and asset-class levels.

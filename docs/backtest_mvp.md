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

## 5) Trade execution assumptions

- Primary fill rule: **next-open fill**.
- Fill price starts from bar `open` and is adjusted by slippage model.
- Fees are charged per fill and deducted from cash/PnL.
- Partial fills are out of scope for MVP (assume full fill).
- Market impact is represented only through pluggable slippage (no order book simulation).

## 6) Required MVP metrics

- **CAGR**: annualized growth rate from equity curve.
- **Sharpe ratio**: annualized mean return divided by annualized volatility.
- **Max drawdown**: minimum peak-to-trough drawdown.
- **Turnover**: sum of absolute traded notional (or units) normalized by average equity.
- **Win rate**: fraction of closed trades with positive PnL.

These are computed from backtest outputs and reported in a single metrics object.

## 7) Next after MVP backlog

1. **Walk-forward optimization**
   - Add rolling train/validation/test windows for momentum and cost hyperparameters.
   - Store each fold's parameter choice and out-of-sample metrics for auditability.

2. **Regime filters**
   - Gate exposure with simple market-state filters (trend/volatility/liquidity regimes).
   - Track filter hit-rate and incremental alpha vs. always-on baseline.

3. **Multi-asset risk parity**
   - Extend sizing to equal-risk-contribution allocations across symbols.
   - Add leverage and concentration controls at portfolio and asset-class levels.

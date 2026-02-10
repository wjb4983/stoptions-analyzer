from __future__ import annotations

import json
import random
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime
from pathlib import Path

import numpy as np

from config import BACKTEST_CACHE_DIR, BACKTEST_OUTPUT_DIR
from data_access.api_client import MassiveApiClient
from data_access.cache import _safe_ticker_name

from data_access.engine_loader import EngineArrayBundle, load_canonical_price_arrays
from utils.parsing import build_npz_payload, chunk_results_by_year


def run_backtest_cache(
    tickers: list[str],
    start_date: date,
    end_date: date,
    cache_root: Path,
    api_key: str,
) -> str:
    api_client = MassiveApiClient(api_key)
    cache_root.mkdir(parents=True, exist_ok=True)
    BACKTEST_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    def _process_ticker(ticker: str) -> str:
        safe_ticker = _safe_ticker_name(ticker)
        ticker_dir = cache_root / safe_ticker / "1m"
        ticker_dir.mkdir(parents=True, exist_ok=True)
        index_path = ticker_dir / "index.json"
        expected_years = list(range(start_date.year, end_date.year + 1))
        try:
            cache_ready = False
            if index_path.exists():
                index_data = json.loads(index_path.read_text())
                years = index_data.get("years", [])
                cache_ready = (
                    index_data.get("full_range") is True
                    and set(expected_years).issubset(set(years))
                )
            if cache_ready:
                sample_text = f"{ticker}: cached data ready"
                sample_year = random.choice(expected_years)
                sample_path = ticker_dir / f"{safe_ticker}_1m_{sample_year}.npz"
                if sample_path.exists():
                    with np.load(sample_path, mmap_mode="r") as data:
                        if data["t"].size > 0:
                            idx = random.randrange(data["t"].size)
                            sample_text = (
                                f"{ticker}: sample close={data['c'][idx]} "
                                f"timestamp={int(data['t'][idx])}"
                            )
                return sample_text
            legacy_path = (
                cache_root
                / f"{safe_ticker}_1m_{start_date.isoformat()}_{end_date.isoformat()}.json"
            )
            if not legacy_path.exists():
                legacy_path = (
                    BACKTEST_CACHE_DIR
                    / f"{safe_ticker}_1m_{start_date.isoformat()}_{end_date.isoformat()}.json"
                )
            if legacy_path.exists():
                results = json.loads(legacy_path.read_text()).get("results", [])
            else:
                results = api_client.fetch_aggregates_range(
                    ticker, start_date, end_date, minutes_per_bar=1
                )
            buckets = chunk_results_by_year(results)
            for year, entries in buckets.items():
                payload = build_npz_payload(entries)
                np.savez_compressed(
                    ticker_dir / f"{safe_ticker}_1m_{year}.npz", **payload
                )
            index_payload = {
                "ticker": ticker,
                "start_date": start_date.isoformat(),
                "end_date": end_date.isoformat(),
                "full_range": True,
                "fetched_at": datetime.now().isoformat(),
                "years": sorted(buckets.keys()),
            }
            index_path.write_text(json.dumps(index_payload, indent=2))
            if results:
                sample = random.choice(results)
                return (
                    f"{ticker}: sample close={sample.get('c')} "
                    f"timestamp={sample.get('t')}"
                )
            return f"{ticker}: no data returned"
        except Exception as exc:
            return f"{ticker}: error fetching data ({exc})"

    lines: list[str] = []
    max_workers = min(8, max(1, len(tickers)))
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_map = {executor.submit(_process_ticker, ticker): ticker for ticker in tickers}
        for future in as_completed(future_map):
            lines.append(future.result())

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = BACKTEST_OUTPUT_DIR / f"backtest_cache_{timestamp}.txt"
    output_path.write_text("\n".join(lines))
    return "\n".join(lines) + f"\n\nSaved summary to: {output_path}"



def load_backtest_engine_arrays(
    tickers: list[str],
    start: datetime | str,
    end: datetime | str,
    *,
    cache_root: Path | None = None,
    timeframe: str = "1m",
    lookback_window: int = 0,
) -> EngineArrayBundle:
    """Load canonical float64 arrays and metadata for backtest engines."""

    return load_canonical_price_arrays(
        symbols=tickers,
        start=start,
        end=end,
        cache_root=cache_root,
        timeframe=timeframe,
        lookback_window=lookback_window,
        validate_split_adjustment=True,
    )

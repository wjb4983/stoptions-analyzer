from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, Iterable, List, Optional

import numpy as np


class BacktestBarStoreConnector:
    """Loads cached OHLCV(+n) bars from NPZ yearly shards and merges them."""

    REQUIRED_KEYS = ("t", "o", "h", "l", "c", "v", "n")

    def __init__(self, roots: Optional[Iterable[Path]] = None) -> None:
        self.roots = list(roots) if roots else [
            Path("src/data/backtest_cache"),
            Path("../stoptions_analyzer/src/data/backtest_cache"),
            Path("../stoptions-analyzer/src/data/backtest_cache"),
        ]

    def _base_dir(self, safe_ticker: str, timeframe: str) -> Path:
        for root in self.roots:
            candidate = root / safe_ticker / timeframe
            if candidate.exists():
                return candidate
        searched = ", ".join(str(r / safe_ticker / timeframe) for r in self.roots)
        raise FileNotFoundError(f"Backtest cache directory not found. Searched: {searched}")

    def load_index(self, safe_ticker: str, timeframe: str) -> Dict:
        base = self._base_dir(safe_ticker, timeframe)
        index_path = base / "index.json"
        if not index_path.exists():
            raise FileNotFoundError(f"Missing {index_path}")
        with index_path.open("r", encoding="utf-8") as f:
            return json.load(f)

    def load_bars(self, safe_ticker: str, timeframe: str, years: Optional[List[int]] = None) -> Dict[str, np.ndarray]:
        base = self._base_dir(safe_ticker, timeframe)
        index = self.load_index(safe_ticker, timeframe)
        selected_years = years or sorted(index.get("years", []))
        merged = {k: [] for k in self.REQUIRED_KEYS}

        for year in selected_years:
            shard = base / f"{safe_ticker}_{timeframe}_{year}.npz"
            if not shard.exists():
                continue
            with np.load(shard) as arr:
                for key in self.REQUIRED_KEYS:
                    if key not in arr:
                        raise KeyError(f"Missing key {key} in {shard}")
                    merged[key].append(arr[key])

        return {k: np.concatenate(v) if v else np.array([]) for k, v in merged.items()}

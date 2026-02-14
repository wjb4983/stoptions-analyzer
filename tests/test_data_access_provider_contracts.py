from __future__ import annotations

import json
from datetime import datetime, timezone

import numpy as np

from src.data_access.actions_schema import REQUIRED_ACTION_FIELDS, validate_actions_frame
from src.data_access.bars_schema import REQUIRED_BAR_FIELDS, validate_bars_frame
from src.data_access.provider_base import DataProvider
from src.data_access.providers.massive_provider import MassiveCacheProvider
from tests.fixtures_datasets import synthetic_corporate_actions_dataset, synthetic_vendor_bars_dataset


def _seed_provider_cache(tmp_path) -> MassiveCacheProvider:
    symbol = "AAPL"
    symbol_dir = tmp_path / symbol / "1m"
    symbol_dir.mkdir(parents=True, exist_ok=True)
    bars = synthetic_vendor_bars_dataset()
    np.savez_compressed(
        symbol_dir / f"{symbol}_1m_2024.npz",
        t=np.array([row["t"] for row in bars], dtype=np.int64),
        o=np.array([row["o"] for row in bars], dtype=float),
        h=np.array([row["h"] for row in bars], dtype=float),
        l=np.array([row["l"] for row in bars], dtype=float),
        c=np.array([row["c"] for row in bars], dtype=float),
        v=np.array([row["v"] for row in bars], dtype=float),
        n=np.array([row["n"] for row in bars], dtype=np.int64),
    )
    (symbol_dir / "corporate_actions.json").write_text(json.dumps(synthetic_corporate_actions_dataset()))
    return MassiveCacheProvider(cache_root=tmp_path, symbols=[symbol])


def test_provider_base_required_interface_methods(tmp_path) -> None:
    provider = _seed_provider_cache(tmp_path)

    assert isinstance(provider, DataProvider)
    assert callable(provider.list_symbols)
    assert callable(provider.get_bars)
    assert callable(provider.get_corporate_actions)


def test_massive_provider_bars_schema_contract(tmp_path) -> None:
    provider = _seed_provider_cache(tmp_path)
    bars = provider.get_bars(
        ["AAPL"],
        start=datetime(2024, 1, 2, 14, 30, tzinfo=timezone.utc),
        end=datetime(2024, 1, 2, 14, 32, tzinfo=timezone.utc),
    )

    if hasattr(bars, "columns"):
        validate_bars_frame(bars)
    if hasattr(bars, "to_dict"):
        rows = bars.to_dict(orient="records")
    else:
        rows = list(bars)

    assert rows
    for row in rows:
        for field in REQUIRED_BAR_FIELDS:
            assert field in row
        assert row["symbol"] == "AAPL"


def test_massive_provider_actions_schema_contract(tmp_path) -> None:
    provider = _seed_provider_cache(tmp_path)
    actions = provider.get_corporate_actions(
        ["AAPL"],
        start=datetime(2024, 1, 1, tzinfo=timezone.utc),
        end=datetime(2024, 2, 2, tzinfo=timezone.utc),
    )

    if hasattr(actions, "columns"):
        validate_actions_frame(actions)
    if hasattr(actions, "to_dict"):
        rows = actions.to_dict(orient="records")
    else:
        rows = list(actions)

    assert rows
    for row in rows:
        for field in REQUIRED_ACTION_FIELDS:
            assert field in row
        assert row["symbol"] == "AAPL"

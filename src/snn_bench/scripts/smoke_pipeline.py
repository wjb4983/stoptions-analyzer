from __future__ import annotations

import argparse
from pathlib import Path

from snn_bench.data_connectors import BacktestBarStoreConnector, SnapshotCacheConnector
from snn_bench.eval import classification_metrics
from snn_bench.feature_pipelines import BasicFeaturePipeline
from snn_bench.models import DummySpikingModel
from snn_bench.tasks import DirectionClassificationTask
from snn_bench.trainers import SimpleTrainer


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="SNN benchmark smoke pipeline")
    p.add_argument("--ticker", default="NVDA")
    p.add_argument("--timeframe", default="1D")
    p.add_argument("--data-root", default=None, help="Optional override for src/data root")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    roots = [Path(args.data_root)] if args.data_root else None

    snapshot_conn = SnapshotCacheConnector(roots=roots)
    backtest_conn = BacktestBarStoreConnector(
        roots=[Path(args.data_root) / "backtest_cache"] if args.data_root else None
    )

    try:
        snapshot = snapshot_conn.load(args.ticker)
        print(f"Loaded snapshot for {args.ticker}: keys={list(snapshot)[:5]}")
    except FileNotFoundError as e:
        print(f"Snapshot unavailable: {e}")

    try:
        bars = backtest_conn.load_bars(args.ticker, args.timeframe)
    except FileNotFoundError as e:
        print(f"Backtest cache unavailable: {e}; using synthetic bars for smoke run")
        import numpy as np
        bars = {
            "t": np.arange(256),
            "o": np.linspace(100, 120, 256),
            "h": np.linspace(101, 121, 256),
            "l": np.linspace(99, 119, 256),
            "c": np.linspace(100, 120, 256) + np.sin(np.linspace(0, 10, 256)),
            "v": np.ones(256),
            "n": np.ones(256),
        }
    pipeline = BasicFeaturePipeline()
    features = pipeline.transform(bars)

    task = DirectionClassificationTask()
    x, y = task.make_dataset(features, bars)

    split = int(len(x) * 0.8) if len(x) else 0
    x_train, y_train = x[:split], y[:split]
    x_eval, y_eval = x[split:], y[split:]

    trainer = SimpleTrainer(DummySpikingModel(seed=42))
    y_pred = trainer.run(x_train, y_train, x_eval)

    metrics = classification_metrics(y_eval, y_pred)
    print(f"Smoke run complete: {metrics}")


if __name__ == "__main__":
    main()

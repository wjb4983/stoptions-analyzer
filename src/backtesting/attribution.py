from __future__ import annotations

import csv
import json
from pathlib import Path
from typing import Any


def write_attribution_artifacts(*, run_dir: Path, payload: dict[str, Any]) -> None:
    series = payload.get("time_series", []) if isinstance(payload, dict) else []
    summary = payload.get("summary", []) if isinstance(payload, dict) else []

    if isinstance(series, list) and series:
        (run_dir / "attribution_timeseries.json").write_text(json.dumps(series, indent=2))
        fieldnames = sorted({str(key) for row in series if isinstance(row, dict) for key in row.keys()})
        with (run_dir / "attribution_timeseries.csv").open("w", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=fieldnames)
            writer.writeheader()
            for row in series:
                if isinstance(row, dict):
                    writer.writerow(row)

    if isinstance(summary, list) and summary:
        (run_dir / "attribution_summary.json").write_text(json.dumps(summary, indent=2))
        fieldnames = sorted({str(key) for row in summary if isinstance(row, dict) for key in row.keys()})
        with (run_dir / "attribution_summary.csv").open("w", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=fieldnames)
            writer.writeheader()
            for row in summary:
                if isinstance(row, dict):
                    writer.writerow(row)


from __future__ import annotations

import argparse
import sys
from pathlib import Path

import numpy as np

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from src.modeling_nextgen.features.no_arb import detect_and_repair_no_arb, export_no_arb_diagnostics


REPORTS_DIR = Path(__file__).resolve().parent
DEFAULT_SURFACE_NPZ_PATH = REPORTS_DIR / "surface_total_variance_snapshot.npz"
DEFAULT_OUTPUT_PATH = REPORTS_DIR / "no_arb_diagnostics_report.json"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate no-arbitrage diagnostics for model gating/reporting.")
    parser.add_argument("--surface", default=str(DEFAULT_SURFACE_NPZ_PATH))
    parser.add_argument("--out", default=str(DEFAULT_OUTPUT_PATH))
    parser.add_argument("--threshold", type=int, default=0)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    payload = np.load(Path(args.surface))
    total_variance = np.asarray(payload["total_variance"], dtype=np.float64)
    moneyness = np.asarray(payload["moneyness"], dtype=np.float64)

    result = detect_and_repair_no_arb(total_variance, moneyness)
    report = export_no_arb_diagnostics(result.diagnostics, out_path=args.out, model_gate_threshold=args.threshold)
    return 0 if bool(report["model_gate"]["pass"]) else 1


if __name__ == "__main__":
    raise SystemExit(main())

from __future__ import annotations

import json
from pathlib import Path

from execution.run_summary import build_workflow_summary


def test_build_workflow_summary_writes_summary_bundle(tmp_path: Path) -> None:
    run_dir = tmp_path / "run"
    run_dir.mkdir()
    (run_dir / "leaderboard.json").write_text(json.dumps([{"sharpe": 1.4, "total_return": 0.2}]), encoding="utf-8")
    (run_dir / "metrics.json").write_text(json.dumps([{"metric": "sharpe", "value": 1.4}, {"metric": "max_drawdown", "value": -0.1}]), encoding="utf-8")
    (run_dir / "manifest.json").write_text(json.dumps({"governance_metadata": {"gate_checks": {"approval": False}}}), encoding="utf-8")
    (run_dir / "trade_explainability.json").write_text(json.dumps([{"regime": "bull", "red_flags": ["slippage"]}]), encoding="utf-8")

    created = build_workflow_summary(run_dir=run_dir, workflow="backtesting")

    assert "summary/leaderboard_stats.json" in created
    assert (run_dir / "summary" / "summary_metrics.csv").exists()
    governance = json.loads((run_dir / "summary" / "governance_checks.json").read_text(encoding="utf-8"))
    assert governance["failed_gate_checks"] == ["approval"]

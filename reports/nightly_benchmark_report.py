from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


REPORTS_DIR = Path(__file__).resolve().parent
DEFAULT_PRIOR_BASELINE_PATH = REPORTS_DIR / "prior_baseline_metrics_summary.json"
DEFAULT_MAINLINE_PATH = REPORTS_DIR / "baseline_metrics_summary.json"
DEFAULT_JSON_OUT = REPORTS_DIR / "nightly_benchmark_comparison.json"
DEFAULT_MD_OUT = REPORTS_DIR / "nightly_benchmark_report.md"



def _load_json(path: Path) -> Any:
    return json.loads(path.read_text())



def _index_by_scenario(rows: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    return {str(row.get("scenario", "unknown")): row for row in rows}



def build_comparison(*, mainline: list[dict[str, Any]], prior: list[dict[str, Any]]) -> dict[str, Any]:
    current = _index_by_scenario(mainline)
    baseline = _index_by_scenario(prior)

    scenarios = sorted(set(current) | set(baseline))
    rows: list[dict[str, Any]] = []
    for scenario in scenarios:
        main_row = current.get(scenario, {})
        prior_row = baseline.get(scenario, {})

        main_sharpe = float(main_row.get("sharpe", 0.0))
        prior_sharpe = float(prior_row.get("sharpe", 0.0))
        main_return = float(main_row.get("total_return", 0.0))
        prior_return = float(prior_row.get("total_return", 0.0))

        rows.append(
            {
                "scenario": scenario,
                "mainline": {"sharpe": main_sharpe, "total_return": main_return},
                "prior_baseline": {"sharpe": prior_sharpe, "total_return": prior_return},
                "delta": {
                    "sharpe": main_sharpe - prior_sharpe,
                    "total_return": main_return - prior_return,
                },
            }
        )

    improved = sum(1 for row in rows if row["delta"]["sharpe"] >= 0.0)
    return {
        "scenarios_compared": len(rows),
        "mainline_beats_or_matches_prior_on_sharpe": improved,
        "rows": rows,
    }



def render_markdown(comparison: dict[str, Any]) -> str:
    lines = [
        "# Nightly Benchmark Report",
        "",
        f"Scenarios compared: **{comparison['scenarios_compared']}**",
        f"Mainline sharpe >= prior baseline: **{comparison['mainline_beats_or_matches_prior_on_sharpe']}** scenarios",
        "",
        "| Scenario | Mainline Sharpe | Prior Sharpe | Δ Sharpe | Mainline Return | Prior Return | Δ Return |",
        "|---|---:|---:|---:|---:|---:|---:|",
    ]
    for row in comparison["rows"]:
        lines.append(
            "| {scenario} | {m_sh:.4f} | {p_sh:.4f} | {d_sh:.4f} | {m_rt:.4f} | {p_rt:.4f} | {d_rt:.4f} |".format(
                scenario=row["scenario"],
                m_sh=row["mainline"]["sharpe"],
                p_sh=row["prior_baseline"]["sharpe"],
                d_sh=row["delta"]["sharpe"],
                m_rt=row["mainline"]["total_return"],
                p_rt=row["prior_baseline"]["total_return"],
                d_rt=row["delta"]["total_return"],
            )
        )
    lines.append("")
    return "\n".join(lines)



def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Publish nightly benchmark comparison report.")
    parser.add_argument("--mainline", default=str(DEFAULT_MAINLINE_PATH))
    parser.add_argument("--prior", default=str(DEFAULT_PRIOR_BASELINE_PATH))
    parser.add_argument("--json-out", default=str(DEFAULT_JSON_OUT))
    parser.add_argument("--md-out", default=str(DEFAULT_MD_OUT))
    return parser.parse_args()



def main() -> int:
    args = parse_args()
    comparison = build_comparison(
        mainline=_load_json(Path(args.mainline)),
        prior=_load_json(Path(args.prior)),
    )
    Path(args.json_out).write_text(json.dumps(comparison, indent=2) + "\n")
    Path(args.md_out).write_text(render_markdown(comparison))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

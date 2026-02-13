from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path


PROMPT_TEMPLATES: dict[str, str] = {
    "strategy_ideation": """## Strategy Ideation Prompt
You are a quantitative strategy researcher. Use the supplied context to propose 3 distinct strategy extensions.

### Deliverables
1. Thesis for each strategy in 1-2 sentences.
2. Entry/exit logic and position sizing assumptions.
3. What market regime each strategy should outperform in.
4. Failure modes and observable invalidation triggers.

### Constraints
- Use only information explicitly present in the context payload.
- Separate assumptions from evidence.
- Rank ideas by expected robustness and implementation difficulty.
""",
    "validation_design": """## Validation Design Prompt
You are designing a validation protocol for the candidate strategy and configuration.

### Deliverables
1. Walk-forward or holdout design with clear splits.
2. Metrics to report (performance, risk, turnover, capacity, stability).
3. Leakage and overfitting checks.
4. Decision thresholds for promotion/rejection.

### Constraints
- Explicitly justify why each validation step is required.
- Include at least one stress scenario and one ablation.
- Map each check to a specific expected failure mode.
""",
    "risk_model_critique": """## Risk Model Critique Prompt
You are an independent risk reviewer.

### Deliverables
1. Critique assumptions behind the current risk model.
2. Identify blind spots in exposure controls and concentration constraints.
3. Recommend improvements to drawdown, tail-risk, and scenario monitoring.
4. Prioritize fixes by severity and implementation effort.

### Constraints
- Focus on practical portfolio and execution risks.
- Flag where available diagnostics are insufficient to support conclusions.
- Provide concrete checks that can be automated.
""",
    "execution_realism_audit": """## Execution Realism Audit Prompt
You are auditing whether backtest execution assumptions are realistic.

### Deliverables
1. Identify optimistic assumptions in fills, slippage, and latency.
2. Recommend parameter ranges for conservative/neutral/aggressive execution settings.
3. Define experiments to quantify execution sensitivity.
4. Explain how execution assumptions could invert the strategy conclusion.

### Constraints
- Tie every critique to a specific config field or output artifact.
- Distinguish structural issues from calibration issues.
- Include at least 2 edge cases (liquidity shock, gap move, etc.).
""",
    "statistical_significance_review": """## Statistical Significance Review Prompt
You are reviewing whether observed performance is statistically credible.

### Deliverables
1. Evaluate sample size adequacy and independence assumptions.
2. Recommend significance tests and confidence interval framing.
3. Assess multiple-testing / selection-bias risk.
4. Provide a go/no-go confidence statement with caveats.

### Constraints
- Distinguish economic significance vs statistical significance.
- Request additional evidence where current outputs are insufficient.
- Call out any metric instability across splits/regimes.
""",
}


def build_prompt_pack_markdown(*, title: str, config: dict[str, object], recent_outputs: dict[str, object] | None = None) -> str:
    lines = [
        f"# {title}",
        "",
        f"Generated at: {datetime.utcnow().isoformat()}Z",
        "",
        "## Context",
        "### Current Configuration",
        "```json",
        json.dumps(config, indent=2, sort_keys=True, default=str),
        "```",
    ]
    if recent_outputs:
        lines.extend(
            [
                "",
                "### Recent Run Outputs",
                "```json",
                json.dumps(recent_outputs, indent=2, sort_keys=True, default=str),
                "```",
            ]
        )

    lines.append("")
    lines.append("## Structured Prompts")
    for key, prompt in PROMPT_TEMPLATES.items():
        lines.append("")
        lines.append(f"### {key.replace('_', ' ').title()}")
        lines.append(prompt.strip())
    lines.append("")
    return "\n".join(lines)


def write_prompt_pack(*, output_dir: Path, file_stem: str, title: str, config: dict[str, object], recent_outputs: dict[str, object] | None = None) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    output_path = output_dir / f"{file_stem}_{timestamp}.md"
    markdown = build_prompt_pack_markdown(title=title, config=config, recent_outputs=recent_outputs)
    output_path.write_text(markdown, encoding="utf-8")
    return output_path


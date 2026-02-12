from __future__ import annotations

import json
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Protocol

import numpy as np


ObjectiveSense = str


@dataclass(frozen=True)
class Objective:
    name: str
    sense: ObjectiveSense = "maximize"  # maximize | minimize


@dataclass(frozen=True)
class Constraint:
    metric: str
    min_value: float | None = None
    max_value: float | None = None


@dataclass
class Trial:
    trial_id: int
    params: dict[str, Any]
    metrics: dict[str, float]
    feasible: bool
    stopped_early: bool
    period_fraction: float


class Sampler(Protocol):
    def sample(self, *, trials: list[Trial], space: dict[str, list[Any]], rng: np.random.Generator) -> dict[str, Any]:
        ...


class RandomSampler:
    def sample(self, *, trials: list[Trial], space: dict[str, list[Any]], rng: np.random.Generator) -> dict[str, Any]:
        return {k: values[int(rng.integers(0, len(values)))] for k, values in space.items()}


class TPESampler:
    """Lightweight TPE-style sampler over discrete spaces."""

    def __init__(self, *, gamma: float = 0.25, random_fraction: float = 0.2) -> None:
        self.gamma = float(gamma)
        self.random_fraction = float(random_fraction)

    def sample(self, *, trials: list[Trial], space: dict[str, list[Any]], rng: np.random.Generator) -> dict[str, Any]:
        if not trials or rng.random() < self.random_fraction:
            return RandomSampler().sample(trials=trials, space=space, rng=rng)

        feasible = [trial for trial in trials if trial.feasible and not trial.stopped_early]
        if not feasible:
            return RandomSampler().sample(trials=trials, space=space, rng=rng)

        ranked = sorted(feasible, key=lambda t: float(t.metrics.get("_scalar_score", -math.inf)), reverse=True)
        n_good = max(1, int(math.ceil(len(ranked) * self.gamma)))
        good = ranked[:n_good]

        sampled: dict[str, Any] = {}
        for key, choices in space.items():
            counts = {repr(choice): 1.0 for choice in choices}
            rep_to_value = {repr(choice): choice for choice in choices}
            for trial in good:
                selected = repr(trial.params[key])
                counts[selected] = counts.get(selected, 1.0) + 1.0
            labels = list(counts.keys())
            probs = np.array([counts[label] for label in labels], dtype=float)
            probs /= probs.sum()
            sampled[key] = rep_to_value[str(rng.choice(labels, p=probs))]
        return sampled


def compute_scalar_score(metrics: dict[str, float], objectives: list[Objective]) -> float:
    score = 0.0
    for obj in objectives:
        value = float(metrics.get(obj.name, -math.inf if obj.sense == "maximize" else math.inf))
        score += value if obj.sense == "maximize" else -value
    return score


def check_constraints(metrics: dict[str, float], constraints: list[Constraint]) -> bool:
    for cons in constraints:
        value = float(metrics.get(cons.metric, float("nan")))
        if math.isnan(value):
            return False
        if cons.min_value is not None and value < float(cons.min_value):
            return False
        if cons.max_value is not None and value > float(cons.max_value):
            return False
    return True


def pareto_frontier(rows: list[dict[str, Any]], objectives: list[Objective]) -> list[dict[str, Any]]:
    frontier: list[dict[str, Any]] = []
    for row in rows:
        dominated = False
        for other in rows:
            if other is row:
                continue
            better_or_equal = True
            strictly_better = False
            for obj in objectives:
                rv = float(row["metrics"].get(obj.name, float("nan")))
                ov = float(other["metrics"].get(obj.name, float("nan")))
                if obj.sense == "maximize":
                    if ov < rv:
                        better_or_equal = False
                        break
                    if ov > rv:
                        strictly_better = True
                else:
                    if ov > rv:
                        better_or_equal = False
                        break
                    if ov < rv:
                        strictly_better = True
            if better_or_equal and strictly_better:
                dominated = True
                break
        if not dominated:
            frontier.append(row)
    return frontier


def optimize(
    *,
    space: dict[str, list[Any]],
    evaluate: Callable[[dict[str, Any], float], dict[str, float]],
    objectives: list[Objective],
    constraints: list[Constraint],
    sampler: Sampler,
    n_trials: int,
    seed: int,
    partial_period_fractions: list[float] | None,
    output_dir: Path,
) -> dict[str, Any]:
    output_dir.mkdir(parents=True, exist_ok=True)
    fractions = sorted(set(partial_period_fractions or [1.0]))
    if fractions[-1] != 1.0:
        fractions.append(1.0)

    rng = np.random.default_rng(seed)
    trials: list[Trial] = []
    jsonl_path = output_dir / "trials.jsonl"

    with jsonl_path.open("w", encoding="utf-8") as handle:
        for trial_id in range(int(n_trials)):
            params = sampler.sample(trials=trials, space=space, rng=rng)
            last_metrics: dict[str, float] = {}
            feasible = False
            stopped_early = False
            period_fraction = 1.0
            for period_fraction in fractions:
                last_metrics = {k: float(v) for k, v in evaluate(params, float(period_fraction)).items()}
                last_metrics["_scalar_score"] = compute_scalar_score(last_metrics, objectives)
                feasible = check_constraints(last_metrics, constraints)
                if not feasible and period_fraction < 1.0:
                    stopped_early = True
                    break
            trial = Trial(
                trial_id=trial_id,
                params=dict(params),
                metrics=last_metrics,
                feasible=feasible,
                stopped_early=stopped_early,
                period_fraction=float(period_fraction),
            )
            trials.append(trial)
            handle.write(
                json.dumps(
                    {
                        "trial_id": trial.trial_id,
                        "params": trial.params,
                        "metrics": trial.metrics,
                        "feasible": trial.feasible,
                        "stopped_early": trial.stopped_early,
                        "period_fraction": trial.period_fraction,
                    },
                    sort_keys=True,
                )
                + "\n"
            )

    feasible_rows = [
        {"trial_id": t.trial_id, "params": t.params, "metrics": t.metrics}
        for t in trials
        if t.feasible and not t.stopped_early
    ]
    frontier = pareto_frontier(feasible_rows, objectives)
    (output_dir / "pareto_frontier.json").write_text(json.dumps(frontier, indent=2, sort_keys=True), encoding="utf-8")

    return {
        "trial_count": len(trials),
        "feasible_count": len(feasible_rows),
        "pareto_count": len(frontier),
        "best_scalar": max((float(t.metrics.get("_scalar_score", -math.inf)) for t in trials), default=-math.inf),
        "trials": trials,
        "pareto_frontier": frontier,
        "trials_path": str(jsonl_path),
        "pareto_path": str(output_dir / "pareto_frontier.json"),
    }

from __future__ import annotations

import json
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Protocol

import numpy as np


ObjectiveSense = str


@dataclass(frozen=True)
class ContinuousDimension:
    low: float
    high: float
    log_scale: bool = False
    step: float | None = None


@dataclass(frozen=True)
class DiscreteDimension:
    values: list[Any]


SearchDimension = ContinuousDimension | DiscreteDimension


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
    uncertainty: float | None = None
    fold_scores: list[float] | None = None
    robust_score: float | None = None


class Sampler(Protocol):
    def sample(self, *, trials: list[Trial], space: dict[str, SearchDimension], rng: np.random.Generator) -> dict[str, Any]:
        ...

    def predict(self, *, trials: list[Trial], space: dict[str, SearchDimension], params: dict[str, Any]) -> tuple[float, float] | None:
        ...


def _normalize_space(space: dict[str, Any]) -> dict[str, SearchDimension]:
    normalized: dict[str, SearchDimension] = {}
    for key, raw in space.items():
        if isinstance(raw, dict):
            dim_type = str(raw.get("type", "discrete")).lower()
            if dim_type == "continuous":
                low = float(raw["low"])
                high = float(raw["high"])
                if high <= low:
                    raise ValueError(f"Continuous dimension '{key}' requires high > low")
                normalized[key] = ContinuousDimension(
                    low=low,
                    high=high,
                    log_scale=bool(raw.get("log", False)),
                    step=None if raw.get("step") is None else float(raw.get("step")),
                )
            else:
                values = list(raw.get("values", []))
                if not values:
                    raise ValueError(f"Discrete dimension '{key}' requires non-empty values")
                normalized[key] = DiscreteDimension(values=values)
            continue
        values = list(raw)
        if not values:
            raise ValueError(f"Discrete dimension '{key}' requires non-empty values")
        normalized[key] = DiscreteDimension(values=values)
    return normalized


class RandomSampler:
    def sample(self, *, trials: list[Trial], space: dict[str, SearchDimension], rng: np.random.Generator) -> dict[str, Any]:
        sampled: dict[str, Any] = {}
        for key, dim in space.items():
            if isinstance(dim, DiscreteDimension):
                sampled[key] = dim.values[int(rng.integers(0, len(dim.values)))]
            else:
                if dim.log_scale:
                    value = float(np.exp(rng.uniform(np.log(dim.low), np.log(dim.high))))
                else:
                    value = float(rng.uniform(dim.low, dim.high))
                if dim.step and dim.step > 0:
                    n_steps = round((value - dim.low) / dim.step)
                    value = dim.low + (n_steps * dim.step)
                    value = float(min(dim.high, max(dim.low, value)))
                sampled[key] = value
        return sampled

    def predict(self, *, trials: list[Trial], space: dict[str, SearchDimension], params: dict[str, Any]) -> tuple[float, float] | None:
        return None


class TPESampler:
    """Lightweight TPE-style sampler over discrete spaces."""

    def __init__(self, *, gamma: float = 0.25, random_fraction: float = 0.2) -> None:
        self.gamma = float(gamma)
        self.random_fraction = float(random_fraction)

    def sample(self, *, trials: list[Trial], space: dict[str, SearchDimension], rng: np.random.Generator) -> dict[str, Any]:
        if not trials or rng.random() < self.random_fraction:
            return RandomSampler().sample(trials=trials, space=space, rng=rng)

        feasible = [trial for trial in trials if trial.feasible and not trial.stopped_early]
        if not feasible:
            return RandomSampler().sample(trials=trials, space=space, rng=rng)

        ranked = sorted(feasible, key=lambda t: float(t.metrics.get("_scalar_score", -math.inf)), reverse=True)
        n_good = max(1, int(math.ceil(len(ranked) * self.gamma)))
        good = ranked[:n_good]

        sampled: dict[str, Any] = {}
        for key, dim in space.items():
            if not isinstance(dim, DiscreteDimension):
                return RandomSampler().sample(trials=trials, space=space, rng=rng)
            choices = dim.values
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

    def predict(self, *, trials: list[Trial], space: dict[str, SearchDimension], params: dict[str, Any]) -> tuple[float, float] | None:
        return None


class BayesianSampler:
    """Simple GP-style Bayesian sampler over mixed continuous/discrete spaces."""

    def __init__(self, *, random_fraction: float = 0.15, candidate_pool_size: int = 128, beta: float = 1.5) -> None:
        self.random_fraction = float(random_fraction)
        self.candidate_pool_size = int(candidate_pool_size)
        self.beta = float(beta)

    def _encode(self, params: dict[str, Any], space: dict[str, SearchDimension]) -> np.ndarray:
        vec: list[float] = []
        for key, dim in space.items():
            value = params[key]
            if isinstance(dim, DiscreteDimension):
                idx = next((i for i, item in enumerate(dim.values) if item == value), 0)
                denom = max(1, len(dim.values) - 1)
                vec.append(float(idx) / float(denom))
            else:
                v = float(value)
                if dim.log_scale:
                    lo = np.log(dim.low)
                    hi = np.log(dim.high)
                    vec.append(float((np.log(v) - lo) / max(1e-9, hi - lo)))
                else:
                    vec.append(float((v - dim.low) / max(1e-9, dim.high - dim.low)))
        return np.array(vec, dtype=float)

    def _fit_predict(self, *, trials: list[Trial], space: dict[str, SearchDimension], params: dict[str, Any]) -> tuple[float, float] | None:
        observed = [t for t in trials if not t.stopped_early and t.metrics.get("_scalar_score") is not None]
        if len(observed) < 3:
            return None
        x = np.vstack([self._encode(t.params, space) for t in observed])
        y = np.array([float(t.metrics["_scalar_score"]) for t in observed], dtype=float)
        xq = self._encode(params, space)
        d2 = np.sum((x[:, None, :] - x[None, :, :]) ** 2, axis=2)
        k = np.exp(-2.0 * d2) + np.eye(len(observed)) * 1e-6
        k_inv = np.linalg.pinv(k)
        kq = np.exp(-2.0 * np.sum((x - xq[None, :]) ** 2, axis=1))
        mean = float(kq @ k_inv @ y)
        var = max(1e-9, float(1.0 - (kq @ k_inv @ kq)))
        return mean, math.sqrt(var)

    def sample(self, *, trials: list[Trial], space: dict[str, SearchDimension], rng: np.random.Generator) -> dict[str, Any]:
        random_sampler = RandomSampler()
        if len(trials) < 4 or rng.random() < self.random_fraction:
            return random_sampler.sample(trials=trials, space=space, rng=rng)
        candidates = [random_sampler.sample(trials=trials, space=space, rng=rng) for _ in range(self.candidate_pool_size)]
        scored: list[tuple[float, dict[str, Any]]] = []
        for candidate in candidates:
            pred = self._fit_predict(trials=trials, space=space, params=candidate)
            if pred is None:
                continue
            mean, std = pred
            scored.append((mean + self.beta * std, candidate))
        if not scored:
            return random_sampler.sample(trials=trials, space=space, rng=rng)
        scored.sort(key=lambda item: item[0], reverse=True)
        return dict(scored[0][1])

    def predict(self, *, trials: list[Trial], space: dict[str, SearchDimension], params: dict[str, Any]) -> tuple[float, float] | None:
        return self._fit_predict(trials=trials, space=space, params=params)


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
    space: dict[str, Any],
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
    normalized_space = _normalize_space(space)
    fractions = sorted(set(partial_period_fractions or [1.0]))
    if fractions[-1] != 1.0:
        fractions.append(1.0)

    rng = np.random.default_rng(seed)
    trials: list[Trial] = []
    jsonl_path = output_dir / "trials.jsonl"

    with jsonl_path.open("w", encoding="utf-8") as handle:
        for trial_id in range(int(n_trials)):
            params = sampler.sample(trials=trials, space=normalized_space, rng=rng)
            last_metrics: dict[str, float] = {}
            feasible = False
            stopped_early = False
            period_fraction = 1.0
            fold_scores: list[float] = []
            pred = sampler.predict(trials=trials, space=normalized_space, params=params)
            predicted_uncertainty = None if pred is None else float(pred[1])
            for period_fraction in fractions:
                last_metrics = {k: float(v) for k, v in evaluate(params, float(period_fraction)).items()}
                last_metrics["_scalar_score"] = compute_scalar_score(last_metrics, objectives)
                fold_scores.append(float(last_metrics["_scalar_score"]))
                feasible = check_constraints(last_metrics, constraints)
                if not feasible and period_fraction < 1.0:
                    stopped_early = True
                    break
                if period_fraction < 1.0:
                    completed_scores = [
                        float(t.metrics.get("_scalar_score", -math.inf))
                        for t in trials
                        if t.feasible and not t.stopped_early and t.period_fraction >= 1.0
                    ]
                    if len(completed_scores) >= 5:
                        prune_bar = float(np.quantile(np.array(completed_scores), 0.25))
                        if float(last_metrics["_scalar_score"]) < (prune_bar - 0.25):
                            stopped_early = True
                            feasible = False
                            break

            robust_score = float(np.mean(fold_scores) - np.std(fold_scores)) if fold_scores else None
            trial = Trial(
                trial_id=trial_id,
                params=dict(params),
                metrics=last_metrics,
                feasible=feasible,
                stopped_early=stopped_early,
                period_fraction=float(period_fraction),
                uncertainty=predicted_uncertainty,
                fold_scores=list(fold_scores),
                robust_score=robust_score,
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
                        "uncertainty": trial.uncertainty,
                        "fold_scores": trial.fold_scores,
                        "robust_score": trial.robust_score,
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

    trace_rows = [
        {
            "trial_id": t.trial_id,
            "candidate": t.params,
            "score": float(t.metrics.get("_scalar_score", -math.inf)),
            "uncertainty": t.uncertainty,
            "robust_score": t.robust_score,
            "stopped_early": t.stopped_early,
        }
        for t in trials
    ]
    (output_dir / "optimization_trace.json").write_text(json.dumps(trace_rows, indent=2, sort_keys=True), encoding="utf-8")
    robust_best = [
        {
            "trial_id": t.trial_id,
            "params": t.params,
            "metrics": t.metrics,
            "robust_score": t.robust_score,
            "uncertainty": t.uncertainty,
        }
        for t in sorted(
            [trial for trial in trials if trial.feasible and not trial.stopped_early and trial.period_fraction >= 1.0],
            key=lambda trial: float(trial.robust_score if trial.robust_score is not None else -math.inf),
            reverse=True,
        )[:5]
    ]
    (output_dir / "best_robust_params.json").write_text(json.dumps(robust_best, indent=2, sort_keys=True), encoding="utf-8")
    (output_dir / "artifact_metadata.json").write_text(
        json.dumps(
            {
                "schema_version": "1.0",
                "run_type": "optimization_trials",
                "random_seeds": {"run_seed": int(seed), "numpy_random_seed": int(seed)},
            },
            indent=2,
            sort_keys=True,
        ),
        encoding="utf-8",
    )

    return {
        "trial_count": len(trials),
        "feasible_count": len(feasible_rows),
        "pareto_count": len(frontier),
        "best_scalar": max((float(t.metrics.get("_scalar_score", -math.inf)) for t in trials), default=-math.inf),
        "trials": trials,
        "pareto_frontier": frontier,
        "pareto_trials": frontier,
        "optimization_trace_path": str(output_dir / "optimization_trace.json"),
        "best_robust_params_path": str(output_dir / "best_robust_params.json"),
        "best_robust_params": robust_best,
        "trials_path": str(jsonl_path),
        "pareto_path": str(output_dir / "pareto_frontier.json"),
    }

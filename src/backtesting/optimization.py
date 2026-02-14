from __future__ import annotations

import itertools
import json
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Protocol

import numpy as np

from .perf import benjamini_hochberg_adjusted_pvalues


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
    uncertainty: dict[str, float] | None = None
    fold_scores: list[float] | None = None
    oos_score_vector: list[float] | None = None
    robust_score: float | None = None


@dataclass(frozen=True)
class OverfittingPenaltyConfig:
    train_validation_gap_weight: float = 0.2
    fold_instability_weight: float = 0.1
    tail_risk_weight: float = 0.1


def _resolve_metric_value(metrics: dict[str, float], objective_name: str) -> float:
    if objective_name in metrics:
        return float(metrics[objective_name])
    alias_candidates = {
        "tail_risk": ("cvar_95", "expected_shortfall", "tail_expected_shortfall"),
        "max_drawdown": ("drawdown",),
        "turnover": ("turnover_total",),
    }
    for alias in alias_candidates.get(objective_name, ()):  # pragma: no branch - tiny mapping
        if alias in metrics:
            return float(metrics[alias])
    return float("nan")


def compute_bootstrap_reality_check(
    *,
    score_vectors: list[list[float]],
    n_bootstrap: int = 500,
    seed: int = 42,
) -> dict[str, float]:
    """Compute simple White Reality Check / SPA-style bootstrap diagnostics."""
    valid = [np.asarray(vec, dtype=float) for vec in score_vectors if len(vec) > 0]
    if not valid:
        return {
            "observed_stat": 0.0,
            "white_reality_check_pvalue": 1.0,
            "spa_pvalue": 1.0,
            "n_candidates": 0.0,
            "n_observations": 0.0,
        }

    n_obs = int(max(arr.size for arr in valid))
    aligned = np.column_stack([
        arr if arr.size == n_obs else np.pad(arr, (0, n_obs - arr.size), mode="edge")
        for arr in valid
    ])
    observed_means = np.mean(aligned, axis=0)
    observed = float(np.max(observed_means))
    centered = aligned - observed_means
    std = np.std(aligned, axis=0, ddof=1)
    std = np.where(std > 1e-12, std, 1.0)

    rng = np.random.default_rng(seed)
    white_stats: list[float] = []
    spa_stats: list[float] = []
    for _ in range(max(1, int(n_bootstrap))):
        idx = rng.integers(0, n_obs, size=n_obs)
        sampled = centered[idx, :]
        boot_means = np.mean(sampled, axis=0)
        white_stats.append(float(np.max(boot_means)))
        spa_stats.append(float(np.max(np.sqrt(n_obs) * boot_means / std)))

    observed_spa = float(np.max(np.sqrt(n_obs) * observed_means / std))
    return {
        "observed_stat": observed,
        "white_reality_check_pvalue": float(np.mean(np.asarray(white_stats) >= observed)) if white_stats else 1.0,
        "spa_pvalue": float(np.mean(np.asarray(spa_stats) >= observed_spa)) if spa_stats else 1.0,
        "n_candidates": float(aligned.shape[1]),
        "n_observations": float(n_obs),
    }


def compute_corrected_pvalues(*, score_vectors: list[list[float]], seed: int = 42) -> dict[str, Any]:
    """Combine BH-adjusted and bootstrap RC/SPA p-values into corrected significance terms."""
    valid = [np.asarray(vec, dtype=float) for vec in score_vectors if len(vec) > 0]
    if not valid:
        return {
            "raw_pvalues": [],
            "bh_adjusted_pvalues": [],
            "corrected_pvalues": [],
            "white_reality_check_pvalue": 1.0,
            "spa_pvalue": 1.0,
        }

    raw_pvalues: list[float] = []
    for vec in valid:
        mean = float(np.mean(vec))
        std = float(np.std(vec, ddof=1)) if vec.size > 1 else 0.0
        if std <= 1e-12:
            raw_pvalues.append(1.0 if mean <= 0.0 else 0.0)
            continue
        t_stat = mean / (std / math.sqrt(max(1, vec.size)))
        raw_pvalues.append(float(0.5 * math.erfc(t_stat / math.sqrt(2.0))))

    bh_adjusted = benjamini_hochberg_adjusted_pvalues(raw_pvalues)
    bootstrap = compute_bootstrap_reality_check(score_vectors=[vec.tolist() for vec in valid], seed=seed)
    white_p = float(bootstrap["white_reality_check_pvalue"])
    spa_p = float(bootstrap["spa_pvalue"])
    corrected = [float(max(p, white_p, spa_p)) for p in bh_adjusted]
    return {
        "raw_pvalues": raw_pvalues,
        "bh_adjusted_pvalues": bh_adjusted,
        "corrected_pvalues": corrected,
        "white_reality_check_pvalue": white_p,
        "spa_pvalue": spa_p,
    }


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
        observed = [t for t in trials if not t.stopped_early and t.metrics.get("_scalar_score") is not None]
        if len(observed) < 3:
            return None

        global_scores = np.array([float(t.metrics["_scalar_score"]) for t in observed], dtype=float)
        matched = [
            float(t.metrics["_scalar_score"])
            for t in observed
            if all(t.params.get(key) == params.get(key) for key in space)
        ]
        if matched:
            local_scores = np.array(matched, dtype=float)
            support = len(matched)
            blend = min(1.0, support / 3.0)
            mean = float((1.0 - blend) * np.mean(global_scores) + blend * np.mean(local_scores))
            std = float(max(np.std(local_scores), np.std(global_scores) / max(1.0, math.sqrt(support))))
            return mean, max(1e-6, std)

        return float(np.mean(global_scores)), max(1e-6, float(np.std(global_scores)))


def _pruning_reference_lcb(*, trials: list[Trial], risk_beta: float, min_completed: int = 5) -> float | None:
    completed = [
        t
        for t in trials
        if t.feasible and not t.stopped_early and t.period_fraction >= 1.0 and t.metrics.get("_scalar_score") is not None
    ]
    if len(completed) < min_completed:
        return None
    lcb_scores = []
    for trial in completed:
        trial_std = float((trial.uncertainty or {}).get("predicted_std", 0.0))
        lcb_scores.append(float(trial.metrics["_scalar_score"]) - risk_beta * trial_std)
    return max(lcb_scores, default=None)


def _build_uncertainty(
    *,
    predicted_mean: float | None,
    predicted_std: float | None,
    observed_scores: list[float],
    risk_beta: float,
) -> dict[str, float] | None:
    if predicted_mean is None and predicted_std is None and not observed_scores:
        return None
    running_mean = float(np.mean(observed_scores)) if observed_scores else float("nan")
    running_std = float(np.std(observed_scores)) if len(observed_scores) > 1 else 0.0
    total_std = math.sqrt(max(1e-9, running_std**2 + float(predicted_std or 0.0) ** 2))
    blended_mean = float(predicted_mean) if predicted_mean is not None else running_mean
    lcb = blended_mean - risk_beta * total_std
    return {
        "predicted_mean": blended_mean,
        "predicted_std": total_std,
        "running_mean": running_mean,
        "running_std": running_std,
        "lcb": lcb,
    }


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


class CMASampler:
    """Small CMA-ES-inspired sampler for mixed spaces."""

    def __init__(self, *, elite_fraction: float = 0.3, sigma_decay: float = 0.98, min_sigma: float = 0.03) -> None:
        self.elite_fraction = float(elite_fraction)
        self.sigma_decay = float(sigma_decay)
        self.min_sigma = float(min_sigma)

    def sample(self, *, trials: list[Trial], space: dict[str, SearchDimension], rng: np.random.Generator) -> dict[str, Any]:
        observed = [t for t in trials if not t.stopped_early and t.metrics.get("_scalar_score") is not None]
        if len(observed) < 5:
            return RandomSampler().sample(trials=trials, space=space, rng=rng)

        ranked = sorted(observed, key=lambda t: float(t.metrics.get("_scalar_score", -math.inf)), reverse=True)
        n_elites = max(2, int(math.ceil(len(ranked) * self.elite_fraction)))
        elites = ranked[:n_elites]

        sampled: dict[str, Any] = {}
        for key, dim in space.items():
            if isinstance(dim, DiscreteDimension):
                counts = {repr(v): 1.0 for v in dim.values}
                mapping = {repr(v): v for v in dim.values}
                for trial in elites:
                    value = repr(trial.params.get(key))
                    counts[value] = counts.get(value, 1.0) + 2.0
                labels = list(counts.keys())
                probs = np.array([counts[label] for label in labels], dtype=float)
                probs /= probs.sum()
                sampled[key] = mapping[str(rng.choice(labels, p=probs))]
                continue

            elite_values = np.array([float(trial.params.get(key, dim.low)) for trial in elites], dtype=float)
            center = float(np.mean(elite_values))
            spread = float(max(np.std(elite_values), self.min_sigma * (dim.high - dim.low))) * self.sigma_decay
            spread = max(self.min_sigma * (dim.high - dim.low), spread)
            value = float(rng.normal(loc=center, scale=spread))
            value = float(min(dim.high, max(dim.low, value)))
            if dim.step and dim.step > 0:
                n_steps = round((value - dim.low) / dim.step)
                value = dim.low + (n_steps * dim.step)
                value = float(min(dim.high, max(dim.low, value)))
            sampled[key] = value

        return sampled

    def predict(self, *, trials: list[Trial], space: dict[str, SearchDimension], params: dict[str, Any]) -> tuple[float, float] | None:
        return BayesianSampler().predict(trials=trials, space=space, params=params)


class GridSampler:
    """Deterministic cartesian traversal for discrete spaces."""

    def __init__(self) -> None:
        self._cursor = 0
        self._grid_cache: tuple[tuple[str, tuple[Any, ...]], ...] | None = None
        self._grid_rows: list[dict[str, Any]] = []

    def _rebuild_if_needed(self, space: dict[str, SearchDimension]) -> None:
        signature: tuple[tuple[str, tuple[Any, ...]], ...] = tuple(
            (key, tuple(dim.values) if isinstance(dim, DiscreteDimension) else tuple()) for key, dim in sorted(space.items())
        )
        if self._grid_cache == signature and self._grid_rows:
            return
        if any(not isinstance(dim, DiscreteDimension) for dim in space.values()):
            self._grid_cache = signature
            self._grid_rows = []
            self._cursor = 0
            return

        keys = list(space.keys())
        dims = [space[key] for key in keys]
        rows: list[dict[str, Any]] = []
        for values in itertools.product(*[dim.values for dim in dims]):
            rows.append({key: value for key, value in zip(keys, values)})
        self._grid_cache = signature
        self._grid_rows = rows
        self._cursor = 0

    def sample(self, *, trials: list[Trial], space: dict[str, SearchDimension], rng: np.random.Generator) -> dict[str, Any]:
        self._rebuild_if_needed(space)
        if not self._grid_rows:
            return RandomSampler().sample(trials=trials, space=space, rng=rng)
        sampled = dict(self._grid_rows[self._cursor % len(self._grid_rows)])
        self._cursor += 1
        return sampled

    def predict(self, *, trials: list[Trial], space: dict[str, SearchDimension], params: dict[str, Any]) -> tuple[float, float] | None:
        return None


def compute_scalar_score(
    metrics: dict[str, float],
    objectives: list[Objective],
    *,
    objective_weights: dict[str, float] | None = None,
) -> float:
    score = 0.0
    for obj in objectives:
        value = _resolve_metric_value(metrics, obj.name)
        if math.isnan(value):
            value = -math.inf if obj.sense == "maximize" else math.inf
        weight = float((objective_weights or {}).get(obj.name, 1.0))
        score += weight * (value if obj.sense == "maximize" else -value)
    return score


def compute_overfitting_penalty(
    *,
    metrics: dict[str, float],
    fold_scores: list[float],
    config: OverfittingPenaltyConfig,
) -> float:
    train_sharpe = float(metrics.get("train_sharpe", metrics.get("in_sample_sharpe", float("nan"))))
    validation_sharpe = float(metrics.get("validation_sharpe", metrics.get("oos_sharpe", float("nan"))))
    gap_penalty = 0.0
    if not math.isnan(train_sharpe) and not math.isnan(validation_sharpe):
        gap_penalty = max(0.0, train_sharpe - validation_sharpe) * float(config.train_validation_gap_weight)

    instability_penalty = (float(np.std(fold_scores)) if len(fold_scores) > 1 else 0.0) * float(config.fold_instability_weight)

    tail_risk_value = _resolve_metric_value(metrics, "tail_risk")
    tail_risk_penalty = 0.0 if math.isnan(tail_risk_value) else abs(float(tail_risk_value)) * float(config.tail_risk_weight)

    direct_penalty = max(0.0, float(metrics.get("probability_of_overfitting", 0.0)))
    return float(gap_penalty + instability_penalty + tail_risk_penalty + direct_penalty)


def _aggregate_walk_forward_metrics(
    metric_history: list[dict[str, float]],
    objectives: list[Objective],
    final_metrics: dict[str, float],
) -> dict[str, float]:
    if not metric_history:
        return dict(final_metrics)
    out = dict(final_metrics)
    for obj in objectives:
        values = [
            _resolve_metric_value(metrics=row, objective_name=obj.name)
            for row in metric_history
        ]
        finite_vals = [float(v) for v in values if math.isfinite(v)]
        if not finite_vals:
            continue
        out[f"wf_{obj.name}_mean"] = float(np.mean(finite_vals))
        out[f"wf_{obj.name}_worst"] = float(min(finite_vals) if obj.sense == "maximize" else max(finite_vals))
        out[f"wf_{obj.name}_final"] = float(finite_vals[-1])
    return out


def _load_prior_trials(
    *,
    history_path: Path | None,
    strategy_key: str | None,
    prior_strategy_keys: list[str] | None,
) -> list[Trial]:
    if history_path is None or not history_path.exists():
        return []
    accepted = {str(strategy_key)} if strategy_key else set()
    accepted.update(str(item) for item in (prior_strategy_keys or []))
    loaded: list[Trial] = []
    for raw_line in history_path.read_text(encoding="utf-8").splitlines():
        if not raw_line.strip():
            continue
        try:
            row = json.loads(raw_line)
        except json.JSONDecodeError:
            continue
        if accepted and str(row.get("strategy_key", "")) not in accepted:
            continue
        params = row.get("params") if isinstance(row.get("params"), dict) else None
        metrics = row.get("metrics") if isinstance(row.get("metrics"), dict) else None
        if not params or not metrics:
            continue
        loaded.append(
            Trial(
                trial_id=int(row.get("trial_id", -len(loaded) - 1)),
                params={k: v for k, v in params.items()},
                metrics={k: float(v) for k, v in metrics.items() if isinstance(v, (int, float))},
                feasible=bool(row.get("feasible", True)),
                stopped_early=bool(row.get("stopped_early", False)),
                period_fraction=float(row.get("period_fraction", 1.0)),
            )
        )
    return loaded


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
    objective_weights: dict[str, float] | None = None,
    overfitting_penalty: OverfittingPenaltyConfig | None = None,
    use_walk_forward_objective_metrics: bool = True,
    history_path: Path | None = None,
    strategy_key: str | None = None,
    prior_strategy_keys: list[str] | None = None,
    enable_pruning: bool = True,
    prune_on_constraint_violation: bool = True,
    prune_on_lcb: bool = True,
    min_completed_for_pruning: int = 5,
) -> dict[str, Any]:
    risk_beta = 1.5
    output_dir.mkdir(parents=True, exist_ok=True)
    normalized_space = _normalize_space(space)
    fractions = sorted(set(partial_period_fractions or [1.0]))
    if fractions[-1] != 1.0:
        fractions.append(1.0)

    rng = np.random.default_rng(seed)
    trials: list[Trial] = []
    seeded_trials = _load_prior_trials(
        history_path=history_path,
        strategy_key=strategy_key,
        prior_strategy_keys=prior_strategy_keys,
    )
    jsonl_path = output_dir / "trials.jsonl"
    penalty_cfg = overfitting_penalty or OverfittingPenaltyConfig()

    with jsonl_path.open("w", encoding="utf-8") as handle:
        for trial_id in range(int(n_trials)):
            context_trials = [*seeded_trials, *trials]
            params = sampler.sample(trials=context_trials, space=normalized_space, rng=rng)
            last_metrics: dict[str, float] = {}
            metric_history: list[dict[str, float]] = []
            feasible = False
            stopped_early = False
            period_fraction = 1.0
            fold_scores: list[float] = []
            pred = sampler.predict(trials=context_trials, space=normalized_space, params=params)
            pred_mean = None if pred is None else float(pred[0])
            pred_std = None if pred is None else float(pred[1])
            uncertainty: dict[str, float] | None = None
            for period_fraction in fractions:
                last_metrics = {k: float(v) for k, v in evaluate(params, float(period_fraction)).items()}
                metric_history.append(dict(last_metrics))
                scoring_metrics = (
                    _aggregate_walk_forward_metrics(metric_history, objectives, last_metrics)
                    if use_walk_forward_objective_metrics
                    else last_metrics
                )
                raw_scalar = compute_scalar_score(scoring_metrics, objectives, objective_weights=objective_weights)
                overfit_pen = compute_overfitting_penalty(metrics=scoring_metrics, fold_scores=fold_scores, config=penalty_cfg)
                last_metrics.update(scoring_metrics)
                last_metrics["_overfit_penalty"] = float(overfit_pen)
                last_metrics["_raw_scalar_score"] = float(raw_scalar)
                last_metrics["_scalar_score"] = float(raw_scalar - overfit_pen)
                fold_scores.append(float(last_metrics["_scalar_score"]))
                feasible = check_constraints(last_metrics, constraints)
                uncertainty = _build_uncertainty(
                    predicted_mean=pred_mean,
                    predicted_std=pred_std,
                    observed_scores=fold_scores,
                    risk_beta=risk_beta,
                )
                if enable_pruning and prune_on_constraint_violation and not feasible and period_fraction < 1.0:
                    stopped_early = True
                    break
                if enable_pruning and prune_on_lcb and period_fraction < 1.0:
                    reference_lcb = _pruning_reference_lcb(trials=trials, risk_beta=risk_beta, min_completed=max(1, int(min_completed_for_pruning)))
                    current_lcb = None if uncertainty is None else uncertainty.get("lcb")
                    if reference_lcb is not None and current_lcb is not None and float(current_lcb) < float(reference_lcb):
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
                uncertainty=uncertainty,
                fold_scores=list(fold_scores),
                oos_score_vector=list(fold_scores),
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
                        "oos_score_vector": trial.oos_score_vector,
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
    oos_vectors = [list(t.oos_score_vector or t.fold_scores or []) for t in trials if t.feasible and not t.stopped_early]
    corrected_significance = compute_corrected_pvalues(score_vectors=oos_vectors, seed=seed)
    (output_dir / "oos_score_vectors.json").write_text(
        json.dumps(
            [
                {
                    "trial_id": t.trial_id,
                    "params": t.params,
                    "oos_score_vector": list(t.oos_score_vector or t.fold_scores or []),
                }
                for t in trials
            ],
            indent=2,
            sort_keys=True,
        ),
        encoding="utf-8",
    )
    (output_dir / "corrected_significance.json").write_text(
        json.dumps(corrected_significance, indent=2, sort_keys=True),
        encoding="utf-8",
    )
    (output_dir / "artifact_metadata.json").write_text(
        json.dumps(
            {
                "schema_version": "1.0",
                "run_type": "optimization_trials",
                "random_seeds": {"run_seed": int(seed), "numpy_random_seed": int(seed)},
                "uncertainty_diagnostics": {
                    "risk_beta": risk_beta,
                    "fields": ["predicted_mean", "predicted_std", "running_mean", "running_std", "lcb"],
                },
                "seeded_prior_trials": len(seeded_trials),
                "strategy_key": strategy_key,
                "prior_strategy_keys": list(prior_strategy_keys or []),
                "pruning": {
                    "enabled": bool(enable_pruning),
                    "prune_on_constraint_violation": bool(prune_on_constraint_violation),
                    "prune_on_lcb": bool(prune_on_lcb),
                    "min_completed_for_pruning": int(min_completed_for_pruning),
                },
            },
            indent=2,
            sort_keys=True,
        ),
        encoding="utf-8",
    )

    if history_path is not None:
        history_path.parent.mkdir(parents=True, exist_ok=True)
        with history_path.open("a", encoding="utf-8") as handle:
            for trial in trials:
                handle.write(
                    json.dumps(
                        {
                            "strategy_key": strategy_key,
                            "trial_id": int(trial.trial_id),
                            "params": trial.params,
                            "metrics": trial.metrics,
                            "feasible": bool(trial.feasible),
                            "stopped_early": bool(trial.stopped_early),
                            "period_fraction": float(trial.period_fraction),
                        },
                        sort_keys=True,
                    )
                    + "\n"
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
        "corrected_significance": corrected_significance,
        "trials_path": str(jsonl_path),
        "pareto_path": str(output_dir / "pareto_frontier.json"),
    }

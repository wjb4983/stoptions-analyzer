from __future__ import annotations

from dataclasses import dataclass, field

import numpy as np

from ...core.contracts import PredictionResult

try:
    import torch
    from torch import nn

    _HAS_TORCH = True
except Exception:  # pragma: no cover - optional dependency fallback
    torch = None
    nn = None
    _HAS_TORCH = False


_REQUIRED_INPUTS = ("returns",)


@dataclass
class CrossAssetGraphModel:
    """
    Cross-asset graph model with dynamic connectivity and node-level outputs.

    The model builds a dynamic adjacency matrix from:
      1. Cross-sectional return correlation,
      2. Lead/lag relationships,
      3. Sector membership edges,
      4. Macro exposure similarity edges.

    If torch is available and ``use_deep_stack`` is True, a compact graph-aware
    deep head is trained to produce node-level alpha and uncertainty.
    Otherwise, the model falls back to sparse linear graph filters.
    """

    use_deep_stack: bool = True
    hidden_dim: int = 24
    epochs: int = 200
    learning_rate: float = 1e-2
    ridge: float = 1e-3

    corr_weight: float = 0.55
    lead_lag_weight: float = 0.25
    sector_edge_weight: float = 0.15
    macro_edge_weight: float = 0.05
    sparsity_top_k: int = 8
    seed: int = 11

    _rng: np.random.Generator = field(init=False, repr=False)
    _adjacency: np.ndarray | None = field(default=None, init=False, repr=False)
    _linear_w: np.ndarray | None = field(default=None, init=False, repr=False)
    _linear_resid_scale: np.ndarray | None = field(default=None, init=False, repr=False)
    _deep_model: nn.Module | None = field(default=None, init=False, repr=False)

    def __post_init__(self) -> None:
        if self.hidden_dim <= 0:
            raise ValueError("hidden_dim must be > 0")
        if self.epochs <= 0:
            raise ValueError("epochs must be > 0")
        if self.learning_rate <= 0:
            raise ValueError("learning_rate must be > 0")
        if self.ridge < 0:
            raise ValueError("ridge must be >= 0")
        if self.sparsity_top_k <= 0:
            raise ValueError("sparsity_top_k must be > 0")
        self._rng = np.random.default_rng(self.seed)

    @property
    def using_deep_stack(self) -> bool:
        return bool(self.use_deep_stack and _HAS_TORCH)

    def fit(self, features: dict[str, np.ndarray], labels: np.ndarray) -> "CrossAssetGraphModel":
        returns = self._require_returns(features)
        y = self._normalize_labels(labels=labels, n_samples=returns.shape[0], n_nodes=returns.shape[1])

        adjacency = self._build_dynamic_adjacency(features)
        self._adjacency = adjacency
        graph_features = self._graph_filter_bank(returns, adjacency)

        if self.using_deep_stack:
            self._fit_deep(graph_features, y)
            self._linear_w = None
            self._linear_resid_scale = None
        else:
            self._fit_linear(graph_features, y)
            self._deep_model = None

        return self

    def predict(self, features: dict[str, np.ndarray]) -> PredictionResult:
        alpha, uncertainty = self.predict_alpha_uncertainty(features)
        return PredictionResult(
            predictions=alpha,
            uncertainty=uncertainty,
            metadata={
                "backend": "deep" if self.using_deep_stack else "sparse_linear_filter",
                "n_nodes": int(alpha.shape[1]),
            },
        )

    def predict_alpha_uncertainty(self, features: dict[str, np.ndarray]) -> tuple[np.ndarray, np.ndarray]:
        returns = self._require_returns(features)
        adjacency = self._adjacency if self._adjacency is not None else self._build_dynamic_adjacency(features)
        graph_features = self._graph_filter_bank(returns, adjacency)

        if self.using_deep_stack and self._deep_model is not None:
            xt = torch.from_numpy(graph_features.astype(np.float32))
            with torch.no_grad():
                out = self._deep_model(xt)
            alpha = out[..., 0].numpy()
            uncertainty = np.log1p(np.exp(out[..., 1].numpy())) + 1e-6
            return alpha, uncertainty

        if self._linear_w is None:
            self._fit_linear(graph_features, np.zeros((returns.shape[0], returns.shape[1]), dtype=float))

        assert self._linear_w is not None and self._linear_resid_scale is not None
        alpha = np.einsum("snf,nf->sn", graph_features, self._linear_w)
        uncertainty = np.broadcast_to(self._linear_resid_scale[None, :], alpha.shape)
        return alpha, uncertainty

    def _fit_deep(self, graph_features: np.ndarray, labels: np.ndarray) -> None:
        assert nn is not None and torch is not None
        torch.manual_seed(self.seed)

        model = nn.Sequential(
            nn.Linear(graph_features.shape[2], self.hidden_dim),
            nn.ReLU(),
            nn.Linear(self.hidden_dim, 2),
        )
        optimizer = torch.optim.Adam(model.parameters(), lr=self.learning_rate)

        x = torch.from_numpy(graph_features.astype(np.float32))
        y = torch.from_numpy(labels.astype(np.float32))

        for _ in range(self.epochs):
            optimizer.zero_grad()
            pred = model(x)
            mu = pred[..., 0]
            raw_sigma = pred[..., 1]
            sigma = torch.nn.functional.softplus(raw_sigma) + 1e-6
            loss = torch.mean(0.5 * ((y - mu) / sigma) ** 2 + torch.log(sigma))
            loss.backward()
            optimizer.step()

        self._deep_model = model.eval()

    def _fit_linear(self, graph_features: np.ndarray, labels: np.ndarray) -> None:
        n_samples, n_nodes, n_filters = graph_features.shape
        x = graph_features.reshape(n_samples, n_nodes, n_filters)
        w = np.zeros((n_nodes, n_filters), dtype=float)
        resid_scale = np.zeros(n_nodes, dtype=float)

        eye = np.eye(n_filters)
        for node in range(n_nodes):
            x_node = x[:, node, :]
            y_node = labels[:, node]
            gram = x_node.T @ x_node + self.ridge * eye
            rhs = x_node.T @ y_node
            w[node] = np.linalg.solve(gram, rhs)
            resid = y_node - x_node @ w[node]
            resid_scale[node] = float(np.sqrt(np.mean(resid**2)) + 1e-6)

        self._linear_w = w
        self._linear_resid_scale = resid_scale

    def _graph_filter_bank(self, returns: np.ndarray, adjacency: np.ndarray) -> np.ndarray:
        a_sparse = self._sparsify_top_k(adjacency, top_k=min(self.sparsity_top_k, adjacency.shape[1]))
        a2_sparse = self._sparsify_top_k(a_sparse @ a_sparse, top_k=min(self.sparsity_top_k, adjacency.shape[1]))

        f0 = returns
        f1 = returns @ a_sparse
        f2 = returns @ a2_sparse
        return np.stack([f0, f1, f2], axis=2)

    def _build_dynamic_adjacency(self, features: dict[str, np.ndarray]) -> np.ndarray:
        returns = self._require_returns(features)
        n_nodes = returns.shape[1]

        corr = np.corrcoef(returns, rowvar=False)
        corr = np.nan_to_num(corr, nan=0.0, posinf=0.0, neginf=0.0)
        corr = np.abs(corr)

        lead_lag = self._lead_lag_matrix(returns)

        sector = self._sector_edges(features, n_nodes=n_nodes)
        macro = self._macro_edges(features, n_nodes=n_nodes)

        adj = (
            self.corr_weight * corr
            + self.lead_lag_weight * lead_lag
            + self.sector_edge_weight * sector
            + self.macro_edge_weight * macro
        )
        np.fill_diagonal(adj, 0.0)

        row_sum = adj.sum(axis=1, keepdims=True)
        row_sum[row_sum == 0.0] = 1.0
        return adj / row_sum

    def _lead_lag_matrix(self, returns: np.ndarray) -> np.ndarray:
        if returns.shape[0] < 2:
            return np.zeros((returns.shape[1], returns.shape[1]), dtype=float)

        lead = returns[1:, :]
        lag = returns[:-1, :]
        centered_lead = lead - np.mean(lead, axis=0, keepdims=True)
        centered_lag = lag - np.mean(lag, axis=0, keepdims=True)

        cov = centered_lag.T @ centered_lead / max(1, lead.shape[0] - 1)
        std_lag = np.std(lag, axis=0, ddof=1)
        std_lead = np.std(lead, axis=0, ddof=1)
        denom = np.outer(std_lag, std_lead)
        denom[denom == 0.0] = 1.0

        ll = np.abs(cov / denom)
        ll = np.nan_to_num(ll, nan=0.0, posinf=0.0, neginf=0.0)
        np.fill_diagonal(ll, 0.0)
        return ll

    def _sector_edges(self, features: dict[str, np.ndarray], *, n_nodes: int) -> np.ndarray:
        sector_ids = features.get("sector_ids")
        if sector_ids is None:
            return np.zeros((n_nodes, n_nodes), dtype=float)

        sector = np.asarray(sector_ids)
        if sector.ndim != 1 or sector.shape[0] != n_nodes:
            raise ValueError("sector_ids must have shape (n_nodes,)")

        equal = sector[:, None] == sector[None, :]
        out = equal.astype(float)
        np.fill_diagonal(out, 0.0)
        return out

    def _macro_edges(self, features: dict[str, np.ndarray], *, n_nodes: int) -> np.ndarray:
        macro = features.get("macro_exposures")
        if macro is None:
            return np.zeros((n_nodes, n_nodes), dtype=float)

        x = np.asarray(macro, dtype=float)
        if x.ndim != 2 or x.shape[0] != n_nodes:
            raise ValueError("macro_exposures must have shape (n_nodes, n_factors)")

        norms = np.linalg.norm(x, axis=1, keepdims=True)
        norms[norms == 0.0] = 1.0
        x_unit = x / norms
        sim = np.clip(x_unit @ x_unit.T, 0.0, 1.0)
        np.fill_diagonal(sim, 0.0)
        return sim

    def _sparsify_top_k(self, mat: np.ndarray, *, top_k: int) -> np.ndarray:
        if top_k >= mat.shape[1]:
            return mat

        out = np.zeros_like(mat)
        idx = np.argpartition(mat, kth=-top_k, axis=1)[:, -top_k:]
        rows = np.arange(mat.shape[0])[:, None]
        out[rows, idx] = mat[rows, idx]

        row_sum = out.sum(axis=1, keepdims=True)
        row_sum[row_sum == 0.0] = 1.0
        return out / row_sum

    def _require_returns(self, features: dict[str, np.ndarray]) -> np.ndarray:
        missing = [name for name in _REQUIRED_INPUTS if name not in features]
        if missing:
            raise ValueError(f"Missing required inputs: {missing}")

        returns = np.asarray(features["returns"], dtype=float)
        if returns.ndim != 2:
            raise ValueError("returns must have shape (n_samples, n_nodes)")
        return returns

    def _normalize_labels(self, *, labels: np.ndarray, n_samples: int, n_nodes: int) -> np.ndarray:
        y = np.asarray(labels, dtype=float)
        if y.ndim == 1:
            if y.shape[0] != n_samples:
                raise ValueError(f"labels must have shape ({n_samples},) or ({n_samples}, {n_nodes})")
            y = np.repeat(y[:, None], n_nodes, axis=1)
        elif y.ndim == 2:
            if y.shape != (n_samples, n_nodes):
                raise ValueError(f"labels must have shape ({n_samples}, {n_nodes})")
        else:
            raise ValueError("labels must be 1D or 2D")
        return y

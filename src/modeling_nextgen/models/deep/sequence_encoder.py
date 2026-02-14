from __future__ import annotations

from dataclasses import dataclass, field

import numpy as np

from ...core.contracts import PredictionResult


_REQUIRED_INPUTS = ("order_flow", "imbalance", "gamma_proxy", "realized_vol")


@dataclass
class SequenceEncoder:
    """
    Lightweight sequence encoder with pluggable backbones (TCN/LSTM/Transformer).

    Inputs are expected in a feature dictionary containing at minimum:
      - order_flow
      - imbalance
      - gamma_proxy
      - realized_vol

    The model emits horizon-conditioned class probabilities for downstream
    signal-engine consumption.
    """

    architecture: str = "tcn"
    horizons: tuple[str, ...] = ("short", "medium", "long")
    n_classes: int = 3
    hidden_dim: int = 32
    learning_rate: float = 5e-2
    epochs: int = 250
    seed: int = 13

    _rng: np.random.Generator = field(init=False, repr=False)
    _input_channels: int | None = field(default=None, init=False, repr=False)
    _proj_w: np.ndarray | None = field(default=None, init=False, repr=False)
    _proj_b: np.ndarray | None = field(default=None, init=False, repr=False)
    _lstm_params: dict[str, np.ndarray] | None = field(default=None, init=False, repr=False)
    _attn_params: dict[str, np.ndarray] | None = field(default=None, init=False, repr=False)
    _head_w: np.ndarray | None = field(default=None, init=False, repr=False)
    _head_b: np.ndarray | None = field(default=None, init=False, repr=False)

    def __post_init__(self) -> None:
        if self.architecture not in {"tcn", "lstm", "transformer"}:
            raise ValueError("architecture must be one of: tcn, lstm, transformer")
        if self.n_classes < 2:
            raise ValueError("n_classes must be >= 2")
        if self.hidden_dim <= 0:
            raise ValueError("hidden_dim must be > 0")
        if self.epochs <= 0:
            raise ValueError("epochs must be > 0")
        if self.learning_rate <= 0:
            raise ValueError("learning_rate must be > 0")

        self._rng = np.random.default_rng(self.seed)

    def fit(self, features: dict[str, np.ndarray], labels: np.ndarray) -> "SequenceEncoder":
        seq = self._stack_features(features)
        self._init_encoder(seq.shape[2])

        encoded = self._encode(seq)
        labels_2d = self._normalize_labels(labels=labels, n_samples=seq.shape[0])
        n_samples = encoded.shape[0]

        assert self._head_w is not None and self._head_b is not None
        for _ in range(self.epochs):
            logits = np.einsum("nh,hkc->nkc", encoded, self._head_w) + self._head_b
            probs = _softmax_last(logits)

            one_hot = _one_hot_3d(labels_2d, self.n_classes)
            grad_logits = (probs - one_hot) / n_samples

            grad_w = np.einsum("nh,nkc->hkc", encoded, grad_logits)
            grad_b = np.sum(grad_logits, axis=0)

            self._head_w -= self.learning_rate * grad_w
            self._head_b -= self.learning_rate * grad_b

        return self

    def predict_proba(self, features: dict[str, np.ndarray], horizon: str | None = None) -> np.ndarray:
        seq = self._stack_features(features)
        if self._head_w is None or self._head_b is None:
            self._init_encoder(seq.shape[2])

        encoded = self._encode(seq)
        logits = np.einsum("nh,hkc->nkc", encoded, self._head_w) + self._head_b
        probs = _softmax_last(logits)

        if horizon is None:
            return probs
        idx = self._horizon_index(horizon)
        return probs[:, idx, :]

    def predict(self, features: dict[str, np.ndarray]) -> PredictionResult:
        probs = self.predict_proba(features)
        preds = np.argmax(probs, axis=2)
        return PredictionResult(
            predictions=preds,
            probabilities=probs,
            metadata={"horizons": list(self.horizons), "architecture": self.architecture},
        )

    def _normalize_labels(self, *, labels: np.ndarray, n_samples: int) -> np.ndarray:
        y = np.asarray(labels, dtype=int)
        if y.ndim == 1:
            if y.shape[0] != n_samples:
                raise ValueError(f"labels must have shape ({n_samples},) or ({n_samples}, {len(self.horizons)})")
            y = np.tile(y[:, None], (1, len(self.horizons)))
        elif y.ndim == 2:
            if y.shape != (n_samples, len(self.horizons)):
                raise ValueError(f"labels must have shape ({n_samples}, {len(self.horizons)})")
        else:
            raise ValueError("labels must be 1D or 2D")

        if np.min(y) < 0 or np.max(y) >= self.n_classes:
            raise ValueError(f"labels values must be in [0, {self.n_classes - 1}]")
        return y

    def _horizon_index(self, horizon: str) -> int:
        try:
            return self.horizons.index(horizon)
        except ValueError as exc:
            raise ValueError(f"Unknown horizon {horizon!r}; available={self.horizons}") from exc

    def _stack_features(self, features: dict[str, np.ndarray]) -> np.ndarray:
        missing = [name for name in _REQUIRED_INPUTS if name not in features]
        if missing:
            raise ValueError(f"Missing required inputs: {missing}")

        converted = [self._as_sequence(np.asarray(features[name], dtype=float)) for name in _REQUIRED_INPUTS]
        n_samples = converted[0].shape[0]
        seq_len = max(arr.shape[1] for arr in converted)

        broadcasted: list[np.ndarray] = []
        for arr in converted:
            if arr.shape[0] != n_samples:
                raise ValueError("All feature arrays must have same number of samples")
            if arr.shape[1] != seq_len:
                if arr.shape[1] == 1:
                    arr = np.repeat(arr, seq_len, axis=1)
                else:
                    raise ValueError("Sequence lengths must match or be broadcastable from length=1")
            broadcasted.append(arr)

        return np.concatenate(broadcasted, axis=2)

    def _as_sequence(self, arr: np.ndarray) -> np.ndarray:
        if arr.ndim == 1:
            return arr[:, None, None]
        if arr.ndim == 2:
            return arr[:, :, None]
        if arr.ndim == 3:
            return arr
        raise ValueError("Feature arrays must be 1D, 2D, or 3D")

    def _init_encoder(self, input_channels: int) -> None:
        if self._input_channels == input_channels and self._head_w is not None:
            return

        self._input_channels = input_channels
        self._proj_w = self._rng.normal(0.0, 0.08, size=(input_channels, self.hidden_dim))
        self._proj_b = np.zeros(self.hidden_dim, dtype=float)

        self._lstm_params = {
            "W_i": self._rng.normal(0.0, 0.08, size=(input_channels, self.hidden_dim)),
            "U_i": self._rng.normal(0.0, 0.08, size=(self.hidden_dim, self.hidden_dim)),
            "W_f": self._rng.normal(0.0, 0.08, size=(input_channels, self.hidden_dim)),
            "U_f": self._rng.normal(0.0, 0.08, size=(self.hidden_dim, self.hidden_dim)),
            "W_o": self._rng.normal(0.0, 0.08, size=(input_channels, self.hidden_dim)),
            "U_o": self._rng.normal(0.0, 0.08, size=(self.hidden_dim, self.hidden_dim)),
            "W_g": self._rng.normal(0.0, 0.08, size=(input_channels, self.hidden_dim)),
            "U_g": self._rng.normal(0.0, 0.08, size=(self.hidden_dim, self.hidden_dim)),
            "b_i": np.zeros(self.hidden_dim),
            "b_f": np.zeros(self.hidden_dim),
            "b_o": np.zeros(self.hidden_dim),
            "b_g": np.zeros(self.hidden_dim),
        }
        self._attn_params = {
            "W_q": self._rng.normal(0.0, 0.08, size=(input_channels, self.hidden_dim)),
            "W_k": self._rng.normal(0.0, 0.08, size=(input_channels, self.hidden_dim)),
            "W_v": self._rng.normal(0.0, 0.08, size=(input_channels, self.hidden_dim)),
        }

        self._head_w = self._rng.normal(0.0, 0.08, size=(self.hidden_dim, len(self.horizons), self.n_classes))
        self._head_b = np.zeros((len(self.horizons), self.n_classes), dtype=float)

    def _encode(self, seq: np.ndarray) -> np.ndarray:
        if self.architecture == "tcn":
            return self._encode_tcn(seq)
        if self.architecture == "lstm":
            return self._encode_lstm(seq)
        return self._encode_transformer(seq)

    def _encode_tcn(self, seq: np.ndarray) -> np.ndarray:
        assert self._proj_w is not None and self._proj_b is not None

        pooled_states: list[np.ndarray] = []
        for dilation in (1, 2, 4):
            lag = min(dilation, seq.shape[1] - 1)
            shifted = np.roll(seq, shift=lag, axis=1)
            shifted[:, :lag, :] = seq[:, :1, :]
            conv_like = 0.6 * seq + 0.4 * shifted
            pooled_states.append(np.mean(conv_like, axis=1))

        merged = np.mean(np.stack(pooled_states, axis=1), axis=1)
        hidden = np.tanh(merged @ self._proj_w + self._proj_b)
        return hidden

    def _encode_lstm(self, seq: np.ndarray) -> np.ndarray:
        assert self._lstm_params is not None

        h = np.zeros((seq.shape[0], self.hidden_dim), dtype=float)
        c = np.zeros_like(h)
        p = self._lstm_params

        for t in range(seq.shape[1]):
            x_t = seq[:, t, :]
            i = _sigmoid(x_t @ p["W_i"] + h @ p["U_i"] + p["b_i"])
            f = _sigmoid(x_t @ p["W_f"] + h @ p["U_f"] + p["b_f"])
            o = _sigmoid(x_t @ p["W_o"] + h @ p["U_o"] + p["b_o"])
            g = np.tanh(x_t @ p["W_g"] + h @ p["U_g"] + p["b_g"])

            c = f * c + i * g
            h = o * np.tanh(c)

        return h

    def _encode_transformer(self, seq: np.ndarray) -> np.ndarray:
        assert self._attn_params is not None

        q = seq @ self._attn_params["W_q"]
        k = seq @ self._attn_params["W_k"]
        v = seq @ self._attn_params["W_v"]

        scale = np.sqrt(float(self.hidden_dim))
        scores = np.einsum("nth,nsh->nts", q, k) / scale
        weights = _softmax_last(scores)
        attended = np.einsum("nts,nsh->nth", weights, v)
        return np.mean(attended, axis=1)


def _sigmoid(x: np.ndarray) -> np.ndarray:
    return 1.0 / (1.0 + np.exp(-x))


def _softmax_last(x: np.ndarray) -> np.ndarray:
    shifted = x - np.max(x, axis=-1, keepdims=True)
    exp = np.exp(shifted)
    return exp / np.sum(exp, axis=-1, keepdims=True)


def _one_hot_3d(y: np.ndarray, n_classes: int) -> np.ndarray:
    out = np.zeros((y.shape[0], y.shape[1], n_classes), dtype=float)
    rows = np.arange(y.shape[0])[:, None]
    cols = np.arange(y.shape[1])[None, :]
    out[rows, cols, y] = 1.0
    return out

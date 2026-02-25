from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
from typing import Any, Callable

import yaml

NN_PRESETS_PATH = Path("config/nn_presets.json")
NN_EXPORT_DIR = Path("data/nn_architectures")


def default_architecture_spec() -> dict[str, Any]:
    return {
        "schema_version": 1,
        "layers": [
            {"type": "Dense", "units": 64, "activation": "relu"},
            {"type": "Dropout", "rate": 0.2},
            {"type": "Dense", "units": 32, "activation": "relu"},
            {"type": "Dense", "units": 1, "activation": "sigmoid"},
        ],
        "optimizer": {"name": "adam", "learning_rate": 0.001, "weight_decay": 0.0},
        "loss": {"name": "binary_cross_entropy"},
        "scheduler": {"name": "none"},
        "training": {
            "batch_size": 32,
            "epochs": 50,
            "early_stopping": {"enabled": True, "patience": 6, "min_delta": 0.0001},
        },
    }


def _as_float(value: Any, fallback: float) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return fallback


def _as_int(value: Any, fallback: int) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return fallback


def normalize_architecture_spec(payload: dict[str, Any] | None) -> dict[str, Any]:
    base = default_architecture_spec()
    if not isinstance(payload, dict):
        return base

    normalized = dict(base)
    layers = payload.get("layers")
    normalized_layers: list[dict[str, Any]] = []
    if isinstance(layers, list):
        for layer in layers:
            if isinstance(layer, dict) and str(layer.get("type", "")).strip():
                normalized_layers.append(dict(layer))
    normalized["layers"] = normalized_layers or base["layers"]

    for section in ("optimizer", "loss", "scheduler", "training"):
        section_value = payload.get(section)
        if isinstance(section_value, dict):
            normalized[section] = {**dict(base[section]), **dict(section_value)}

    training = dict(normalized["training"])
    early = training.get("early_stopping")
    if isinstance(early, dict):
        training["early_stopping"] = {
            "enabled": bool(early.get("enabled", True)),
            "patience": _as_int(early.get("patience", 6), 6),
            "min_delta": _as_float(early.get("min_delta", 0.0001), 0.0001),
        }
    else:
        training["early_stopping"] = dict(base["training"]["early_stopping"])
    training["batch_size"] = max(1, _as_int(training.get("batch_size", 32), 32))
    training["epochs"] = max(1, _as_int(training.get("epochs", 50), 50))
    normalized["training"] = training
    return normalized


def load_nn_presets() -> dict[str, dict[str, Any]]:
    if not NN_PRESETS_PATH.exists():
        return {}
    try:
        payload = json.loads(NN_PRESETS_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}
    presets = payload.get("presets", {}) if isinstance(payload, dict) else {}
    if not isinstance(presets, dict):
        return {}
    return {
        str(key): normalize_architecture_spec(value)
        for key, value in presets.items()
        if isinstance(key, str) and isinstance(value, dict)
    }


def save_nn_presets(presets: dict[str, dict[str, Any]]) -> None:
    payload = {
        "default_preset": next(iter(presets), ""),
        "presets": {name: normalize_architecture_spec(spec) for name, spec in presets.items()},
    }
    NN_PRESETS_PATH.parent.mkdir(parents=True, exist_ok=True)
    NN_PRESETS_PATH.write_text(json.dumps(payload, indent=2, sort_keys=True), encoding="utf-8")


def architecture_to_yaml(payload: dict[str, Any]) -> str:
    return yaml.safe_dump(normalize_architecture_spec(payload), sort_keys=False)


def architecture_to_python(payload: dict[str, Any], *, function_name: str = "build_model_spec") -> str:
    spec_literal = repr(normalize_architecture_spec(payload))
    return (
        '"""Generated neural network architecture spec for regime training."""\n\n'
        "from __future__ import annotations\n\n"
        "from typing import Any\n\n"
        f"def {function_name}() -> dict[str, Any]:\n"
        f"    return {spec_literal}\n"
    )


class NeuralNetworkDesignerPage(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Misc,
        *,
        initial_spec: dict[str, Any] | None,
        on_save: Callable[[dict[str, Any]], None],
        custom_presets: dict[str, dict[str, Any]] | None = None,
    ) -> None:
        super().__init__(parent)
        self.title("Neural Network Designer")
        self.geometry("860x620")
        self.transient(parent)

        self._on_save = on_save
        self._built_in_presets = load_nn_presets()
        self._custom_presets = {k: normalize_architecture_spec(v) for k, v in (custom_presets or {}).items()}
        self._spec = normalize_architecture_spec(initial_spec)

        self.layer_type_var = tk.StringVar(value="Dense")
        self.preset_var = tk.StringVar(value="")
        self.layers_var = tk.StringVar(value="")

        self.optimizer_name_var = tk.StringVar(value=str(self._spec["optimizer"].get("name", "adam")))
        self.optimizer_lr_var = tk.StringVar(value=str(self._spec["optimizer"].get("learning_rate", 0.001)))
        self.optimizer_wd_var = tk.StringVar(value=str(self._spec["optimizer"].get("weight_decay", 0.0)))

        self.loss_name_var = tk.StringVar(value=str(self._spec["loss"].get("name", "binary_cross_entropy")))
        self.scheduler_name_var = tk.StringVar(value=str(self._spec["scheduler"].get("name", "none")))

        self.batch_size_var = tk.StringVar(value=str(self._spec["training"].get("batch_size", 32)))
        self.epochs_var = tk.StringVar(value=str(self._spec["training"].get("epochs", 50)))
        self.early_stop_var = tk.BooleanVar(value=bool(self._spec["training"]["early_stopping"].get("enabled", True)))
        self.early_patience_var = tk.StringVar(value=str(self._spec["training"]["early_stopping"].get("patience", 6)))

        self._build_layout()
        self._refresh_layers_text()

    def _build_layout(self) -> None:
        root = ttk.Frame(self, padding=10)
        root.pack(fill="both", expand=True)

        preset_row = ttk.Frame(root)
        preset_row.pack(fill="x", pady=(0, 6))
        ttk.Label(preset_row, text="Preset").pack(side="left")
        preset_names = ["<built-in> " + name for name in sorted(self._built_in_presets)] + [
            "<custom> " + name for name in sorted(self._custom_presets)
        ]
        ttk.Combobox(preset_row, textvariable=self.preset_var, values=preset_names, state="readonly", width=34).pack(side="left", padx=6)
        ttk.Button(preset_row, text="Load preset", command=self._load_selected_preset).pack(side="left", padx=4)
        ttk.Button(preset_row, text="Save as preset", command=self._save_as_preset).pack(side="left", padx=4)

        layers = ttk.LabelFrame(root, text="Layer stack builder", padding=8)
        layers.pack(fill="x", pady=6)
        ttk.Label(layers, text="Layer type").grid(row=0, column=0, sticky="w")
        ttk.Combobox(layers, textvariable=self.layer_type_var, state="readonly", values=["Dense", "Dropout", "Norm", "Activation"], width=16).grid(row=0, column=1, sticky="w", padx=6)
        ttk.Button(layers, text="Add layer", command=self._add_layer).grid(row=0, column=2, padx=6)
        ttk.Button(layers, text="Remove last", command=self._remove_last_layer).grid(row=0, column=3, padx=6)
        ttk.Label(layers, textvariable=self.layers_var, justify="left", wraplength=780).grid(row=1, column=0, columnspan=4, sticky="w", pady=(8, 0))

        settings = ttk.LabelFrame(root, text="Optimizer / loss / scheduler", padding=8)
        settings.pack(fill="x", pady=6)
        ttk.Label(settings, text="Optimizer").grid(row=0, column=0, sticky="w")
        ttk.Entry(settings, textvariable=self.optimizer_name_var, width=18).grid(row=0, column=1, sticky="w", padx=4)
        ttk.Label(settings, text="Learning rate").grid(row=0, column=2, sticky="w")
        ttk.Entry(settings, textvariable=self.optimizer_lr_var, width=12).grid(row=0, column=3, sticky="w", padx=4)
        ttk.Label(settings, text="Weight decay").grid(row=0, column=4, sticky="w")
        ttk.Entry(settings, textvariable=self.optimizer_wd_var, width=12).grid(row=0, column=5, sticky="w", padx=4)
        ttk.Label(settings, text="Loss").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(settings, textvariable=self.loss_name_var, width=18).grid(row=1, column=1, sticky="w", padx=4, pady=(6, 0))
        ttk.Label(settings, text="Scheduler").grid(row=1, column=2, sticky="w", pady=(6, 0))
        ttk.Entry(settings, textvariable=self.scheduler_name_var, width=18).grid(row=1, column=3, sticky="w", padx=4, pady=(6, 0))

        training = ttk.LabelFrame(root, text="Train-time knobs", padding=8)
        training.pack(fill="x", pady=6)
        ttk.Label(training, text="Batch size").grid(row=0, column=0, sticky="w")
        ttk.Entry(training, textvariable=self.batch_size_var, width=12).grid(row=0, column=1, sticky="w", padx=4)
        ttk.Label(training, text="Epochs").grid(row=0, column=2, sticky="w")
        ttk.Entry(training, textvariable=self.epochs_var, width=12).grid(row=0, column=3, sticky="w", padx=4)
        ttk.Checkbutton(training, text="Early stopping", variable=self.early_stop_var).grid(row=0, column=4, sticky="w", padx=6)
        ttk.Label(training, text="Patience").grid(row=0, column=5, sticky="w")
        ttk.Entry(training, textvariable=self.early_patience_var, width=12).grid(row=0, column=6, sticky="w", padx=4)

        actions = ttk.Frame(root)
        actions.pack(fill="x", pady=(8, 0))
        ttk.Button(actions, text="Export YAML", command=self._export_yaml).pack(side="left")
        ttk.Button(actions, text="Export Python", command=self._export_python).pack(side="left", padx=6)
        ttk.Button(actions, text="Save architecture", command=self._save_architecture).pack(side="right")

    def _refresh_layers_text(self) -> None:
        lines = []
        for idx, layer in enumerate(self._spec["layers"], start=1):
            lines.append(f"{idx}. {layer}")
        self.layers_var.set("\n".join(lines))

    def _add_layer(self) -> None:
        layer_type = self.layer_type_var.get()
        if layer_type == "Dense":
            self._spec["layers"].append({"type": "Dense", "units": 64, "activation": "relu"})
        elif layer_type == "Dropout":
            self._spec["layers"].append({"type": "Dropout", "rate": 0.2})
        elif layer_type == "Norm":
            self._spec["layers"].append({"type": "Norm", "norm": "batch"})
        else:
            self._spec["layers"].append({"type": "Activation", "name": "relu"})
        self._refresh_layers_text()

    def _remove_last_layer(self) -> None:
        if self._spec["layers"]:
            self._spec["layers"].pop()
            if not self._spec["layers"]:
                self._spec["layers"] = default_architecture_spec()["layers"]
        self._refresh_layers_text()

    def _effective_spec(self) -> dict[str, Any]:
        payload = dict(self._spec)
        payload["optimizer"] = {
            "name": self.optimizer_name_var.get().strip() or "adam",
            "learning_rate": _as_float(self.optimizer_lr_var.get(), 0.001),
            "weight_decay": _as_float(self.optimizer_wd_var.get(), 0.0),
        }
        payload["loss"] = {"name": self.loss_name_var.get().strip() or "binary_cross_entropy"}
        payload["scheduler"] = {"name": self.scheduler_name_var.get().strip() or "none"}
        payload["training"] = {
            "batch_size": max(1, _as_int(self.batch_size_var.get(), 32)),
            "epochs": max(1, _as_int(self.epochs_var.get(), 50)),
            "early_stopping": {
                "enabled": bool(self.early_stop_var.get()),
                "patience": max(1, _as_int(self.early_patience_var.get(), 6)),
                "min_delta": 0.0001,
            },
        }
        return normalize_architecture_spec(payload)

    def _load_selected_preset(self) -> None:
        label = self.preset_var.get()
        if label.startswith("<built-in> "):
            key = label.replace("<built-in> ", "", 1)
            preset = self._built_in_presets.get(key)
        elif label.startswith("<custom> "):
            key = label.replace("<custom> ", "", 1)
            preset = self._custom_presets.get(key)
        else:
            preset = None
        if preset is None:
            return
        self._spec = normalize_architecture_spec(preset)
        self._refresh_layers_text()

    def _save_as_preset(self) -> None:
        name = simpledialog.askstring("Preset name", "Enter new preset name:", parent=self)
        if not name:
            return
        self._custom_presets[name] = self._effective_spec()
        messagebox.showinfo("Preset saved", f"Saved custom preset '{name}'.")

    def _export_yaml(self) -> None:
        export_path = self._write_export("yml", architecture_to_yaml(self._effective_spec()))
        messagebox.showinfo("YAML exported", f"Architecture exported to:\n{export_path}")

    def _export_python(self) -> None:
        export_path = self._write_export("py", architecture_to_python(self._effective_spec()))
        messagebox.showinfo("Python exported", f"Architecture exported to:\n{export_path}")

    def _write_export(self, suffix: str, content: str) -> str:
        NN_EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        path = NN_EXPORT_DIR / f"architecture_{stamp}.{suffix}"
        path.write_text(content, encoding="utf-8")
        return str(path)

    def _save_architecture(self) -> None:
        self._on_save(self._effective_spec())
        self.destroy()

    @property
    def custom_presets(self) -> dict[str, dict[str, Any]]:
        return {key: dict(value) for key, value in self._custom_presets.items()}

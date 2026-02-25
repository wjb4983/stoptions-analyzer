from __future__ import annotations

import copy
import json
from datetime import datetime, timezone
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
from typing import Any, Callable

NN_PRESETS_PATH = Path("config/nn_presets.json")
NN_MODEL_PROFILES_PATH = Path("config/nn_model_profiles.json")
NN_EXPORT_DIR = Path("data/nn_architectures")
_LAYER_TYPES = ["Dense", "Dropout", "Norm", "Activation"]


class LayerEditorRow(ttk.Frame):
    def __init__(self, parent: ttk.Frame, *, on_change: Callable[[], None]) -> None:
        super().__init__(parent)
        self._on_change = on_change
        self._widgets: list[tk.Widget] = []

    def rebuild(self, layer: dict[str, Any]) -> None:
        for widget in self._widgets:
            widget.destroy()
        self._widgets = []

        layer_type = str(layer.get("type", "Dense"))
        row = 0
        ttk.Label(self, text=f"Type: {layer_type}").grid(row=row, column=0, sticky="w", pady=(0, 4))
        row += 1

        if layer_type == "Dense":
            self._add_labeled_entry(row, "Units", "units", layer)
            row += 1
            self._add_labeled_entry(row, "Activation", "activation", layer)
            row += 1
            bias_var = tk.BooleanVar(value=bool(layer.get("use_bias", True)))

            def _set_bias(*_: Any) -> None:
                layer["use_bias"] = bool(bias_var.get())
                self._on_change()

            bias_var.trace_add("write", _set_bias)
            w = ttk.Checkbutton(self, text="Use bias", variable=bias_var)
            w.grid(row=row, column=0, sticky="w")
            self._widgets.append(w)
        elif layer_type == "Dropout":
            self._add_labeled_entry(row, "Rate", "rate", layer)
        elif layer_type == "Norm":
            self._add_labeled_entry(row, "Norm mode", "norm", layer)
        elif layer_type == "Activation":
            self._add_labeled_entry(row, "Activation name", "name", layer)

    def _add_labeled_entry(self, row: int, label: str, key: str, layer: dict[str, Any]) -> None:
        ttk.Label(self, text=label).grid(row=row, column=0, sticky="w")
        var = tk.StringVar(value=str(layer.get(key, "")))

        def _on_var_change(*_: Any) -> None:
            value = var.get().strip()
            if key in {"units"}:
                layer[key] = _as_int(value, 0)
            elif key in {"rate"}:
                layer[key] = _as_float(value, -1.0)
            else:
                layer[key] = value
            self._on_change()

        var.trace_add("write", _on_var_change)
        entry = ttk.Entry(self, textvariable=var, width=20)
        entry.grid(row=row, column=1, sticky="w", padx=(6, 0), pady=2)
        self._widgets.extend([entry])


def default_architecture_spec() -> dict[str, Any]:
    return {
        "schema_version": 1,
        "layers": [
            {"type": "Dense", "units": 64, "activation": "relu", "use_bias": True},
            {"type": "Dropout", "rate": 0.2},
            {"type": "Dense", "units": 32, "activation": "relu", "use_bias": True},
            {"type": "Dense", "units": 1, "activation": "sigmoid", "use_bias": True},
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


def _normalize_metadata(payload: dict[str, Any] | None) -> dict[str, Any]:
    item = dict(payload) if isinstance(payload, dict) else {}
    compatibility = item.get("compatibility")
    if not isinstance(compatibility, dict):
        compatibility = {}
    leg_tags = compatibility.get("leg_tags")
    model_families = compatibility.get("model_families")
    item["compatibility"] = {
        "leg_tags": [str(v).strip() for v in (leg_tags or []) if str(v).strip()],
        "model_families": [str(v).strip() for v in (model_families or []) if str(v).strip()],
    }
    if not str(item.get("updated_at", "")).strip():
        item["updated_at"] = datetime.now(timezone.utc).isoformat()
    return item


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


def _extract_custom_preset_entries(payload: dict[str, Any] | None) -> dict[str, dict[str, Any]]:
    if not isinstance(payload, dict):
        return {}
    raw = payload.get("custom_presets", {})
    result: dict[str, dict[str, Any]] = {}
    if isinstance(raw, dict):
        for name, item in raw.items():
            if not isinstance(name, str) or not isinstance(item, dict):
                continue
            spec = item.get("spec") if isinstance(item.get("spec"), dict) else item
            result[name] = {
                "spec": normalize_architecture_spec(spec),
                "metadata": _normalize_metadata(item.get("metadata") if isinstance(item.get("metadata"), dict) else item),
            }
    return result


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


def save_nn_presets(
    presets: dict[str, dict[str, Any]],
    *,
    custom_presets: dict[str, dict[str, Any]] | None = None,
) -> None:
    payload: dict[str, Any] = {
        "default_preset": next(iter(presets), ""),
        "presets": {name: normalize_architecture_spec(spec) for name, spec in presets.items()},
    }
    if custom_presets:
        payload["custom_presets"] = {
            name: {
                "spec": normalize_architecture_spec(item.get("spec") if isinstance(item, dict) else None),
                "metadata": _normalize_metadata(item.get("metadata") if isinstance(item, dict) else None),
            }
            for name, item in custom_presets.items()
            if isinstance(name, str)
        }
    NN_PRESETS_PATH.parent.mkdir(parents=True, exist_ok=True)
    NN_PRESETS_PATH.write_text(json.dumps(payload, indent=2, sort_keys=True), encoding="utf-8")


def architecture_to_yaml(payload: dict[str, Any]) -> str:
    normalized = normalize_architecture_spec(payload)
    try:
        import yaml
    except ModuleNotFoundError:
        return json.dumps(normalized, indent=2)
    return yaml.safe_dump(normalized, sort_keys=False)


def architecture_to_python(payload: dict[str, Any], *, function_name: str = "build_model_spec") -> str:
    spec_literal = repr(normalize_architecture_spec(payload))
    return (
        '"""Generated neural network architecture spec for regime training."""\n\n'
        "from __future__ import annotations\n\n"
        "from typing import Any\n\n"
        f"def {function_name}() -> dict[str, Any]:\n"
        f"    return {spec_literal}\n"
    )


def _validate_architecture_spec_for_ui(payload: dict[str, Any] | None, *, field_path: str) -> list[str]:
    if not isinstance(payload, dict):
        return [f"{field_path} is required for ANN/neural model legs"]
    errors: list[str] = []
    schema_version = payload.get("schema_version")
    if schema_version is not None and _as_int(schema_version, 0) < 1:
        errors.append(f"{field_path}.schema_version must be >= 1")
    layers = payload.get("layers")
    if not isinstance(layers, list) or not layers:
        errors.append(f"{field_path}.layers must be a non-empty list")
    else:
        for idx, layer in enumerate(layers):
            if not isinstance(layer, dict):
                errors.append(f"{field_path}.layers[{idx}] must be an object")
                continue
            layer_type = str(layer.get("type", "")).strip()
            if layer_type not in set(_LAYER_TYPES):
                errors.append(f"{field_path}.layers[{idx}].type must be one of {_LAYER_TYPES}")
                continue
            if layer_type == "Dense":
                if _as_int(layer.get("units", 0), 0) <= 0:
                    errors.append(f"{field_path}.layers[{idx}].units must be > 0")
                if not str(layer.get("activation", "")).strip():
                    errors.append(f"{field_path}.layers[{idx}].activation is required")
            if layer_type == "Dropout":
                rate = _as_float(layer.get("rate", -1.0), -1.0)
                if rate < 0.0 or rate >= 1.0:
                    errors.append(f"{field_path}.layers[{idx}].rate must be within [0, 1)")
            if layer_type == "Norm" and not str(layer.get("norm", "")).strip():
                errors.append(f"{field_path}.layers[{idx}].norm is required")
            if layer_type == "Activation" and not str(layer.get("name", "")).strip():
                errors.append(f"{field_path}.layers[{idx}].name is required")

    optimizer = payload.get("optimizer")
    if not isinstance(optimizer, dict):
        errors.append(f"{field_path}.optimizer must be an object")
    else:
        if not str(optimizer.get("name", "")).strip():
            errors.append(f"{field_path}.optimizer.name is required")
        if _as_float(optimizer.get("learning_rate", 0.0), 0.0) <= 0.0:
            errors.append(f"{field_path}.optimizer.learning_rate must be > 0")

    loss = payload.get("loss")
    if not isinstance(loss, dict) or not str(loss.get("name", "")).strip():
        errors.append(f"{field_path}.loss.name is required")

    scheduler = payload.get("scheduler")
    if not isinstance(scheduler, dict) or not str(scheduler.get("name", "")).strip():
        errors.append(f"{field_path}.scheduler.name is required")

    training = payload.get("training")
    if not isinstance(training, dict):
        errors.append(f"{field_path}.training must be an object")
    else:
        if _as_int(training.get("batch_size", 0), 0) <= 0:
            errors.append(f"{field_path}.training.batch_size must be > 0")
        if _as_int(training.get("epochs", 0), 0) <= 0:
            errors.append(f"{field_path}.training.epochs must be > 0")
        early = training.get("early_stopping")
        if not isinstance(early, dict):
            errors.append(f"{field_path}.training.early_stopping must be an object")
        else:
            if not isinstance(early.get("enabled"), bool):
                errors.append(f"{field_path}.training.early_stopping.enabled must be boolean")
            if _as_int(early.get("patience", 0), 0) <= 0:
                errors.append(f"{field_path}.training.early_stopping.patience must be > 0")
            if _as_float(early.get("min_delta", -1.0), -1.0) < 0:
                errors.append(f"{field_path}.training.early_stopping.min_delta must be >= 0")
    return errors


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
        self.geometry("1200x780")
        self.transient(parent)

        self._on_save = on_save
        self._built_in_presets = load_nn_presets()
        self._persisted_custom_entries = _extract_custom_preset_entries(
            json.loads(NN_PRESETS_PATH.read_text(encoding="utf-8")) if NN_PRESETS_PATH.exists() else {}
        )
        incoming_custom = {k: {"spec": normalize_architecture_spec(v), "metadata": _normalize_metadata(None)} for k, v in (custom_presets or {}).items()}
        self._custom_preset_entries = {**self._persisted_custom_entries, **incoming_custom}
        self._spec = normalize_architecture_spec(initial_spec)
        self._selected_layer_index: int | None = None

        self.layer_type_var = tk.StringVar(value="Dense")
        self.preset_var = tk.StringVar(value="")
        self.preview_format_var = tk.StringVar(value="json")
        self.optimizer_name_var = tk.StringVar(value=str(self._spec["optimizer"].get("name", "adam")))
        self.optimizer_lr_var = tk.StringVar(value=str(self._spec["optimizer"].get("learning_rate", 0.001)))
        self.optimizer_wd_var = tk.StringVar(value=str(self._spec["optimizer"].get("weight_decay", 0.0)))
        self.optimizer_beta1_var = tk.StringVar(value=str(self._spec["optimizer"].get("beta1", "")))
        self.optimizer_beta2_var = tk.StringVar(value=str(self._spec["optimizer"].get("beta2", "")))
        self.loss_name_var = tk.StringVar(value=str(self._spec["loss"].get("name", "binary_cross_entropy")))
        self.scheduler_name_var = tk.StringVar(value=str(self._spec["scheduler"].get("name", "none")))
        self.batch_size_var = tk.StringVar(value=str(self._spec["training"].get("batch_size", 32)))
        self.epochs_var = tk.StringVar(value=str(self._spec["training"].get("epochs", 50)))
        self.early_stop_var = tk.BooleanVar(value=bool(self._spec["training"]["early_stopping"].get("enabled", True)))
        self.early_patience_var = tk.StringVar(value=str(self._spec["training"]["early_stopping"].get("patience", 6)))
        self.early_min_delta_var = tk.StringVar(value=str(self._spec["training"]["early_stopping"].get("min_delta", 0.0001)))

        self.compat_leg_tags_var = tk.StringVar(value="")
        self.compat_model_families_var = tk.StringVar(value="")
        self.layer_editor: LayerEditorRow | None = None

        self._build_layout()
        self._refresh_preset_options()
        self._refresh_layers_table()
        self._refresh_validation_and_preview()

    def _build_layout(self) -> None:
        root = ttk.Frame(self, padding=10)
        root.pack(fill="both", expand=True)

        preset_row = ttk.Frame(root)
        preset_row.pack(fill="x", pady=(0, 6))
        ttk.Label(preset_row, text="Preset").pack(side="left")
        self.preset_combo = ttk.Combobox(preset_row, textvariable=self.preset_var, state="readonly", width=46)
        self.preset_combo.pack(side="left", padx=6)
        ttk.Button(preset_row, text="Load", command=self._load_selected_preset).pack(side="left", padx=2)
        ttk.Button(preset_row, text="Save/Update", command=self._save_as_preset).pack(side="left", padx=2)
        ttk.Button(preset_row, text="Delete", command=self._delete_selected_preset).pack(side="left", padx=2)
        ttk.Button(preset_row, text="Save as model profile", command=self._save_as_model_profile).pack(side="left", padx=2)

        preset_meta = ttk.Frame(root)
        preset_meta.pack(fill="x", pady=(0, 6))
        ttk.Label(preset_meta, text="Compat leg tags (csv)").grid(row=0, column=0, sticky="w")
        ttk.Entry(preset_meta, textvariable=self.compat_leg_tags_var, width=30).grid(row=0, column=1, sticky="w", padx=4)
        ttk.Label(preset_meta, text="Compat model families (csv)").grid(row=0, column=2, sticky="w")
        ttk.Entry(preset_meta, textvariable=self.compat_model_families_var, width=36).grid(row=0, column=3, sticky="w", padx=4)

        mid = ttk.Panedwindow(root, orient="horizontal")
        mid.pack(fill="both", expand=True, pady=6)

        left = ttk.Frame(mid)
        mid.add(left, weight=3)
        right = ttk.Frame(mid)
        mid.add(right, weight=2)

        layers = ttk.LabelFrame(left, text="Editable layer table", padding=8)
        layers.pack(fill="both", expand=True)
        toolbar = ttk.Frame(layers)
        toolbar.pack(fill="x", pady=(0, 6))
        ttk.Label(toolbar, text="Type").pack(side="left")
        ttk.Combobox(toolbar, textvariable=self.layer_type_var, state="readonly", values=_LAYER_TYPES, width=12).pack(side="left", padx=6)
        ttk.Button(toolbar, text="Add", command=self._add_layer).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Up", command=lambda: self._move_layer(-1)).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Down", command=lambda: self._move_layer(1)).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Duplicate", command=self._duplicate_layer).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Delete", command=self._delete_layer).pack(side="left", padx=2)

        self.layer_tree = ttk.Treeview(layers, columns=("idx", "type", "config"), show="headings", height=9)
        self.layer_tree.heading("idx", text="#")
        self.layer_tree.heading("type", text="Type")
        self.layer_tree.heading("config", text="Config")
        self.layer_tree.column("idx", width=40, anchor="center")
        self.layer_tree.column("type", width=100, anchor="w")
        self.layer_tree.column("config", width=460, anchor="w")
        self.layer_tree.pack(fill="x")
        self.layer_tree.bind("<<TreeviewSelect>>", self._on_layer_selected)

        layer_detail = ttk.LabelFrame(layers, text="Per-layer form", padding=8)
        layer_detail.pack(fill="x", pady=(8, 0))
        self.layer_editor = LayerEditorRow(layer_detail, on_change=self._on_layer_detail_changed)
        self.layer_editor.pack(fill="x")

        settings = ttk.LabelFrame(left, text="Optimizer / loss / scheduler", padding=8)
        settings.pack(fill="x", pady=6)
        ttk.Label(settings, text="Optimizer").grid(row=0, column=0, sticky="w")
        ttk.Entry(settings, textvariable=self.optimizer_name_var, width=16).grid(row=0, column=1, sticky="w", padx=4)
        ttk.Label(settings, text="Learning rate").grid(row=0, column=2, sticky="w")
        ttk.Entry(settings, textvariable=self.optimizer_lr_var, width=10).grid(row=0, column=3, sticky="w", padx=4)
        ttk.Label(settings, text="Weight decay").grid(row=0, column=4, sticky="w")
        ttk.Entry(settings, textvariable=self.optimizer_wd_var, width=10).grid(row=0, column=5, sticky="w", padx=4)
        ttk.Label(settings, text="Beta1").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(settings, textvariable=self.optimizer_beta1_var, width=10).grid(row=1, column=1, sticky="w", padx=4, pady=(6, 0))
        ttk.Label(settings, text="Beta2").grid(row=1, column=2, sticky="w", pady=(6, 0))
        ttk.Entry(settings, textvariable=self.optimizer_beta2_var, width=10).grid(row=1, column=3, sticky="w", padx=4, pady=(6, 0))
        ttk.Label(settings, text="Loss").grid(row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(settings, textvariable=self.loss_name_var, width=16).grid(row=2, column=1, sticky="w", padx=4, pady=(6, 0))
        ttk.Label(settings, text="Scheduler").grid(row=2, column=2, sticky="w", pady=(6, 0))
        ttk.Entry(settings, textvariable=self.scheduler_name_var, width=16).grid(row=2, column=3, sticky="w", padx=4, pady=(6, 0))

        training = ttk.LabelFrame(left, text="Train-time knobs", padding=8)
        training.pack(fill="x", pady=6)
        ttk.Label(training, text="Batch size").grid(row=0, column=0, sticky="w")
        ttk.Entry(training, textvariable=self.batch_size_var, width=12).grid(row=0, column=1, sticky="w", padx=4)
        ttk.Label(training, text="Epochs").grid(row=0, column=2, sticky="w")
        ttk.Entry(training, textvariable=self.epochs_var, width=12).grid(row=0, column=3, sticky="w", padx=4)
        ttk.Checkbutton(training, text="Early stopping", variable=self.early_stop_var).grid(row=0, column=4, sticky="w", padx=6)
        ttk.Label(training, text="Patience").grid(row=0, column=5, sticky="w")
        ttk.Entry(training, textvariable=self.early_patience_var, width=8).grid(row=0, column=6, sticky="w", padx=4)
        ttk.Label(training, text="Min delta").grid(row=0, column=7, sticky="w")
        ttk.Entry(training, textvariable=self.early_min_delta_var, width=10).grid(row=0, column=8, sticky="w", padx=4)

        validation = ttk.LabelFrame(right, text="Real-time validation", padding=8)
        validation.pack(fill="both", expand=True)
        self.validation_text = tk.Text(validation, height=14, wrap="word")
        self.validation_text.pack(fill="both", expand=True)

        preview = ttk.LabelFrame(right, text="Spec preview", padding=8)
        preview.pack(fill="both", expand=True, pady=(6, 0))
        fmt = ttk.Frame(preview)
        fmt.pack(fill="x")
        ttk.Radiobutton(fmt, text="JSON", variable=self.preview_format_var, value="json", command=self._refresh_preview).pack(side="left")
        ttk.Radiobutton(fmt, text="YAML", variable=self.preview_format_var, value="yaml", command=self._refresh_preview).pack(side="left", padx=(8, 0))
        self.preview_text = tk.Text(preview, height=16, wrap="none")
        self.preview_text.pack(fill="both", expand=True, pady=(6, 0))

        actions = ttk.Frame(root)
        actions.pack(fill="x", pady=(8, 0))
        ttk.Button(actions, text="Export YAML", command=self._export_yaml).pack(side="left")
        ttk.Button(actions, text="Export Python", command=self._export_python).pack(side="left", padx=6)
        ttk.Button(actions, text="Save architecture", command=self._save_architecture).pack(side="right")

        for var in (
            self.optimizer_name_var,
            self.optimizer_lr_var,
            self.optimizer_wd_var,
            self.optimizer_beta1_var,
            self.optimizer_beta2_var,
            self.loss_name_var,
            self.scheduler_name_var,
            self.batch_size_var,
            self.epochs_var,
            self.early_patience_var,
            self.early_min_delta_var,
        ):
            var.trace_add("write", lambda *_: self._refresh_validation_and_preview())
        self.early_stop_var.trace_add("write", lambda *_: self._refresh_validation_and_preview())

    def _refresh_preset_options(self) -> None:
        values = [f"<built-in> {name}" for name in sorted(self._built_in_presets)] + [
            f"<custom> {name}" for name in sorted(self._custom_preset_entries)
        ]
        self.preset_combo.configure(values=values)

    def _refresh_layers_table(self) -> None:
        for item in self.layer_tree.get_children():
            self.layer_tree.delete(item)
        for idx, layer in enumerate(self._spec["layers"]):
            config = ", ".join(f"{k}={v}" for k, v in layer.items() if k != "type")
            self.layer_tree.insert("", "end", iid=str(idx), values=(idx + 1, layer.get("type", ""), config))
        if self._selected_layer_index is not None and str(self._selected_layer_index) in self.layer_tree.get_children():
            self.layer_tree.selection_set(str(self._selected_layer_index))

    def _on_layer_selected(self, *_: Any) -> None:
        selection = self.layer_tree.selection()
        self._selected_layer_index = int(selection[0]) if selection else None
        if self._selected_layer_index is None:
            return
        layer = self._spec["layers"][self._selected_layer_index]
        if self.layer_editor is not None:
            self.layer_editor.rebuild(layer)

    def _on_layer_detail_changed(self) -> None:
        self._refresh_layers_table()
        self._refresh_validation_and_preview()

    def _add_layer(self) -> None:
        layer_type = self.layer_type_var.get()
        if layer_type == "Dense":
            layer = {"type": "Dense", "units": 64, "activation": "relu", "use_bias": True}
        elif layer_type == "Dropout":
            layer = {"type": "Dropout", "rate": 0.2}
        elif layer_type == "Norm":
            layer = {"type": "Norm", "norm": "batch"}
        else:
            layer = {"type": "Activation", "name": "relu"}
        self._spec["layers"].append(layer)
        self._selected_layer_index = len(self._spec["layers"]) - 1
        self._refresh_layers_table()
        self.layer_tree.selection_set(str(self._selected_layer_index))
        self._on_layer_selected()
        self._refresh_validation_and_preview()

    def _move_layer(self, delta: int) -> None:
        if self._selected_layer_index is None:
            return
        target = self._selected_layer_index + delta
        if target < 0 or target >= len(self._spec["layers"]):
            return
        layers = self._spec["layers"]
        layers[self._selected_layer_index], layers[target] = layers[target], layers[self._selected_layer_index]
        self._selected_layer_index = target
        self._refresh_layers_table()
        self.layer_tree.selection_set(str(self._selected_layer_index))
        self._refresh_validation_and_preview()

    def _duplicate_layer(self) -> None:
        if self._selected_layer_index is None:
            return
        dup = copy.deepcopy(self._spec["layers"][self._selected_layer_index])
        self._spec["layers"].insert(self._selected_layer_index + 1, dup)
        self._selected_layer_index += 1
        self._refresh_layers_table()
        self.layer_tree.selection_set(str(self._selected_layer_index))
        self._refresh_validation_and_preview()

    def _delete_layer(self) -> None:
        if self._selected_layer_index is None:
            return
        self._spec["layers"].pop(self._selected_layer_index)
        if not self._spec["layers"]:
            self._spec["layers"] = default_architecture_spec()["layers"]
        self._selected_layer_index = min(self._selected_layer_index, len(self._spec["layers"]) - 1)
        self._refresh_layers_table()
        self._refresh_validation_and_preview()

    def _effective_spec(self) -> dict[str, Any]:
        payload = dict(self._spec)
        optimizer: dict[str, Any] = {
            "name": self.optimizer_name_var.get().strip() or "adam",
            "learning_rate": _as_float(self.optimizer_lr_var.get(), 0.001),
            "weight_decay": _as_float(self.optimizer_wd_var.get(), 0.0),
        }
        if self.optimizer_beta1_var.get().strip():
            optimizer["beta1"] = _as_float(self.optimizer_beta1_var.get(), 0.9)
        if self.optimizer_beta2_var.get().strip():
            optimizer["beta2"] = _as_float(self.optimizer_beta2_var.get(), 0.999)
        payload["optimizer"] = optimizer
        payload["loss"] = {"name": self.loss_name_var.get().strip() or "binary_cross_entropy"}
        payload["scheduler"] = {"name": self.scheduler_name_var.get().strip() or "none"}
        payload["training"] = {
            "batch_size": max(1, _as_int(self.batch_size_var.get(), 32)),
            "epochs": max(1, _as_int(self.epochs_var.get(), 50)),
            "early_stopping": {
                "enabled": bool(self.early_stop_var.get()),
                "patience": max(1, _as_int(self.early_patience_var.get(), 6)),
                "min_delta": max(0.0, _as_float(self.early_min_delta_var.get(), 0.0001)),
            },
        }
        return normalize_architecture_spec(payload)

    def _refresh_validation_and_preview(self) -> None:
        spec = self._effective_spec()
        errors = _validate_architecture_spec_for_ui(spec, field_path="architecture_spec")
        self.validation_text.delete("1.0", "end")
        if errors:
            self.validation_text.insert("1.0", "\n".join(f"• {line}" for line in errors))
        else:
            self.validation_text.insert("1.0", "No validation errors.")
        self._refresh_preview()

    def _refresh_preview(self) -> None:
        spec = self._effective_spec()
        content = json.dumps(spec, indent=2) if self.preview_format_var.get() == "json" else architecture_to_yaml(spec)
        self.preview_text.delete("1.0", "end")
        self.preview_text.insert("1.0", content)

    def _load_selected_preset(self) -> None:
        label = self.preset_var.get()
        if label.startswith("<built-in> "):
            key = label.replace("<built-in> ", "", 1)
            preset = self._built_in_presets.get(key)
            metadata = None
        elif label.startswith("<custom> "):
            key = label.replace("<custom> ", "", 1)
            item = self._custom_preset_entries.get(key)
            preset = item.get("spec") if isinstance(item, dict) else None
            metadata = item.get("metadata") if isinstance(item, dict) else None
        else:
            preset = None
            metadata = None
        if preset is None:
            return
        self._spec = normalize_architecture_spec(preset)
        self.compat_leg_tags_var.set(", ".join((metadata or {}).get("compatibility", {}).get("leg_tags", [])))
        self.compat_model_families_var.set(", ".join((metadata or {}).get("compatibility", {}).get("model_families", [])))
        self._selected_layer_index = 0
        self._refresh_layers_table()
        self.layer_tree.selection_set("0")
        self._on_layer_selected()
        self._refresh_validation_and_preview()

    def _save_as_preset(self) -> None:
        name = simpledialog.askstring("Preset name", "Enter preset name:", parent=self)
        if not name:
            return
        metadata = _normalize_metadata(
            {
                "compatibility": {
                    "leg_tags": [v.strip() for v in self.compat_leg_tags_var.get().split(",") if v.strip()],
                    "model_families": [v.strip() for v in self.compat_model_families_var.get().split(",") if v.strip()],
                }
            }
        )
        self._custom_preset_entries[name] = {"spec": self._effective_spec(), "metadata": metadata}
        save_nn_presets(self._built_in_presets, custom_presets=self._custom_preset_entries)
        self._refresh_preset_options()
        messagebox.showinfo("Preset saved", f"Saved custom preset '{name}'.")

    def _delete_selected_preset(self) -> None:
        label = self.preset_var.get()
        if not label.startswith("<custom> "):
            messagebox.showwarning("Delete preset", "Select a custom preset to delete.")
            return
        key = label.replace("<custom> ", "", 1)
        if key in self._custom_preset_entries:
            del self._custom_preset_entries[key]
            save_nn_presets(self._built_in_presets, custom_presets=self._custom_preset_entries)
            self._refresh_preset_options()
            self.preset_var.set("")

    def _save_as_model_profile(self) -> None:
        name = simpledialog.askstring("Model profile", "Model profile name:", parent=self)
        if not name:
            return
        payload: dict[str, Any] = {}
        if NN_MODEL_PROFILES_PATH.exists():
            try:
                payload = json.loads(NN_MODEL_PROFILES_PATH.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                payload = {}
        profiles = payload.get("profiles", {}) if isinstance(payload, dict) else {}
        if not isinstance(profiles, dict):
            profiles = {}
        profiles[name] = {
            "architecture": self._effective_spec(),
            "compatibility": {
                "leg_tags": [v.strip() for v in self.compat_leg_tags_var.get().split(",") if v.strip()],
                "model_families": [v.strip() for v in self.compat_model_families_var.get().split(",") if v.strip()],
            },
            "updated_at": datetime.now(timezone.utc).isoformat(),
        }
        NN_MODEL_PROFILES_PATH.parent.mkdir(parents=True, exist_ok=True)
        payload["profiles"] = profiles
        NN_MODEL_PROFILES_PATH.write_text(json.dumps(payload, indent=2, sort_keys=True), encoding="utf-8")
        messagebox.showinfo("Model profile", f"Saved model profile '{name}'.")

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
        return {key: dict(value.get("spec", {})) for key, value in self._custom_preset_entries.items()}

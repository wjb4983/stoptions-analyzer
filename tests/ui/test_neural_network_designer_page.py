from __future__ import annotations

import builtins
import yaml

from ui import neural_network_designer_page as nn_page


def test_preset_round_trip(tmp_path, monkeypatch):
    path = tmp_path / "nn_presets.json"
    monkeypatch.setattr(nn_page, "NN_PRESETS_PATH", path)

    presets = {
        "small_mlp": nn_page.default_architecture_spec(),
        "custom": {
            "layers": [{"type": "Dense", "units": 32, "activation": "relu"}],
            "optimizer": {"name": "adam", "learning_rate": 0.001},
            "loss": {"name": "binary_cross_entropy"},
            "scheduler": {"name": "none"},
            "training": {"batch_size": 16, "epochs": 10, "early_stopping": {"enabled": True, "patience": 3}},
        },
    }
    nn_page.save_nn_presets(presets)
    loaded = nn_page.load_nn_presets()

    assert set(loaded) == {"small_mlp", "custom"}
    assert loaded["custom"]["training"]["epochs"] == 10


def test_yaml_and_python_export_correctness():
    spec = nn_page.default_architecture_spec()

    yml = nn_page.architecture_to_yaml(spec)
    parsed = yaml.safe_load(yml)
    assert parsed["layers"][0]["type"] == "Dense"
    assert parsed["optimizer"]["name"] == "adam"

    py_source = nn_page.architecture_to_python(spec, function_name="factory")
    namespace: dict[str, object] = {}
    exec(py_source, namespace)  # noqa: S102
    built = namespace["factory"]()
    assert built["training"]["batch_size"] == 32
    assert built["loss"]["name"] == "binary_cross_entropy"


def test_normalize_architecture_spec_fills_defaults():
    normalized = nn_page.normalize_architecture_spec({"layers": [{"type": "Dense", "units": 8, "activation": "relu"}]})
    assert normalized["optimizer"]["name"] == "adam"
    assert normalized["training"]["early_stopping"]["enabled"] is True


def test_yaml_export_falls_back_to_json_when_pyyaml_missing(monkeypatch):
    original_import = builtins.__import__

    def raising_import(name, *args, **kwargs):
        if name == "yaml":
            raise ModuleNotFoundError("No module named 'yaml'")
        return original_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", raising_import)
    content = nn_page.architecture_to_yaml(nn_page.default_architecture_spec())
    parsed = yaml.safe_load(content)

    assert parsed["layers"][0]["type"] == "Dense"
    assert parsed["optimizer"]["name"] == "adam"


def test_normalize_architecture_spec_handles_new_layer_types():
    normalized = nn_page.normalize_architecture_spec(
        {
            "layers": [
                {"type": "Conv1D", "filters": "32", "kernel_size": "5", "stride": 1, "padding": "same", "activation": "relu"},
                {"type": "Attention", "heads": "4", "key_dim": "16", "dropout": "0.2", "causal": True},
                {"type": "PolicyHead", "action_dim": "3", "distribution": "categorical"},
                {"type": "ValueHead", "value_dim": "1", "activation": "linear"},
            ]
        }
    )

    assert normalized["layers"][0]["filters"] == 32
    assert normalized["layers"][1]["dropout"] == 0.2
    assert normalized["layers"][2]["action_dim"] == 3
    assert normalized["layers"][3]["value_dim"] == 1


def test_validate_architecture_spec_flags_malformed_new_layers():
    bad_spec = {
        "schema_version": 1,
        "layers": [
            {"type": "Conv1D", "filters": 0, "kernel_size": 0, "stride": 0, "padding": "invalid", "activation": ""},
            {"type": "Attention", "heads": 0, "key_dim": 0, "dropout": 1.2, "causal": "yes"},
            {"type": "PolicyHead", "action_dim": 0, "distribution": ""},
            {"type": "ValueHead", "value_dim": 0, "activation": ""},
        ],
        "optimizer": {"name": "adam", "learning_rate": 0.001},
        "loss": {"name": "mse"},
        "scheduler": {"name": "none"},
        "training": {"batch_size": 32, "epochs": 10, "early_stopping": {"enabled": True, "patience": 3, "min_delta": 0.0}},
    }

    errors = nn_page._validate_architecture_spec_for_ui(bad_spec, field_path="architecture_spec")

    assert "architecture_spec.layers[0].filters must be > 0" in errors
    assert "architecture_spec.layers[0].padding must be one of ['same', 'valid', 'causal']" in errors
    assert "architecture_spec.layers[1].causal must be boolean" in errors
    assert "architecture_spec.layers[2].distribution is required" in errors
    assert "architecture_spec.layers[3].value_dim must be > 0" in errors

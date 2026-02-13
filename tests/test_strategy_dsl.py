from __future__ import annotations

import json

import pytest

from src.backtesting.strategies.dsl import (
    BUILTIN_STRATEGY_TEMPLATES,
    StrategyDSLValidationError,
    compile_strategy,
    compile_template,
    list_template_names,
    parse_strategy_payload,
    parse_strategy_text,
)


def _payload() -> dict[str, object]:
    return {
        "name": "My Strategy",
        "family": "trend_following",
        "universe": {"symbols": ["SPY", "QQQ"]},
        "features": [{"name": "returns", "source": "price", "params": {"lookback_days": 90}}],
        "logic": {
            "entry": {"name": "ts_momentum", "params": {"lookback_days": 90, "skip_days": 5}},
            "exit": {"name": "momentum_flip", "params": {"lookback_days": 90, "skip_days": 5}},
        },
        "risk": {"rules": [{"name": "max_position", "params": {"weight": 0.3}}]},
        "execution": {"model": "bps", "params": {"bps": 5.0}},
    }


def test_parse_and_compile_json_strategy() -> None:
    definition = parse_strategy_payload(_payload())
    compiled = compile_strategy(definition)

    assert definition.universe.symbols == ("SPY", "QQQ")
    assert compiled.entry_config.name == "ts_momentum"
    assert compiled.exit_config.name == "momentum_flip"
    assert compiled.execution_model.name == "bps"


def test_parse_yaml_strategy_text_when_yaml_available() -> None:
    yaml = pytest.importorskip("yaml")
    text = yaml.safe_dump(_payload())
    definition = parse_strategy_text(text, format_hint="yaml")
    assert definition.name == "My Strategy"


def test_yaml_text_without_dependency_has_clear_error(monkeypatch: pytest.MonkeyPatch) -> None:
    import builtins

    real_import = builtins.__import__

    def fake_import(name, *args, **kwargs):
        if name == "yaml":
            raise ModuleNotFoundError("No module named 'yaml'")
        return real_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", fake_import)
    with pytest.raises(StrategyDSLValidationError, match="PyYAML is not installed"):
        parse_strategy_text("name: test", format_hint="yaml")


def test_schema_validation_has_clear_error_path() -> None:
    bad = _payload()
    bad["universe"] = {"symbols": ["SPY", ""]}

    with pytest.raises(StrategyDSLValidationError, match=r"universe\.symbols\[1\]"):
        parse_strategy_payload(bad)


def test_invalid_signal_in_config_is_reported() -> None:
    bad = _payload()
    logic = dict(bad["logic"])  # type: ignore[arg-type]
    logic["entry"] = {"name": "does_not_exist", "params": {}}
    bad["logic"] = logic

    definition = parse_strategy_payload(bad)
    with pytest.raises(ValueError, match="Unsupported entry signal"):
        compile_strategy(definition)


def test_builtin_templates_have_15_plus_and_compile() -> None:
    names = list_template_names()
    assert len(names) >= 15
    assert len(BUILTIN_STRATEGY_TEMPLATES) == len(names)

    for name in names:
        compiled = compile_template(name)
        assert compiled.definition.name


def test_unknown_top_level_key_rejected() -> None:
    bad = _payload()
    bad["alpha"] = {"foo": "bar"}

    with pytest.raises(StrategyDSLValidationError, match="unknown top-level keys"):
        parse_strategy_payload(bad)


def test_parse_strategy_text_json_autodetect() -> None:
    definition = parse_strategy_text(json.dumps(_payload()))
    assert definition.family == "trend_following"

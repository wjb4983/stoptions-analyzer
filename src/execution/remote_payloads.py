from __future__ import annotations

from dataclasses import asdict, fields, is_dataclass
from datetime import date, datetime
import importlib
from pathlib import Path
from typing import Any


_DATACLASS_TAG = "__dataclass__"
_CALLABLE_TAG = "__callable__"
_PATH_TAG = "__path__"
_DATETIME_TAG = "__datetime__"
_DATE_TAG = "__date__"
_TUPLE_TAG = "__tuple__"


def _qualname(obj: Any) -> str:
    return f"{obj.__module__}:{obj.__qualname__}"


def _resolve_path(reference: str) -> Any:
    module_name, sep, qualname = reference.partition(":")
    if not sep:
        raise ValueError(f"Invalid reference path: {reference}")
    module = importlib.import_module(module_name)
    current: Any = module
    for part in qualname.split("."):
        current = getattr(current, part)
    return current


def serialize_for_json(value: Any) -> Any:
    if is_dataclass(value):
        payload = {k: serialize_for_json(v) for k, v in asdict(value).items()}
        payload[_DATACLASS_TAG] = _qualname(type(value))
        return payload
    if isinstance(value, tuple):
        return {_TUPLE_TAG: [serialize_for_json(item) for item in value]}
    if isinstance(value, list):
        return [serialize_for_json(item) for item in value]
    if isinstance(value, dict):
        return {str(k): serialize_for_json(v) for k, v in value.items()}
    if isinstance(value, Path):
        return {_PATH_TAG: str(value)}
    if isinstance(value, datetime):
        return {_DATETIME_TAG: value.isoformat()}
    if isinstance(value, date):
        return {_DATE_TAG: value.isoformat()}
    if callable(value):
        return {_CALLABLE_TAG: _qualname(value)}
    return value


def deserialize_from_json(value: Any) -> Any:
    if isinstance(value, list):
        return [deserialize_from_json(item) for item in value]
    if isinstance(value, dict):
        if _TUPLE_TAG in value:
            return tuple(deserialize_from_json(item) for item in value[_TUPLE_TAG])
        if _PATH_TAG in value:
            return Path(str(value[_PATH_TAG]))
        if _DATETIME_TAG in value:
            return datetime.fromisoformat(str(value[_DATETIME_TAG]))
        if _DATE_TAG in value:
            return date.fromisoformat(str(value[_DATE_TAG]))
        if _CALLABLE_TAG in value:
            return _resolve_path(str(value[_CALLABLE_TAG]))
        if _DATACLASS_TAG in value:
            cls = _resolve_path(str(value[_DATACLASS_TAG]))
            if not is_dataclass(cls):
                raise ValueError(f"Resolved class is not a dataclass: {cls}")
            kwargs: dict[str, Any] = {}
            for dataclass_field in fields(cls):
                if dataclass_field.name in value:
                    kwargs[dataclass_field.name] = deserialize_from_json(value[dataclass_field.name])
            return cls(**kwargs)
        return {k: deserialize_from_json(v) for k, v in value.items()}
    return value

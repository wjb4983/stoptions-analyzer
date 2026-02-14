from __future__ import annotations

import sys
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
for path in (ROOT, SRC):
    path_str = str(path)
    if path_str not in sys.path:
        sys.path.insert(0, path_str)


SMOKE_TEST_FILES = {
    "tests/test_feature_store.py",
    "tests/test_vectorized_no_lookahead.py",
    "tests/test_backtest_invariants_and_edge_cases.py",
    "tests/test_validation.py",
}

CORE_TEST_FILES = {
    *SMOKE_TEST_FILES,
    "tests/test_quality_gate_artifacts.py",
    "tests/test_benchmark_bundle.py",
    "tests/test_governance_artifact_validator.py",
    "tests/test_performance_quality_thresholds.py",
}

SLOW_TEST_PATH_SNIPPETS = (
    "modeling_nextgen/test_nextgen_integration_regression.py",
    "test_stress_",
    "test_feature_store_performance.py",
)


def pytest_collection_modifyitems(items: list[pytest.Item]) -> None:
    for item in items:
        node_path = Path(str(item.fspath)).resolve().as_posix()
        try:
            rel_path = Path(node_path).relative_to(ROOT).as_posix()
        except ValueError:
            rel_path = node_path

        if rel_path in SMOKE_TEST_FILES:
            item.add_marker(pytest.mark.smoke)

        if rel_path in CORE_TEST_FILES:
            item.add_marker(pytest.mark.core)

        if any(snippet in rel_path for snippet in SLOW_TEST_PATH_SNIPPETS):
            item.add_marker(pytest.mark.slow)

from __future__ import annotations

from datetime import datetime

from src.data_access.provider_base import DataProvider, FORWARD_KNOWN_FIELD_NAMES


class CompliantFakeProvider:
    def list_symbols(self) -> list[str]:
        return ["AAPL"]

    def get_bars(
        self,
        symbols: list[str],
        start: datetime | str,
        end: datetime | str,
        timeframe: str = "1m",
    ) -> object:
        return []

    def get_corporate_actions(
        self,
        symbols: list[str],
        start: datetime | str,
        end: datetime | str,
    ) -> object:
        return []


class NonCompliantFakeProvider:
    def list_symbols(self) -> list[str]:
        return ["AAPL"]


def test_runtime_checkable_protocol_accepts_compliant_fake() -> None:
    fake = CompliantFakeProvider()

    assert isinstance(fake, DataProvider)


def test_runtime_checkable_protocol_rejects_non_compliant_fake() -> None:
    fake = NonCompliantFakeProvider()

    assert not isinstance(fake, DataProvider)


def test_forward_known_field_names_contains_expected_leak_prone_labels() -> None:
    expected_labels = {
        "future_return",
        "future_close",
        "next_close",
        "next_open",
        "label",
        "target",
        "target_return",
        "alpha_label",
        "lookahead_return",
    }

    assert expected_labels.issubset(FORWARD_KNOWN_FIELD_NAMES)


def test_forward_known_field_names_snapshot_sorted() -> None:
    # Snapshot-style guardrail: intentional set changes must update this list in PR diff.
    assert sorted(FORWARD_KNOWN_FIELD_NAMES) == [
        "alpha_label",
        "future_close",
        "future_return",
        "label",
        "lookahead_return",
        "next_close",
        "next_open",
        "target",
        "target_return",
    ]

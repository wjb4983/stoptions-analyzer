from __future__ import annotations

import json
from urllib.error import HTTPError

import pytest

from src.data_access.api_client import MassiveApiClient


class _FakeResponse:
    def __init__(self, payload: str, status: int = 200):
        self._payload = payload.encode("utf-8")
        self.status = status

    def __enter__(self):
        if self.status != 200:
            raise HTTPError(url="https://example.test", code=self.status, msg="boom", hdrs=None, fp=None)
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def read(self) -> bytes:
        return self._payload


def test_request_uses_expected_timeout(monkeypatch: pytest.MonkeyPatch) -> None:
    captured: dict[str, object] = {}

    def _fake_urlopen(url: str, timeout: int):
        captured["url"] = url
        captured["timeout"] = timeout
        return _FakeResponse('{"results": []}')

    monkeypatch.setattr("src.data_access.api_client.urlopen", _fake_urlopen)
    client = MassiveApiClient(api_key="k", base_url="https://example.test")

    payload = client._request("/v1/data", {"ticker": "AAPL"})

    assert payload == {"results": []}
    assert captured["timeout"] == 10
    assert "apiKey=k" in str(captured["url"])


def test_request_surfaces_non_200_http_errors(monkeypatch: pytest.MonkeyPatch) -> None:
    def _fake_urlopen(url: str, timeout: int):
        return _FakeResponse("{}", status=503)

    monkeypatch.setattr("src.data_access.api_client.urlopen", _fake_urlopen)
    client = MassiveApiClient(api_key="k", base_url="https://example.test")

    with pytest.raises(HTTPError):
        client._request("/v1/data", {})


def test_request_rejects_malformed_json_payload(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        "src.data_access.api_client.urlopen",
        lambda url, timeout: _FakeResponse("not-json"),
    )
    client = MassiveApiClient(api_key="k", base_url="https://example.test")

    with pytest.raises(json.JSONDecodeError):
        client._request("/v1/data", {})


def test_retry_backoff_not_implemented_single_attempt(monkeypatch: pytest.MonkeyPatch) -> None:
    calls = {"count": 0}

    def _fake_urlopen(url: str, timeout: int):
        calls["count"] += 1
        raise TimeoutError("simulated timeout")

    monkeypatch.setattr("src.data_access.api_client.urlopen", _fake_urlopen)
    client = MassiveApiClient(api_key="k", base_url="https://example.test")

    with pytest.raises(TimeoutError):
        client._request("/v1/data", {})

    assert calls["count"] == 1

"""Shared test fixtures — a fake Graph transport, no network, no real tokens.

FakeGraph replaces graph.request(): tests queue canned responses per
(method, path-prefix) and assert on the calls the tools actually made.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import graph  # noqa: E402


class FakeGraph:
    def __init__(self):
        self.calls: list[dict] = []
        self._responses: list[tuple[str, str, object]] = []

    def queue(self, method: str, path_prefix: str, response):
        """Queue a response for the next call matching method + path prefix."""
        self._responses.append((method.upper(), path_prefix, response))

    def request(self, method, path, account=None, params=None, json_body=None,
                headers=None, data=None, raw=False, auth=True):
        call = {
            "method": method.upper(), "path": path, "account": account,
            "params": params, "json_body": json_body, "headers": headers,
            "data": data, "raw": raw, "auth": auth,
        }
        self.calls.append(call)
        for i, (m, prefix, resp) in enumerate(self._responses):
            if m == method.upper() and path.startswith(prefix):
                self._responses.pop(i)
                if isinstance(resp, Exception):
                    raise resp
                return resp
        return {} if not raw else b""


@pytest.fixture
def fake_graph(monkeypatch):
    fake = FakeGraph()
    monkeypatch.setattr(graph, "request", fake.request)
    # get_all is a thin pagination loop over request(); patching request alone
    # covers it, but tools imported `graph` as a module so both resolve live.
    return fake

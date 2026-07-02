"""Thin Microsoft Graph REST helper shared by all tool modules.

One seam for every Graph call: request() / get_all(). Tests monkeypatch
request() with a fake — no tool module ever imports `requests` directly.

Also owns the per-call audit log (4-field JSONL schema: execution_time_ms,
io, token_usage, error_class) at ~/.claude/microsoft-365-mcp/audit.log.jsonl.
Input/output payloads are summarized, not dumped — this MCP handles a user's
corporate mail; full io capture is opt-in via M365_AUDIT_IO=1.
"""

from __future__ import annotations

import datetime
import json
import os
import time
from pathlib import Path
from typing import Any

import requests

from accounts import get_token

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TIMEOUT = 30

_AUDIT_DIR = Path.home() / ".claude" / "microsoft-365-mcp"
_AUDIT_LOG = _AUDIT_DIR / "audit.log.jsonl"

_ERROR_CLASSES = {
    401: "auth",
    403: "auth",
    404: "validation",
    400: "validation",
    429: "rate_limit",
}


class GraphError(Exception):
    """Graph API returned an error. Message is safe to surface (no tokens)."""

    def __init__(self, status: int, code: str, message: str):
        self.status = status
        self.code = code
        self.error_class = _ERROR_CLASSES.get(status, "upstream_error")
        super().__init__(f"Graph {status} {code}: {message}")


def _raise_for(resp: requests.Response) -> None:
    try:
        err = resp.json().get("error", {})
    except ValueError:
        err = {}
    raise GraphError(
        resp.status_code,
        err.get("code", "unknown"),
        err.get("message", resp.text[:300]),
    )


def request(
    method: str,
    path: str,
    account: str | None = None,
    params: dict | None = None,
    json_body: dict | None = None,
    headers: dict | None = None,
    data: bytes | None = None,
    raw: bool = False,
    auth: bool = True,
) -> Any:
    """Make one Graph call. `path` is relative ('/me/messages') or absolute
    (a @odata.nextLink). Returns parsed JSON ({} on 204), or bytes when raw=True.
    Retries once on 429/503 honoring Retry-After.

    auth=False sends NO Authorization header. Required for OneDrive upload-session
    chunk PUTs: the uploadUrl is pre-authenticated and Graph rejects (401) a
    bearer token on it (per the large-file-upload spec).
    """
    url = path if path.startswith("http") else f"{GRAPH_BASE}{path}"
    hdrs = {"Authorization": f"Bearer {get_token(account)}"} if auth else {}
    if headers:
        hdrs.update(headers)

    for attempt in (1, 2):
        resp = requests.request(
            method, url, params=params, json=json_body, data=data,
            headers=hdrs, timeout=TIMEOUT,
        )
        if resp.status_code in (429, 503) and attempt == 1:
            time.sleep(min(int(resp.headers.get("Retry-After", "2")), 30))
            continue
        break

    if resp.status_code >= 400:
        _raise_for(resp)
    if raw:
        return resp.content
    if resp.status_code == 204 or not resp.content:
        return {}
    return resp.json()


def get_all(
    path: str,
    account: str | None = None,
    params: dict | None = None,
    limit: int = 50,
    headers: dict | None = None,
) -> list[dict]:
    """GET with @odata.nextLink pagination, up to `limit` items."""
    items: list[dict] = []
    page = request("GET", path, account=account, params=params, headers=headers)
    while True:
        items.extend(page.get("value", []))
        next_link = page.get("@odata.nextLink")
        if len(items) >= limit or not next_link:
            break
        page = request("GET", next_link, account=account)
    return items[:limit]


# --- Per-call observability (4-field JSONL) ---------------------------------


def _preview(value: Any, cap: int = 120) -> Any:
    if isinstance(value, str):
        return value if len(value) <= cap else value[:cap] + "…"
    if isinstance(value, (int, float, bool)) or value is None:
        return value
    if isinstance(value, list):
        return f"list[{len(value)}]"
    if isinstance(value, dict):
        return f"dict[{len(value)}]"
    return type(value).__name__


def audit(
    tool: str,
    kwargs: dict,
    started: float,
    error_class: str | None = None,
    result: Any = None,
) -> None:
    """Append one 4-field observability record. Never raises."""
    try:
        _AUDIT_DIR.mkdir(parents=True, exist_ok=True)
        if os.environ.get("M365_AUDIT_IO") == "1":
            io = {"input": kwargs, "output": result}
        else:
            io = {
                "input": {k: _preview(v) for k, v in kwargs.items()},
                "output": _preview(result),
            }
        record = {
            "ts": datetime.datetime.now(datetime.timezone.utc).isoformat(),
            "tool": tool,
            "execution_time_ms": int((time.monotonic() - started) * 1000),
            "io": io,
            "token_usage": None,  # no LLM calls in this MCP
            "error_class": error_class,
        }
        with open(_AUDIT_LOG, "a") as f:
            f.write(json.dumps(record, default=str) + "\n")
    except Exception:
        pass

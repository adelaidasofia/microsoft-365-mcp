"""Server-import smoke test.

Imports the FastMCP server entry module and asserts it constructs. This is the
runtime-verification gate the install step alone does not provide: a dependency
bump (e.g. a fastmcp / msal major) that breaks tool registration or server
construction fails here instead of shipping silently.
"""
from __future__ import annotations

import asyncio
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def test_server_module_constructs() -> None:
    from fastmcp import FastMCP
    import server
    assert isinstance(server.mcp, FastMCP), f"server.mcp is not FastMCP: {type(server.mcp)!r}"


def test_all_tools_registered_with_schemas() -> None:
    import server
    tools = asyncio.run(server.mcp.list_tools())
    names = {t.name for t in tools}
    assert len(names) == 33, f"expected 33 tools, got {len(names)}"
    for family, expect in [("m365_account_", 4), ("outlook_", 9), ("mscal_", 7), ("onedrive_", 13)]:
        got = sum(1 for n in names if n.startswith(family))
        assert got == expect, f"{family}*: expected {expect}, got {got}"
    # The observability wrapper must not erase real parameter schemas.
    search = next(t for t in tools if t.name == "outlook_search")
    assert "query" in search.parameters["properties"]
    assert "kwargs" not in search.parameters["properties"]

"""The MCP surface: a server assembled from tool groups around one service."""

from __future__ import annotations

from fastmcp import FastMCP

from .service import DocumentService
from .tools import REGISTRARS, escape


def create_server(service: DocumentService, *, allow_raw: bool = False) -> FastMCP:
    """Build the MCP server around an already constructed service.

    The service is injected rather than discovered, so the tools can be exercised
    in tests against a stub engine. The escape hatch is registered only when it is
    allowed, so a client that may not run arbitrary code never sees the tool.
    """
    mcp = FastMCP("word")
    for register in REGISTRARS:
        register(mcp, service)
    if allow_raw:
        escape.register(mcp, service)
    return mcp

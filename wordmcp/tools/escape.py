"""The escape hatch, registered only when it is explicitly allowed."""

from __future__ import annotations

from fastmcp import FastMCP

from ..service import DocumentService


def register(mcp: FastMCP, service: DocumentService) -> None:
    """Add the raw execution tool to a server."""

    @mcp.tool
    def word_run_raw(code: str) -> str:
        """Run engine-native code against the live session.

        On macOS the code is AppleScript. On Windows and Linux it is a Python
        expression, or a block that assigns `result`, evaluated with the live
        application objects in scope. Intended for diagnosing a failing call
        without editing the server.
        """
        return service.raw(code)

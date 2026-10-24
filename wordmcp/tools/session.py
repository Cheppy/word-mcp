"""Tools that govern which document is current."""

from __future__ import annotations

from fastmcp import FastMCP

from ..models import EngineStatus
from ..service import DocumentService


def register(mcp: FastMCP, service: DocumentService) -> None:
    """Add the session tools to a server."""

    @mcp.tool
    def word_status() -> EngineStatus:
        """Report which word processor drives documents here and whether it answers."""
        return service.status()

    @mcp.tool
    def word_open(path: str) -> str:
        """Open a document and make it the current one."""
        return service.open(path)

    @mcp.tool
    def word_close(save: bool = False) -> str:
        """Close the current document, saving it or discarding changes."""
        service.close(save)
        return "saved and closed" if save else "closed without saving"

    @mcp.tool
    def word_save() -> str:
        """Save the current document in place."""
        service.save()
        return "saved"

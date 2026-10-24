"""Tools that read the layout without changing the document."""

from __future__ import annotations

from fastmcp import FastMCP

from ..models import CheckResult, Heading, Occurrences, Outline, PageText
from ..service import DocumentService


def register(mcp: FastMCP, service: DocumentService) -> None:
    """Add the measuring tools to a server."""

    @mcp.tool
    def word_page_count(path: str | None = None) -> int:
        """Count the pages of a document as the word processor lays it out."""
        return service.page_count(path)

    @mcp.tool
    def word_find_page(text: str, path: str | None = None, limit: int = 20) -> Occurrences:
        """List the pages on which a piece of text occurs, in document order."""
        return service.find(text, path, limit)

    @mcp.tool
    def word_headings(path: str | None = None, max_level: int = 3) -> list[Heading]:
        """List every heading with its outline level and the page it starts on."""
        return service.headings(path, max_level)

    @mcp.tool
    def word_outline(path: str | None = None, markers: list[str] | None = None) -> Outline:
        """Give the start page and length of each top-level part of the document.

        Pass markers to define the parts by text instead of by heading level.
        """
        return service.outline(path, markers)

    @mcp.tool
    def word_page_text(page: int, path: str | None = None) -> PageText:
        """Read the text that the word processor placed on a given page."""
        return service.page_text(page, path)

    @mcp.tool
    def word_selftest(path: str) -> list[CheckResult]:
        """Try every operation on a document and report which ones this machine supports."""
        return service.selftest(path)

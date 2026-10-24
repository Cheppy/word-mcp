"""Tools that change the document or produce a file from it."""

from __future__ import annotations

from fastmcp import FastMCP

from ..models import ExportResult, ParagraphFormat, UpdateResult
from ..service import DocumentService


def register(mcp: FastMCP, service: DocumentService) -> None:
    """Add the modifying tools to a server."""

    @mcp.tool
    def word_update_fields(toc: bool = True, path: str | None = None) -> UpdateResult:
        """Refresh fields and tables of contents so their page numbers become current."""
        return service.update_fields(toc, path)

    @mcp.tool
    def word_format_paragraphs(
        contains: str,
        keep_with_next: bool | None = None,
        keep_together: bool | None = None,
        widow_control: bool | None = None,
        space_before: float | None = None,
        space_after: float | None = None,
        path: str | None = None,
    ) -> int:
        """Set layout properties on every paragraph containing a string.

        keep_with_next ties a heading or a caption to what follows, keep_together
        forbids splitting a paragraph across pages, widow_control switches widow and
        orphan control on. Spacing is in points. Returns how many paragraphs changed.
        """
        fmt = ParagraphFormat(
            keep_with_next=keep_with_next,
            keep_together=keep_together,
            widow_control=widow_control,
            space_before=space_before,
            space_after=space_after,
        )
        return service.format_paragraphs(contains, fmt, path)

    @mcp.tool
    def word_export_pdf(path: str | None = None, pdf_path: str | None = None) -> ExportResult:
        """Export the document to PDF with the word processor's own engine."""
        return service.export_pdf(path, pdf_path)

"""MCP surface. Tools are thin: they validate, delegate to the service, return models."""

from __future__ import annotations

import os
from collections.abc import Mapping

from fastmcp import FastMCP

from .engines import create_engine
from .models import (
    CheckResult,
    EngineStatus,
    ExportResult,
    Heading,
    Occurrences,
    Outline,
    PageText,
    ParagraphFormat,
    UpdateResult,
)
from .service import DocumentService

RAW_VARIABLE = "WORD_MCP_ALLOW_RAW"


def create_server(service: DocumentService, *, allow_raw: bool = False) -> FastMCP:
    """Build the MCP server around an already constructed service.

    The service is injected rather than discovered, so the tools can be exercised
    in tests against a stub engine.
    """
    mcp = FastMCP("word")

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
    def word_update_fields(toc: bool = True) -> UpdateResult:
        """Refresh fields and tables of contents so their page numbers become current."""
        return service.update_fields(toc)

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

    @mcp.tool
    def word_selftest(path: str) -> list[CheckResult]:
        """Try every operation on a document and report which ones this machine supports."""
        return service.selftest(path)

    if allow_raw:

        @mcp.tool
        def word_run_raw(code: str) -> str:
            """Run engine-native code against the live session.

            On macOS the code is AppleScript. On Windows and Linux it is a Python
            expression, or a block that assigns `result`, evaluated with the live
            application objects in scope. Intended for diagnosing a failing call
            without editing the server.
            """
            return service.raw(code)

    return mcp


def build_default_service(env: Mapping[str, str] | None = None) -> DocumentService:
    """Construct the service for this machine from the environment."""
    environment = os.environ if env is None else env
    return DocumentService(create_engine(env=environment))


def main() -> None:
    """Entry point: wire the default service and serve over stdio."""
    environment = os.environ
    server = create_server(
        build_default_service(environment),
        allow_raw=environment.get(RAW_VARIABLE) == "1",
    )
    server.run()


if __name__ == "__main__":
    main()

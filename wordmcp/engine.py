"""The contract every document engine fulfils."""

from __future__ import annotations

from abc import ABC, abstractmethod

from .models import EngineStatus, Heading, ParagraphFormat


class DocumentEngine(ABC):
    """A word processor driven from outside, holding at most one open document.

    Implementations differ only in how they reach the application: Apple Events on
    macOS, COM on Windows, UNO for LibreOffice. `fidelity` says whose pagination the
    numbers carry, since LibreOffice and Word break pages differently.
    """

    name: str = "engine"
    fidelity: str = "unknown"

    @abstractmethod
    def status(self) -> EngineStatus:
        """Report whether the application is installed, reachable and what is open."""

    @abstractmethod
    def open(self, path: str) -> str:
        """Open a document and make it current, returning its resolved path."""

    @abstractmethod
    def close(self, save: bool = False) -> None:
        """Close the current document."""

    @abstractmethod
    def save(self) -> None:
        """Save the current document in place."""

    @abstractmethod
    def page_count(self) -> int:
        """Count the pages of the current document."""

    @abstractmethod
    def find_pages(self, text: str, limit: int = 20) -> list[int]:
        """Pages on which a search string occurs, in document order."""

    @abstractmethod
    def headings(self, max_level: int = 3) -> list[Heading]:
        """Heading paragraphs up to an outline level, with their pages."""

    @abstractmethod
    def page_text(self, page: int) -> str:
        """Text placed on one page."""

    @abstractmethod
    def update_fields(self, toc: bool = True) -> list[str]:
        """Refresh fields and, optionally, tables of contents."""

    @abstractmethod
    def format_paragraphs(self, contains: str, fmt: ParagraphFormat) -> int:
        """Apply layout properties to paragraphs containing a string, returning the count."""

    @abstractmethod
    def export_pdf(self, pdf_path: str) -> str:
        """Export the current document to PDF with the application's own engine."""

    @abstractmethod
    def raw(self, code: str) -> str:
        """Run engine-native code against the live session and return its output."""


class EngineError(RuntimeError):
    """The application refused a request or could not be reached."""


class NoDocumentError(EngineError):
    """An operation needs an open document and none is current."""

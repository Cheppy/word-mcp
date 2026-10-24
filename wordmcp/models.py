"""Value objects exchanged between the engines, the service and the MCP tools."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class EngineStatus:
    """What is driving documents on this machine and whether it answers."""

    engine: str
    fidelity: str
    available: bool
    version: str | None = None
    document: str | None = None
    detail: str | None = None


@dataclass(frozen=True)
class Heading:
    """A heading paragraph and the page it starts on."""

    level: int
    page: int
    text: str


@dataclass(frozen=True)
class Section:
    """A top-level part of the document with its extent in pages."""

    name: str
    page: int
    pages: int


@dataclass(frozen=True)
class Outline:
    """Section map of a document."""

    total_pages: int
    fidelity: str
    sections: list[Section]


@dataclass(frozen=True)
class PageText:
    """Text the engine placed on one page."""

    page: int
    text: str


@dataclass(frozen=True)
class Occurrences:
    """Pages on which a search string occurs, in document order."""

    text: str
    pages: list[int]


@dataclass(frozen=True)
class ExportResult:
    """Where the exported PDF landed."""

    pdf_path: str
    bytes: int


@dataclass(frozen=True)
class UpdateResult:
    """Outcome of refreshing fields and tables of contents."""

    steps: list[str]


@dataclass(frozen=True)
class ParagraphFormat:
    """Layout properties applied to a set of paragraphs.

    Spacing is in points. A property left as None is not touched.
    """

    keep_with_next: bool | None = None
    keep_together: bool | None = None
    widow_control: bool | None = None
    space_before: float | None = None
    space_after: float | None = None

    def is_empty(self) -> bool:
        """Say whether the format would change nothing."""
        return all(
            getattr(self, f) is None
            for f in (
                "keep_with_next",
                "keep_together",
                "widow_control",
                "space_before",
                "space_after",
            )
        )


@dataclass(frozen=True)
class CheckResult:
    """Outcome of one probe in the self-test."""

    step: str
    ok: bool
    detail: str

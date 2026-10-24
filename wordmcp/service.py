"""Document operations expressed once, independently of the engine behind them."""

from __future__ import annotations

from pathlib import Path

from .engine import DocumentEngine
from .models import (
    CheckResult,
    EngineStatus,
    ExportResult,
    Heading,
    Occurrences,
    Outline,
    PageText,
    ParagraphFormat,
    Section,
    UpdateResult,
)


class DocumentService:
    """Everything the tools offer, built on one injected engine."""

    def __init__(self, engine: DocumentEngine) -> None:
        """Bind the service to the engine it will drive."""
        self._engine = engine

    @property
    def engine(self) -> DocumentEngine:
        """The engine this service was built with."""
        return self._engine

    def _use(self, path: str | None) -> None:
        if path:
            self._engine.open(path)

    def status(self) -> EngineStatus:
        """Which engine is active and whether it answers."""
        return self._engine.status()

    def open(self, path: str) -> str:
        """Open a document and make it current."""
        return self._engine.open(path)

    def close(self, save: bool = False) -> None:
        """Close the current document."""
        self._engine.close(save)

    def save(self) -> None:
        """Save the current document in place."""
        self._engine.save()

    def page_count(self, path: str | None = None) -> int:
        """Pages in the document."""
        self._use(path)
        return self._engine.page_count()

    def find(self, text: str, path: str | None = None, limit: int = 20) -> Occurrences:
        """Pages on which a string occurs."""
        self._use(path)
        return Occurrences(text, self._engine.find_pages(text, limit))

    def headings(self, path: str | None = None, max_level: int = 3) -> list[Heading]:
        """Headings with their outline level and page."""
        self._use(path)
        return self._engine.headings(max_level)

    def outline(self, path: str | None = None, markers: list[str] | None = None) -> Outline:
        """Start page and extent of each top-level part.

        Without markers the document outline is used, so every level one heading
        opens a section. With markers the last occurrence of each string is taken,
        which skips mentions in a table of contents.
        """
        self._use(path)
        total = self._engine.page_count()
        if markers:
            starts = [
                (marker, pages[-1])
                for marker in markers
                if (pages := self._engine.find_pages(marker, limit=20))
            ]
        else:
            starts = [(h.text, h.page) for h in self._engine.headings(max_level=1)]
        starts.sort(key=lambda item: item[1])
        sections = [
            Section(name, page, (starts[i + 1][1] if i + 1 < len(starts) else total + 1) - page)
            for i, (name, page) in enumerate(starts)
        ]
        return Outline(total, self._engine.fidelity, sections)

    def page_text(self, page: int, path: str | None = None) -> PageText:
        """Text placed on one page."""
        self._use(path)
        return PageText(page, self._engine.page_text(page))

    def update_fields(self, toc: bool = True) -> UpdateResult:
        """Refresh fields and tables of contents."""
        return UpdateResult(self._engine.update_fields(toc))

    def format_paragraphs(
        self, contains: str, fmt: ParagraphFormat, path: str | None = None
    ) -> int:
        """Apply layout properties to paragraphs containing a string."""
        self._use(path)
        return self._engine.format_paragraphs(contains, fmt)

    def export_pdf(self, path: str | None = None, pdf_path: str | None = None) -> ExportResult:
        """Export to PDF, defaulting to the document path with a pdf suffix."""
        self._use(path)
        target = pdf_path
        if not target:
            source = path or getattr(self._engine, "_path", None)
            if not source:
                raise ValueError("a destination is needed when no document path is known")
            target = str(Path(source).with_suffix(".pdf"))
        written = self._engine.export_pdf(target)
        return ExportResult(written, Path(written).stat().st_size)

    def raw(self, code: str) -> str:
        """Run engine-native code against the live session."""
        return self._engine.raw(code)

    def selftest(self, path: str) -> list[CheckResult]:
        """Probe every operation on a document and report what this machine supports."""
        probes = [
            ("status", lambda: str(self._engine.status())),
            ("open", lambda: self._engine.open(path)),
            ("page_count", lambda: f"{self._engine.page_count()} pages"),
            ("find_pages", lambda: f"{self._engine.find_pages('ВСТУП', 5)}"),
            ("headings", lambda: f"{len(self._engine.headings(2))} headings"),
            ("page_text", lambda: f"{len(self._engine.page_text(1))} characters"),
        ]
        results = []
        for step, probe in probes:
            try:
                results.append(CheckResult(step, True, probe()))
            except Exception as exc:
                results.append(CheckResult(step, False, str(exc)))
        return results

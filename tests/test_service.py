"""Service behaviour, checked against a stub engine instead of a word processor."""

from __future__ import annotations

import pytest

from wordmcp.engine import DocumentEngine
from wordmcp.engines import resolve_name
from wordmcp.models import EngineStatus, Heading, ParagraphFormat
from wordmcp.service import DocumentService


class StubEngine(DocumentEngine):
    """An engine that answers from fixtures and records what it was asked."""

    name = "stub"
    fidelity = "stub"

    def __init__(self, pages: int = 10, headings: list[Heading] | None = None) -> None:
        self.pages = pages
        self._headings = headings or []
        self.opened: list[str] = []
        self.formatted: list[tuple[str, ParagraphFormat]] = []

    def status(self) -> EngineStatus:
        return EngineStatus(self.name, self.fidelity, True)

    def open(self, path: str) -> str:
        self.opened.append(path)
        return path

    def close(self, save: bool = False) -> None:
        return None

    def save(self) -> None:
        return None

    def page_count(self) -> int:
        return self.pages

    def find_pages(self, text: str, limit: int = 20) -> list[int]:
        return {"ВСТУП": [2, 14], "РОЗДІЛ 1": [3, 24]}.get(text, [])

    def headings(self, max_level: int = 3) -> list[Heading]:
        return [h for h in self._headings if h.level <= max_level]

    def page_text(self, page: int) -> str:
        return f"page {page}"

    def update_fields(self, toc: bool = True) -> list[str]:
        return ["fields refreshed"] + (["tables of contents updated"] if toc else [])

    def format_paragraphs(self, contains: str, fmt: ParagraphFormat) -> int:
        self.formatted.append((contains, fmt))
        return 3

    def export_pdf(self, pdf_path: str) -> str:
        return pdf_path

    def raw(self, code: str) -> str:
        return code


def test_outline_from_headings_measures_each_section() -> None:
    engine = StubEngine(
        pages=40,
        headings=[Heading(1, 1, "ВСТУП"), Heading(1, 11, "РОЗДІЛ 1"), Heading(1, 26, "ВИСНОВКИ")],
    )
    outline = DocumentService(engine).outline()
    assert [s.pages for s in outline.sections] == [10, 15, 15]
    assert outline.total_pages == 40


def test_outline_with_markers_takes_the_last_occurrence() -> None:
    outline = DocumentService(StubEngine(pages=30)).outline(markers=["ВСТУП", "РОЗДІЛ 1"])
    assert [(s.name, s.page) for s in outline.sections] == [("ВСТУП", 14), ("РОЗДІЛ 1", 24)]


def test_a_path_opens_the_document_once_per_call() -> None:
    engine = StubEngine()
    service = DocumentService(engine)
    service.page_count("/tmp/a.docx")
    service.page_count()
    assert engine.opened == ["/tmp/a.docx"]


def test_empty_format_is_refused_before_reaching_the_engine() -> None:
    assert ParagraphFormat().is_empty()


def test_selftest_reports_every_probe() -> None:
    results = DocumentService(StubEngine()).selftest("/tmp/a.docx")
    assert [r.step for r in results] == [
        "status",
        "open",
        "page_count",
        "find_pages",
        "headings",
        "page_text",
    ]
    assert all(r.ok for r in results)


@pytest.mark.parametrize(
    ("platform", "expected"),
    [("darwin", "applescript"), ("win32", "com"), ("linux", "uno")],
)
def test_engine_follows_the_platform(platform: str, expected: str) -> None:
    assert resolve_name(platform=platform, env={}) == expected


def test_environment_overrides_the_platform() -> None:
    assert resolve_name(platform="darwin", env={"WORD_MCP_BACKEND": "uno"}) == "uno"


def test_unknown_engine_is_rejected() -> None:
    with pytest.raises(ValueError):
        resolve_name("nope", env={})

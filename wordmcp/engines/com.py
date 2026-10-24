"""Microsoft Word on Windows, reached through COM, the object model VBA uses."""

from __future__ import annotations

import importlib.util
from pathlib import Path

from ..engine import DocumentEngine, EngineError, NoDocumentError
from ..models import EngineStatus, Heading, ParagraphFormat

STATISTIC_PAGES = 2
INFORMATION_PAGE_NUMBER = 3
GOTO_PAGE = 1
GOTO_ABSOLUTE = 1
EXPORT_FORMAT_PDF = 17
FIND_WRAP_STOP = 0
SAVE_CHANGES_YES = -1
SAVE_CHANGES_NO = 0


class ComEngine(DocumentEngine):
    """Drives Word for Windows invisibly; page numbers are Word's own."""

    name = "word-win"
    fidelity = "word"

    def __init__(self) -> None:
        """Start with no application and no document bound."""
        self._app = None
        self._doc = None
        self._path: str | None = None

    def _application(self):
        if self._app is None:
            import win32com.client

            self._app = win32com.client.Dispatch("Word.Application")
            self._app.Visible = False
            self._app.DisplayAlerts = 0
        return self._app

    def _document(self):
        if self._doc is None:
            raise NoDocumentError("no document is open")
        return self._doc

    def _prepare_find(self, span, text: str):
        finder = span.Find
        finder.ClearFormatting()
        finder.Text = text
        finder.MatchCase = True
        finder.Forward = True
        finder.Wrap = FIND_WRAP_STOP
        return finder

    def status(self) -> EngineStatus:
        if importlib.util.find_spec("win32com") is None:
            return EngineStatus(
                self.name, self.fidelity, False, detail="pywin32 is not installed"
            )
        try:
            version = str(self._application().Version)
        except Exception as exc:
            return EngineStatus(self.name, self.fidelity, False, detail=str(exc))
        return EngineStatus(self.name, self.fidelity, True, version=version, document=self._path)

    def open(self, path: str) -> str:
        resolved = Path(path).expanduser().resolve()
        if not resolved.is_file():
            raise FileNotFoundError(resolved)
        self.close()
        self._doc = self._application().Documents.Open(
            str(resolved), ReadOnly=False, AddToRecentFiles=False
        )
        self._path = str(resolved)
        return self._path

    def close(self, save: bool = False) -> None:
        if self._doc is None:
            return
        self._doc.Close(SaveChanges=SAVE_CHANGES_YES if save else SAVE_CHANGES_NO)
        self._doc = None
        self._path = None

    def save(self) -> None:
        self._document().Save()

    def page_count(self) -> int:
        return int(self._document().ComputeStatistics(STATISTIC_PAGES))

    def find_pages(self, text: str, limit: int = 20) -> list[int]:
        document = self._document()
        span = document.Content
        pages: list[int] = []
        while len(pages) < limit:
            if not self._prepare_find(span, text).Execute():
                break
            pages.append(int(span.Information(INFORMATION_PAGE_NUMBER)))
            if span.End >= document.Content.End:
                break
            span.SetRange(span.End, document.Content.End)
        return pages

    def headings(self, max_level: int = 3) -> list[Heading]:
        found: list[Heading] = []
        for paragraph in self._document().Paragraphs:
            level = int(paragraph.OutlineLevel)
            if level <= max_level:
                found.append(
                    Heading(
                        level,
                        int(paragraph.Range.Information(INFORMATION_PAGE_NUMBER)),
                        " ".join(str(paragraph.Range.Text).split()),
                    )
                )
        return found

    def page_text(self, page: int) -> str:
        span = self._document().GoTo(
            What=GOTO_PAGE, Which=GOTO_ABSOLUTE, Count=int(page)
        )
        return str(span.Bookmarks("\\Page").Range.Text)

    def update_fields(self, toc: bool = True) -> list[str]:
        document = self._document()
        steps: list[str] = []
        try:
            document.Fields.Update()
            steps.append("fields refreshed")
        except Exception as exc:
            steps.append(f"fields failed: {exc}")
        if toc:
            try:
                count = document.TablesOfContents.Count
                for position in range(1, count + 1):
                    document.TablesOfContents(position).Update()
                steps.append(f"tables of contents updated: {count}")
            except Exception as exc:
                steps.append(f"tables of contents failed: {exc}")
        return steps

    def format_paragraphs(self, contains: str, fmt: ParagraphFormat) -> int:
        if fmt.is_empty():
            raise ValueError("no property to set")
        changed = 0
        for paragraph in self._document().Paragraphs:
            if contains not in str(paragraph.Range.Text):
                continue
            style = paragraph.Format
            if fmt.keep_with_next is not None:
                style.KeepWithNext = fmt.keep_with_next
            if fmt.keep_together is not None:
                style.KeepTogether = fmt.keep_together
            if fmt.widow_control is not None:
                style.WidowControl = fmt.widow_control
            if fmt.space_before is not None:
                style.SpaceBefore = fmt.space_before
            if fmt.space_after is not None:
                style.SpaceAfter = fmt.space_after
            changed += 1
        return changed

    def export_pdf(self, pdf_path: str) -> str:
        target = Path(pdf_path).expanduser().resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        self._document().ExportAsFixedFormat(
            OutputFileName=str(target), ExportFormat=EXPORT_FORMAT_PDF
        )
        if not target.is_file():
            raise EngineError(f"Word reported success but {target} is missing")
        return str(target)

    def raw(self, code: str) -> str:
        """Evaluate a Python expression, or run a block that assigns `result`.

        In scope: `app` for the application, `doc` for the open document, `win32`.
        """
        import win32com.client

        scope = {"app": self._application(), "doc": self._doc, "win32": win32com.client}
        try:
            return repr(eval(code, scope))
        except SyntaxError:
            exec(code, scope)
            return repr(scope.get("result"))

"""LibreOffice engine, reached through the UNO bridge.

Used on Linux and anywhere else LibreOffice is the only word processor available.
Its pagination is LibreOffice's own and differs from Word by a few pages on a long
document, which `fidelity` declares.
"""

from __future__ import annotations

import importlib.util
import shutil
import subprocess
import time
from pathlib import Path

from ..engine import DocumentEngine, EngineError, NoDocumentError
from ..models import EngineStatus, Heading, ParagraphFormat

POINT_IN_HUNDREDTH_MM = 35.28


def _properties(**values):
    from com.sun.star.beans import PropertyValue

    return tuple(PropertyValue(Name=k, Value=v) for k, v in values.items())


class UnoEngine(DocumentEngine):
    """Drives a headless LibreOffice instance, starting one if none is listening."""

    name = "libreoffice"
    fidelity = "libreoffice"

    def __init__(self, port: int = 2002, startup_timeout: float = 30.0) -> None:
        """Choose the socket LibreOffice listens on and how long to wait for it."""
        self._port = port
        self._startup_timeout = startup_timeout
        self._desktop = None
        self._doc = None
        self._path: str | None = None

    def _binary(self) -> str:
        for candidate in ("soffice", "libreoffice"):
            found = shutil.which(candidate)
            if found:
                return found
        raise EngineError("LibreOffice is not installed or not in PATH")

    def _resolver(self):
        import uno

        context = uno.getComponentContext()
        return context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", context
        )

    def _desktop_or_start(self):
        if self._desktop is not None:
            return self._desktop
        resolver = self._resolver()
        url = f"uno:socket,host=127.0.0.1,port={self._port};urp;StarOffice.ComponentContext"
        try:
            remote = resolver.resolve(url)
        except Exception:
            remote = self._start_soffice(resolver, url)
        self._desktop = remote.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", remote
        )
        return self._desktop

    def _start_soffice(self, resolver, url: str):
        subprocess.Popen(
            [
                self._binary(),
                "--headless",
                "--norestore",
                "--invisible",
                f"--accept=socket,host=127.0.0.1,port={self._port};urp;",
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        deadline = time.monotonic() + self._startup_timeout
        while time.monotonic() < deadline:
            try:
                return resolver.resolve(url)
            except Exception:
                time.sleep(0.5)
        raise EngineError("LibreOffice did not accept a connection in time")

    def _document(self):
        if self._doc is None:
            raise NoDocumentError("no document is open")
        return self._doc

    def _view_cursor(self):
        return self._document().getCurrentController().getViewCursor()

    def _paragraphs(self):
        enumeration = self._document().Text.createEnumeration()
        while enumeration.hasMoreElements():
            element = enumeration.nextElement()
            if element.supportsService("com.sun.star.text.Paragraph"):
                yield element

    def status(self) -> EngineStatus:
        try:
            binary = self._binary()
        except EngineError as exc:
            return EngineStatus(self.name, self.fidelity, False, detail=str(exc))
        if importlib.util.find_spec("uno") is None:
            return EngineStatus(
                self.name,
                self.fidelity,
                False,
                detail="the python3-uno bridge is missing",
            )
        return EngineStatus(
            self.name, self.fidelity, True, version=binary, document=self._path
        )

    def open(self, path: str) -> str:
        import uno

        resolved = Path(path).expanduser().resolve()
        if not resolved.is_file():
            raise FileNotFoundError(resolved)
        desktop = self._desktop_or_start()
        self.close()
        self._doc = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(str(resolved)), "_blank", 0, _properties(Hidden=True)
        )
        self._path = str(resolved)
        return self._path

    def close(self, save: bool = False) -> None:
        if self._doc is None:
            return
        if save:
            self._doc.store()
        self._doc.close(False)
        self._doc = None
        self._path = None

    def save(self) -> None:
        self._document().store()

    def page_count(self) -> int:
        return int(self._document().getCurrentController().PageCount)

    def find_pages(self, text: str, limit: int = 20) -> list[int]:
        document = self._document()
        cursor = self._view_cursor()
        descriptor = document.createSearchDescriptor()
        descriptor.SearchString = text
        descriptor.SearchCaseSensitive = True
        pages: list[int] = []
        match = document.findFirst(descriptor)
        while match is not None and len(pages) < limit:
            cursor.gotoRange(match, False)
            pages.append(int(cursor.getPage()))
            match = document.findNext(match.End, descriptor)
        return pages

    def headings(self, max_level: int = 3) -> list[Heading]:
        cursor = self._view_cursor()
        found: list[Heading] = []
        for paragraph in self._paragraphs():
            level = int(getattr(paragraph, "OutlineLevel", 0) or 0)
            if level and level <= max_level:
                cursor.gotoRange(paragraph.Start, False)
                found.append(
                    Heading(level, int(cursor.getPage()), " ".join(paragraph.getString().split()))
                )
        return found

    def page_text(self, page: int) -> str:
        document = self._document()
        cursor = self._view_cursor()
        cursor.jumpToPage(int(page))
        start = cursor.getStart()
        cursor.jumpToEndOfPage()
        span = document.Text.createTextCursorByRange(start)
        span.gotoRange(cursor.getEnd(), True)
        return span.getString()

    def update_fields(self, toc: bool = True) -> list[str]:
        document = self._document()
        steps: list[str] = []
        try:
            document.getTextFields().refresh()
            steps.append("fields refreshed")
        except Exception as exc:
            steps.append(f"fields failed: {exc}")
        if toc:
            try:
                indexes = document.getDocumentIndexes()
                for position in range(indexes.getCount()):
                    indexes.getByIndex(position).update()
                steps.append(f"indexes updated: {indexes.getCount()}")
            except Exception as exc:
                steps.append(f"indexes failed: {exc}")
        return steps

    def format_paragraphs(self, contains: str, fmt: ParagraphFormat) -> int:
        if fmt.is_empty():
            raise ValueError("no property to set")
        changed = 0
        for paragraph in self._paragraphs():
            if contains not in paragraph.getString():
                continue
            if fmt.keep_with_next is not None:
                paragraph.ParaKeepTogether = fmt.keep_with_next
            if fmt.keep_together is not None:
                paragraph.ParaSplit = not fmt.keep_together
            if fmt.widow_control is not None:
                lines = 2 if fmt.widow_control else 0
                paragraph.ParaWidows = lines
                paragraph.ParaOrphans = lines
            if fmt.space_before is not None:
                paragraph.ParaTopMargin = int(fmt.space_before * POINT_IN_HUNDREDTH_MM)
            if fmt.space_after is not None:
                paragraph.ParaBottomMargin = int(fmt.space_after * POINT_IN_HUNDREDTH_MM)
            changed += 1
        return changed

    def export_pdf(self, pdf_path: str) -> str:
        import uno

        target = Path(pdf_path).expanduser().resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        self._document().storeToURL(
            uno.systemPathToFileUrl(str(target)), _properties(FilterName="writer_pdf_Export")
        )
        if not target.is_file():
            raise EngineError(f"LibreOffice reported success but {target} is missing")
        return str(target)

    def raw(self, code: str) -> str:
        """Evaluate a Python expression, or run a block that assigns `result`.

        In scope: `doc`, `desktop`, `uno`, and `vc` for the view cursor.
        """
        import uno

        scope = {
            "uno": uno,
            "doc": self._doc,
            "desktop": self._desktop_or_start(),
            "vc": self._view_cursor() if self._doc is not None else None,
        }
        try:
            return repr(eval(code, scope))
        except SyntaxError:
            exec(code, scope)
            return repr(scope.get("result"))

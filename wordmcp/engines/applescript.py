"""Microsoft Word on macOS, reached through Apple Events with osascript."""

from __future__ import annotations

import subprocess
from pathlib import Path

from ..engine import DocumentEngine, EngineError
from ..models import EngineStatus, Heading, ParagraphFormat

WORD_APP = Path("/Applications/Microsoft Word.app")


def _quote(text: str) -> str:
    """Escape a string for embedding in AppleScript double quotes."""
    return text.replace("\\", "\\\\").replace('"', '\\"')


class AppleScriptEngine(DocumentEngine):
    """Drives Word for Mac; page numbers are Word's own."""

    name = "word-mac"
    fidelity = "word"

    def __init__(self, timeout: float = 300.0) -> None:
        """Set how long a single AppleScript call may take, in seconds."""
        self._timeout = timeout
        self._path: str | None = None

    def _run(self, script: str, *arguments: str) -> str:
        completed = subprocess.run(
            ["osascript", "-", *arguments],
            input=script,
            capture_output=True,
            text=True,
            timeout=self._timeout,
        )
        if completed.returncode != 0:
            message = completed.stderr.strip() or completed.stdout.strip() or "osascript failed"
            raise EngineError(message.replace("\n", " "))
        return completed.stdout.strip()

    def _first_working(self, scripts: list[str]) -> str:
        failures = []
        for script in scripts:
            try:
                return self._run(script)
            except EngineError as exc:
                failures.append(str(exc))
        raise EngineError("; ".join(failures))

    def status(self) -> EngineStatus:
        if not WORD_APP.exists():
            return EngineStatus(self.name, self.fidelity, False, detail="Word is not installed")
        try:
            version = self._run('tell application "Microsoft Word" to return version')
        except EngineError as exc:
            return EngineStatus(
                self.name,
                self.fidelity,
                False,
                detail=(
                    f"{exc}. Allow automation in System Settings, Privacy and Security, "
                    "Automation, for the application running this server."
                ),
            )
        return EngineStatus(self.name, self.fidelity, True, version=version, document=self._path)

    def open(self, path: str) -> str:
        resolved = Path(path).expanduser().resolve()
        if not resolved.is_file():
            raise FileNotFoundError(resolved)
        self._run(
            'on run argv\n'
            '  tell application "Microsoft Word" to open (POSIX file (item 1 of argv))\n'
            'end run',
            str(resolved),
        )
        self._path = str(resolved)
        return self._path

    def close(self, save: bool = False) -> None:
        verb = "yes" if save else "no"
        self._run(f'tell application "Microsoft Word" to close active document saving {verb}')
        self._path = None

    def save(self) -> None:
        self._run('tell application "Microsoft Word" to save active document')

    def page_count(self) -> int:
        value = self._first_working(
            [
                'tell application "Microsoft Word" to return '
                "(compute statistics active document statistic statistic pages)",
                'tell application "Microsoft Word"\n'
                "  set total to compute statistics active document statistic statistic pages\n"
                "  return total as text\n"
                "end tell",
            ]
        )
        return int(float(value))

    def find_pages(self, text: str, limit: int = 20) -> list[int]:
        script = (
            'tell application "Microsoft Word"\n'
            '  set collected to ""\n'
            "  set span to text object of active document\n"
            "  set finish to end of content of text object of active document\n"
            "  set guard to 0\n"
            "  repeat\n"
            "    set guard to guard + 1\n"
            f"    if guard > {int(limit)} then exit repeat\n"
            "    tell (find object of span)\n"
            "      clear formatting\n"
            f'      set content to "{_quote(text)}"\n'
            "      set match case to true\n"
            "      set forward to true\n"
            "      set wrap to find stop\n"
            "      set hit to execute find\n"
            "    end tell\n"
            "    if not hit then exit repeat\n"
            "    set page to (get information span information type page number)\n"
            "    set collected to collected & (page as text) & \",\"\n"
            "    set resume to end of content of span\n"
            "    if resume ≥ finish then exit repeat\n"
            "    set span to create range active document start resume end finish\n"
            "  end repeat\n"
            "  return collected\n"
            "end tell"
        )
        return [int(part) for part in self._run(script).split(",") if part.strip().isdigit()]

    def headings(self, max_level: int = 3) -> list[Heading]:
        script = (
            'tell application "Microsoft Word"\n'
            '  set collected to ""\n'
            "  set total to count of paragraphs of active document\n"
            "  repeat with index from 1 to total\n"
            "    set para to paragraph index of active document\n"
            "    set level to (outline level of paragraph format of para) as integer\n"
            f"    if level ≤ {int(max_level)} then\n"
            "      set target to text object of para\n"
            "      set page to (get information target information type page number)\n"
            "      set line to (level as text) & tab & (page as text) & tab\n"
            "      set collected to collected & line & (content of target) & linefeed\n"
            "    end if\n"
            "  end repeat\n"
            "  return collected\n"
            "end tell"
        )
        found: list[Heading] = []
        for line in self._run(script).splitlines():
            parts = line.split("\t")
            if len(parts) >= 3 and parts[1].strip().isdigit():
                found.append(Heading(int(parts[0]), int(parts[1]), " ".join(parts[2].split())))
        return found

    def page_text(self, page: int) -> str:
        script = (
            'tell application "Microsoft Word"\n'
            f"  set here to go to active document what go to page "
            f"which go to absolute index {int(page)}\n"
            f"  set after to go to active document what go to page "
            f"which go to absolute index {int(page) + 1}\n"
            "  set from to start of content of here\n"
            "  set upto to start of content of after\n"
            "  if upto ≤ from then set upto to end of content of text object of active document\n"
            "  return content of (create range active document start from end upto)\n"
            "end tell"
        )
        return self._run(script)

    def update_fields(self, toc: bool = True) -> list[str]:
        steps: list[str] = []
        try:
            self._run(
                'tell application "Microsoft Word" to update fields '
                "of text object of active document"
            )
            steps.append("fields refreshed")
        except EngineError as exc:
            steps.append(f"fields failed: {exc}")
        if toc:
            try:
                self._run(
                    'tell application "Microsoft Word"\n'
                    "  repeat with entry in tables of contents of active document\n"
                    "    update entry\n"
                    "  end repeat\n"
                    "end tell"
                )
                steps.append("tables of contents updated")
            except EngineError as exc:
                steps.append(f"tables of contents failed: {exc}")
        return steps

    def format_paragraphs(self, contains: str, fmt: ParagraphFormat) -> int:
        if fmt.is_empty():
            raise ValueError("no property to set")
        assignments = []
        if fmt.keep_with_next is not None:
            assignments.append(f"set keep with next of style to {str(fmt.keep_with_next).lower()}")
        if fmt.keep_together is not None:
            assignments.append(f"set keep together of style to {str(fmt.keep_together).lower()}")
        if fmt.widow_control is not None:
            assignments.append(f"set widow control of style to {str(fmt.widow_control).lower()}")
        if fmt.space_before is not None:
            assignments.append(f"set space before of style to {float(fmt.space_before)}")
        if fmt.space_after is not None:
            assignments.append(f"set space after of style to {float(fmt.space_after)}")
        script = (
            'tell application "Microsoft Word"\n'
            "  set changed to 0\n"
            "  set total to count of paragraphs of active document\n"
            "  repeat with index from 1 to total\n"
            "    set para to paragraph index of active document\n"
            f'    if (content of text object of para) contains "{_quote(contains)}" then\n'
            "      set style to paragraph format of para\n"
            "      " + "\n      ".join(assignments) + "\n"
            "      set changed to changed + 1\n"
            "    end if\n"
            "  end repeat\n"
            "  return changed as text\n"
            "end tell"
        )
        return int(self._run(script))

    def export_pdf(self, pdf_path: str) -> str:
        target = Path(pdf_path).expanduser().resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        self._run(
            "on run argv\n"
            '  tell application "Microsoft Word"\n'
            "    set destination to (POSIX file (item 1 of argv)) as text\n"
            "    save as active document file name destination file format format PDF\n"
            "  end tell\n"
            "end run",
            str(target),
        )
        if not target.is_file():
            raise EngineError(f"Word reported success but {target} is missing")
        return str(target)

    def raw(self, code: str) -> str:
        """Run AppleScript exactly as osascript would."""
        return self._run(code)

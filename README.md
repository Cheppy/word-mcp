# word-mcp

An MCP server that measures and fixes the page layout of `.docx` documents using a real word
processor rather than a converter. Converters paginate differently from Word, and over a few
hundred pages the drift adds up, so the numbers here come from the application itself.

| Platform | Driven through | Whose pagination |
|---|---|---|
| macOS | Microsoft Word over Apple Events | Word |
| Windows | Microsoft Word over COM | Word |
| Linux | LibreOffice over UNO | LibreOffice |

The engine follows the platform; `WORD_MCP_BACKEND=mac|win|uno` overrides it. Every answer
carries a `fidelity` field saying whose numbers these are.

## Example

Asking a document where its parts begin and how long they are:

```
>>> word_status()
{'engine': 'libreoffice', 'fidelity': 'libreoffice', 'available': True,
 'version': '/usr/bin/soffice', 'document': None, 'detail': None}

>>> word_open('/tmp/report.docx')
'/tmp/report.docx'

>>> word_outline()
{
  "total_pages": 13,
  "fidelity": "libreoffice",
  "sections": [
    { "name": "Introduction", "page":  1, "pages": 2 },
    { "name": "Background",   "page":  3, "pages": 3 },
    { "name": "Method",       "page":  6, "pages": 3 },
    { "name": "Results",      "page":  9, "pages": 2 },
    { "name": "Conclusions",  "page": 11, "pages": 2 },
    { "name": "Appendix A",   "page": 13, "pages": 1 }
  ]
}

>>> word_find_page('Complexity')
{'text': 'Complexity', 'pages': [7, 7, 7, 7, 8, 8]}
```

Keeping a caption attached to the figure above it, so a page break never separates them:

```
>>> word_format_paragraphs(contains='Figure 1.', keep_with_next=True, widow_control=True)
1
>>> word_save()
'saved'
```

Sections can also be defined by text instead of by heading level, which is useful when a
document has no outline levels set:

```
>>> word_outline(markers=['Introduction', 'Method', 'Conclusions'])
```

The last occurrence of each marker is taken, so mentions inside a table of contents are skipped.

## Layout

```
wordmcp/
├── models.py          immutable result and parameter types
├── engine.py          DocumentEngine: the contract an engine fulfils
├── engines/
│   ├── __init__.py    engine selection; dependencies are passed in, not read from the module
│   ├── applescript.py Word on macOS
│   ├── com.py         Word on Windows
│   └── uno.py         LibreOffice
├── service.py         document operations, expressed once
└── server.py          MCP tools: validate, delegate, return a model
```

Dependencies are injected through constructors: an engine goes into `DocumentService`, the
service goes into `create_server`. There is no global state, so the logic runs against a stub:

```python
server = create_server(DocumentService(StubEngine()), allow_raw=False)
```

## Tools

| Tool | What it does |
|---|---|
| `word_status` | which engine is active, whether it answers, what is open |
| `word_open` / `word_close` / `word_save` | open, close, save |
| `word_page_count` | number of pages |
| `word_find_page` | pages on which a string occurs |
| `word_headings` | headings with outline level and page |
| `word_outline` | start page and length of every top-level part |
| `word_page_text` | text placed on a given page |
| `word_update_fields` | refresh fields and tables of contents |
| `word_format_paragraphs` | keep with next, keep together, widow control, spacing |
| `word_export_pdf` | export to PDF with the application's own engine |
| `word_selftest` | run every operation and report what this machine supports |
| `word_run_raw` | engine-native code; registered only when `WORD_MCP_ALLOW_RAW=1` |

## Installing

```bash
git clone https://github.com/Cheppy/word-mcp.git
cd word-mcp
uv sync
```

**macOS.** Needs Microsoft Word and automation permission: System Settings, Privacy and
Security, Automation, the application running this server, Microsoft Word. The first call
raises the system prompt.

**Windows.** Needs Microsoft Word: `uv sync --extra windows`, which adds `pywin32`.

**Linux.** Needs LibreOffice and the system Python-UNO package:
`sudo apt install libreoffice python3-uno`. The server starts `soffice` headless by itself;
`WORD_MCP_UNO_PORT` sets the socket. Run it with the interpreter that can see `uno`, which is
usually the system `python3`.

## Connecting

```json
{
  "mcpServers": {
    "word": {
      "command": "uv",
      "args": ["run", "--directory", "/path/to/word-mcp", "word-mcp"],
      "env": { "WORD_MCP_ALLOW_RAW": "1" }
    }
  }
}
```

On Linux the system interpreter is usually required:

```json
{
  "mcpServers": {
    "word": {
      "command": "python3",
      "args": ["-m", "wordmcp.server"],
      "env": { "PYTHONPATH": "/path/to/word-mcp" }
    }
  }
}
```

`WORD_MCP_ALLOW_RAW=1` registers `word_run_raw`, which runs code in the engine's own language:
AppleScript on macOS, a Python expression over the live application objects on Windows and
Linux. It exists so a failing call can be repaired without editing the server, but it is
arbitrary code execution; drop the variable once things settle.

Restart the client after changing its configuration.

## Development

```bash
uv run pytest        # service logic against a stub engine, no word processor needed
uv run ruff check .
```

## State of testing

The LibreOffice engine has been exercised on a real 161-page document: page count, text search,
heading map, section lengths, page text, paragraph formatting, field updates and the escape
hatch.

The Word engines for macOS and Windows follow documented interfaces but have not been run on a
live machine. Start with `word_selftest`. The likely places to adjust on macOS are
`compute statistics` and `get information ... information type page number`; the corresponding
COM calls on Windows are more stable.

## Licence

GNU General Public License v2.0; see `LICENSE`.

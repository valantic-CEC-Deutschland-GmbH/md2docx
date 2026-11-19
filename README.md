# Markdown → Word (valantic CI)

Converts Markdown files into Word documents (`.docx`) with valantic corporate styling, and optionally into PDF.

Write your report, concept or documentation in Markdown, run one command, and get a Word document with the valantic
logo, fonts, colours, headings, tables, page numbers and table of contents.

## Quick Start

```bash
npm install                          # once
node md2docx.js report.md            # → report.docx
node md2docx.js report.md --pdf      # → report.docx + report.pdf
```

The `.docx` (and `.pdf`) is created next to the Markdown file with the same name. Existing files are overwritten.

If the document contains `[TOC]`, open the `.docx` in Word and press **Ctrl+A → F9** (macOS: **Cmd+A → F9** or
**fn+F9**) to fill the table of contents. The PDF does not need this step, the table of contents is filled
automatically.

## Requirements

| What | Why |
|------|-----|
| Node.js 18 or newer | runs the converter |
| Segoe UI, Segoe UI Semibold, Consolas | fonts of the valantic CI (see [Fonts](#fonts)) |
| Microsoft Word (macOS or Windows) | only for `--pdf`, recommended |
| LibreOffice | only for `--pdf` without Word (fallback) |

## Usage

```bash
node md2docx.js input.md                                # → input.docx
node md2docx.js input.md output.docx                    # custom output path
node md2docx.js input.md --pdf                          # additionally input.pdf
node md2docx.js input.md output.docx --pdf              # → output.docx + output.pdf
node md2docx.js input.md output.docx --logo my-logo.png # custom logo in the header
node md2docx.js --help
```

## PDF Export (`--pdf`)

The PDF is created from the finished Word document, so it looks exactly like the `.docx`.

| Engine | Used when | Result |
|--------|-----------|--------|
| **Microsoft Word** | Word is installed (macOS or Windows) | best result: real fonts, table of contents is filled automatically |
| **LibreOffice** | no Word, but LibreOffice is installed | fallback: fonts are replaced if Segoe UI is missing, table of contents may stay empty |

If neither is found, the `.docx` is still created and the converter reports an error for the PDF.

Notes:

- **macOS:** Word runs in the background and is closed again afterwards if it was not open before. The first time,
  macOS may ask whether the terminal is allowed to control Microsoft Word, and Word may ask for access to the output
  folder. Allow both once.
- **macOS:** If Word shows a dialog (e.g. a question about updating fields), the conversion waits. After about 1 minute it
  stops with the error "Word konnte das Dokument nicht öffnen". Close the dialog in Word and run the command again.
- **Windows:** Word is controlled via PowerShell. No extra installation is needed.

## Supported Markdown

| Markdown | Word |
|----------|------|
| `# Title` | Heading 1 — Segoe UI Semibold 22pt, page break |
| `## Heading` | Heading 2 — Segoe UI Semibold 16pt |
| `### Heading` | Heading 3 — Segoe UI Semibold 14pt |
| `#### Heading` | Heading 4 — Segoe UI Semibold 11pt |
| `**bold**` | Bold |
| `*italic*` / `_italic_` | Italic (`***text***` = bold + italic) |
| `` `code` `` | Consolas inline |
| `[text](https://…)` | Clickable hyperlink (blue, underlined). Relative links (e.g. `other.md`) are output as plain text |
| `\*` `\_` `\|` `` \` `` | Escaped character, output literally (e.g. `\|` inside a table cell) |
| ` ```code``` ` | Consolas code block |
| `- item` | Bullet list (`-`, `*` or `+`) |
| `1. item` | Numbered list |
| Indented list items | Nested lists, up to 3 levels; bullets and numbers can be mixed |
| Indented line after a list item | Continues the same list item (no new paragraph, numbering continues) |
| `> quote` | Blockquote with blue border |
| `| table |` | Formatted table with alternating rows. Column widths follow the content (narrow columns stay narrow); 10 pt font from 5 columns, 9 pt from 6, 8 pt from 10 columns. Numbers stay together with their unit or bracket (`5 s`, `120 ms`, `18.6 %`, `4,836 (4,831)`) and are not split across lines |
| `---` | Horizontal line |
| `[TOC]` | Table of contents |

Not supported: images, footnotes, merged table cells, task lists (`- [ ]`), HTML tags, more than 3 list levels.

## Tips for Good Results

- **Title and table of contents:** start with `# Title`, then put `[TOC]` where the table of contents should appear
  (e.g. after an introduction or status table).
- **Wide tables:** the page fits about 8–10 columns. Keep headers short, put units in the header instead of every
  cell (`p95 (ms)` instead of `512 ms` in each row), and explain abbreviations in a legend below or above the table.
- **Pipe in a table cell:** write `\|`, otherwise the cell is split.
- **Links:** only `http(s)://` and `mailto:` links are clickable. Links to other Markdown files become plain text,
  because they do not work in a Word file anyway.
- **Long list items:** an indented line directly below a list item continues that item; an empty line followed by a
  new item keeps the list going.
- **Line breaks in paragraphs:** consecutive lines are joined into one paragraph, as in normal Markdown. Leave an
  empty line to start a new paragraph.

## Troubleshooting

| Problem | Solution |
|---------|----------|
| Table of contents is empty or shows "No table of contents entries found" | In Word: Ctrl+A → F9 (macOS: Cmd+A → F9). Or create the PDF with `--pdf`, which fills it automatically |
| Wrong font on macOS or Linux | Segoe UI is not installed, see [Fonts](#fonts) |
| Table is too wide / numbers wrap | Fewer columns, shorter headers, units in the header (see tips above) |
| `--pdf` stops with "Word konnte das Dokument nicht öffnen" | A dialog is open in Word. Close it and run the command again |
| `--pdf`: "weder Microsoft Word noch LibreOffice gefunden" | Install Word or LibreOffice, or export the PDF manually from the `.docx` |
| `*` or `_` disappears from the text | It was read as italic. Write `\*` or `\_` |

## Logo

The valantic logo (`valantic-logo.png`) from the project directory is used automatically in the document header. Use
`--logo path/to/logo.png` to use a different one.

## Fonts

The document uses the following fonts:

| Font | Usage |
|------|-------|
| Segoe UI | Body text |
| Segoe UI Semibold | Headings |
| Consolas | Code blocks and inline code |

The fonts must be installed on the system where the document is opened. Segoe UI is pre-installed on Windows. On macOS
Microsoft Word usually brings them along; otherwise, and on Linux, the fonts need to be installed separately.
Alternatively, open the document in Word Online — the fonts are available there without local installation.

## Example

A complete example Markdown file is available in the `examples/` folder. You can use it to see the supported
formatting and try out the converter:

```bash
node md2docx.js examples/example.md
node md2docx.js examples/example.md --pdf
```

The files are created in the same directory with the same name (`examples/example.docx`, `examples/example.pdf`).

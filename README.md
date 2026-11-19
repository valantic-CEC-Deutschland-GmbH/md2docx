# Markdown → Word (valantic CI)

Converts Markdown files into Word documents (`.docx`) with valantic corporate styling.

## Features

- Segoe UI / Segoe UI Semibold typography
- Automatic table of contents via `[TOC]`
- Header with valantic logo
- Footer with page numbers
- Code blocks, tables, lists, blockquotes

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
| `| table |` | Formatted table with alternating rows. Column widths follow the content (narrow columns stay narrow); 10 pt font from 5 columns, 9 pt from 6, 8 pt from 10 columns |
| `---` | Horizontal line |
| `[TOC]` | Table of contents |

## Installation

```bash
npm install
```

## Usage

```bash
node md2docx.js input.md
```

The `.docx` file is automatically created in the same directory with the same name. Existing files are overwritten.

Optionally, you can specify a custom output path:

```bash
node md2docx.js input.md output.docx
```

## Logo

The valantic logo (`valantic-logo.png`) from the project directory is used automatically in the document header.

## Fonts

The document uses the following fonts:

| Font | Usage |
|------|-------|
| Segoe UI | Body text |
| Segoe UI Semibold | Headings |
| Consolas | Code blocks and inline code |

The fonts must be installed on the system where the document is opened. Segoe UI is pre-installed on Windows. On macOS or Linux the fonts need to be installed separately. Alternatively, open the document in Word Online — the fonts are available there without local installation.

## Example

A complete example Markdown file is available in the `examples/` folder. You can use it to see the supported formatting and try out the converter:

```bash
node md2docx.js examples/example.md
```

The `.docx` file is automatically created in the same directory with the same name (`examples/example.docx`).

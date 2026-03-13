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
| `` `code` `` | Consolas inline |
| ` ```code``` ` | Consolas code block |
| `- item` | Bullet list |
| `1. item` | Numbered list |
| `> quote` | Blockquote with blue border |
| `| table |` | Formatted table with alternating rows |
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

The `.docx` file is saved next to the input file. Existing files are overwritten.

## Logo

The default logo `valantic-logo.png` from the project directory is used automatically. To use a different logo:

```bash
node md2docx.js input.md output.docx --logo path/to/logo.png
```

## Fonts

The document uses the following fonts:

| Font | Usage |
|------|-------|
| Segoe UI | Body text |
| Segoe UI Semibold | Headings |
| Consolas | Code blocks and inline code |

The fonts must be installed on the system where the document is opened. Segoe UI is pre-installed on Windows. On macOS or Linux the fonts need to be installed separately. Alternatively, open the document in Word Online — the fonts are available there without local installation.

## Example

A complete example is available in the `examples/` folder:

```bash
node md2docx.js examples/example.md examples/example.docx
```

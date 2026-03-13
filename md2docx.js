#!/usr/bin/env node
/**
 * md2docx.js — Markdown → Word Converter (valantic CI)
 *
 * Verwendung:
 *   node md2docx.js eingabe.md
 *   node md2docx.js eingabe.md ausgabe.docx
 *   node md2docx.js eingabe.md ausgabe.docx --logo pfad/zum/logo.png
 *
 * Logo:    Standard ist "valantic-logo.png" im gleichen Verzeichnis wie das Script.
 * Styles:  Alle valantic CI-Styles sind direkt eingebaut (berschrift1–4, Quellcode, Infobox etc.)
 *
 * Unterstützte Markdown-Elemente:
 *   # Titel         → Roter Titel-Style (Segoe UI Semibold 30pt)
 *   ## H1           → Heading 1 (Segoe UI Semibold 22pt)
 *   ### H2          → Heading 2 (Segoe UI Semibold 16pt)
 *   #### H3         → Heading 3 (Segoe UI Semibold 14pt)
 *   ##### H4        → Heading 4 (Segoe UI Semibold 11pt)
 *   **fett**        → Bold
 *   `inline code`   → Consolas 11pt
 *   ```codeblock``` → Consolas eingerückt
 *   - / * Liste     → Bullet
 *   1. Liste        → Nummeriert
 *   > Blockquote    → Blauer linker Rand
 *   | Tabelle |     → CI-Tabelle
 *   ---             → Trennlinie
 *   [TOC]           → Inhaltsverzeichnis (Headings 1-4)
 */

'use strict';

const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    Header, Footer, HeadingLevel, BorderStyle, WidthType, ShadingType,
    AlignmentType, LevelFormat, VerticalAlign, ImageRun, TableOfContents,
    PageNumber
} = require('docx');
const fs   = require('fs');
const path = require('path');

// ─── CLI args ─────────────────────────────────────────────────────────────────
const args = process.argv.slice(2);
if (args.length === 0 || args[0] === '--help' || args[0] === '-h') {
    console.log(`
  md2docx — Markdown → Word (valantic CI)
  ────────────────────────────────────────
  Verwendung:
    node md2docx.js eingabe.md
    node md2docx.js eingabe.md ausgabe.docx
    node md2docx.js eingabe.md ausgabe.docx --logo pfad/logo.png

  Logo-Standard: valantic-logo.png im gleichen Verzeichnis wie das Script.
  `);
    process.exit(0);
}

const inputPath  = path.resolve(args[0]);
const outputPath = (args[1] && !args[1].startsWith('--'))
    ? path.resolve(args[1])
    : inputPath.replace(/\.md$/i, '.docx');

const logoFlag = args.indexOf('--logo');
const logoPath = logoFlag >= 0
    ? path.resolve(args[logoFlag + 1])
    : path.join(__dirname, 'valantic-logo.png');

if (!fs.existsSync(inputPath)) {
    console.error(`Fehler: Datei nicht gefunden: ${inputPath}`);
    process.exit(1);
}

// ─── CI-Konstanten ────────────────────────────────────────────────────────────
const PAGE_W      = 11906;
const PAGE_H      = 16838;
const MARGIN_TOP  = 1418;
const MARGIN_SIDE = 1418;
const MARGIN_BOT  = 1134;
const HEADER_H    = 709;
const CONTENT_W   = PAGE_W - MARGIN_SIDE * 2; // 9070 DXA

const ALT_BG    = "F2F2F2";
const HDR_LINE  = "677A82";
const QUOTE_CLR = "2E4A6B";
const QUOTE_BDR = "2E75B6";

const SZ = { title: 60, h1: 44, h2: 32, h3: 28, h4: 22, body: 22 };

// ─── Logo ─────────────────────────────────────────────────────────────────────
let logoData = null;
let hasLogo  = false;
if (fs.existsSync(logoPath)) {
    logoData = fs.readFileSync(logoPath);
    hasLogo  = true;
} else {
    console.warn(`Warnung: Logo nicht gefunden: ${logoPath}`);
}

// ─── Inline-Parser ────────────────────────────────────────────────────────────
function parseInline(text, base = {}) {
    const runs = [];
    const rx   = /(\*\*(.+?)\*\*|`([^`]+?)`)/g;
    let last   = 0, m;
    while ((m = rx.exec(text)) !== null) {
        if (m.index > last) runs.push({ ...base, text: text.slice(last, m.index) });
        if (m[2] !== undefined) runs.push({ ...base, bold: true,       text: m[2] });
        else                    runs.push({ ...base, font: "Consolas", text: m[3] });
        last = m.index + m[0].length;
    }
    if (last < text.length) runs.push({ ...base, text: text.slice(last) });
    return runs.length ? runs : [{ ...base, text }];
}

function toRuns(text, base = {}) {
    return parseInline(text, base).map(r => new TextRun({
        font:  r.font || (r.bold ? "Segoe UI Semibold" : "Segoe UI"),
        size:  r.size  || SZ.body,
        bold:  !!r.bold,
        color: r.color || undefined,
        text:  r.text
    }));
}

// ─── Absatz-Helfer ────────────────────────────────────────────────────────────
// Style-IDs aus der valantic vorlage.docx (deutsche Umlaute → encoded)
const HEADING_STYLE = {
    1: 'berschrift1',
    2: 'berschrift2',
    3: 'berschrift3',
    4: 'berschrift4',
};

const HEADING_FONT = "Segoe UI Semibold";
const HEADING_SIZE = { 1: 44, 2: 32, 3: 28, 4: SZ.body };

function hPara(text, level) {
    return new Paragraph({
        style:    HEADING_STYLE[level],
        children: [new TextRun({ text, font: HEADING_FONT, size: HEADING_SIZE[level], bold: level === 1 })]
    });
}

function bodyPara(text, opts = {}) {
    return new Paragraph({
        spacing:   { before: opts.before || 60, after: opts.after || 60 },
        numbering:  opts.numbering,
        children:   toRuns(text, { size: SZ.body })
    });
}

function quotePara(text) {
    return new Paragraph({
        spacing: { before: 80, after: 80 },
        border:  { left: { style: BorderStyle.SINGLE, size: 8, color: QUOTE_BDR, space: 10 } },
        indent:  { left: 300 },
        children: toRuns(text, { size: SZ.body, color: QUOTE_CLR })
    });
}

function codePara(text) {
    return new Paragraph({
        spacing: { before: 30, after: 30 },
        indent:  { left: 360 },
        children: [new TextRun({ font: "Consolas", size: SZ.body, text, color: "444444" })]
    });
}

function hrPara() {
    return new Paragraph({
        spacing: { before: 120, after: 120 },
        border:  { bottom: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 1 } },
        children: []
    });
}

function spacer(sz = 120) {
    return new Paragraph({ spacing: { before: sz, after: 0 }, children: [] });
}

// ─── Tabellen ─────────────────────────────────────────────────────────────────
function parseTableRow(line) {
    return line.replace(/^\||\|$/g, '').split('|').map(c => c.trim());
}

function isAlignRow(line) {
    return /^\|[\s|:-]+\|$/.test(line.trim());
}

function buildTable(headerCells, dataRows) {
    const n    = headerCells.length;
    const w    = Math.floor(CONTENT_W / n);
    const colW = Array(n).fill(w);
    colW[n - 1] = CONTENT_W - w * (n - 1);

    const none     = { style: BorderStyle.NONE,   size: 0, color: "FFFFFF" };
    const solidBdr = { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" };

    function makeCell(text, isHdr, colIdx, altRow) {
        const runs = parseInline(text, { size: SZ.body }).map(r => new TextRun({
            font: r.font || "Segoe UI",
            size: SZ.body,
            bold: isHdr ? true : !!r.bold,
            text: r.text
        }));
        const borders = isHdr
            ? { top: none, left: none, right: none, bottom: solidBdr }
            : { top: none, bottom: none, left: colIdx === 0 ? none : solidBdr, right: none };
        const shading = (altRow && !isHdr)
            ? { fill: ALT_BG, type: ShadingType.CLEAR }
            : { fill: "FFFFFF", type: ShadingType.CLEAR };
        return new TableCell({
            width:         { size: colW[colIdx], type: WidthType.DXA },
            borders, shading,
            margins:       { top: 60, bottom: 60, left: 100, right: 100 },
            verticalAlign: VerticalAlign.CENTER,
            children:      [new Paragraph({ children: runs })]
        });
    }

    return new Table({
        width:        { size: CONTENT_W, type: WidthType.DXA },
        columnWidths: colW,
        rows: [
            new TableRow({ tableHeader: true, children: headerCells.map((h, i) => makeCell(h, true, i, false)) }),
            ...dataRows.map((row, ri) => new TableRow({
                children: Array(n).fill('').map((_, i) => makeCell(row[i] || '', false, i, ri % 2 === 1))
            }))
        ]
    });
}

// ─── Header ───────────────────────────────────────────────────────────────────
function buildHeader(titleText) {
    const none    = { style: BorderStyle.NONE,   size: 0, color: "FFFFFF" };
    const botLine = { style: BorderStyle.SINGLE, size: 4, color: HDR_LINE, space: 0 };

    const titleCell = new TableCell({
        width:         { size: 7230, type: WidthType.DXA },
        borders:       { top: none, left: none, right: none, bottom: none },
        margins:       { top: 0, bottom: 0, left: 0, right: 0 },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
            children: [new TextRun({ text: titleText, font: "Segoe UI", size: 18, color: "555555" })]
        })]
    });

    const logoCell = new TableCell({
        width:         { size: 1840, type: WidthType.DXA },
        borders:       { top: none, left: none, right: none, bottom: none },
        margins:       { top: 0, bottom: 0, left: 0, right: 0 },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children:  hasLogo ? [new ImageRun({
                data:           logoData,
                type:           "png",
                transformation: { width: 107, height: 23 }
            })] : []
        })]
    });

    return new Header({
        children: [
            new Table({
                width:        { size: CONTENT_W, type: WidthType.DXA },
                columnWidths: [7230, 1840],
                borders:      { top: none, left: none, right: none, insideH: none, insideV: none, bottom: botLine },
                rows: [new TableRow({
                    height:   { value: 573, rule: "exact" },
                    children: [titleCell, logoCell]
                })]
            }),
            new Paragraph({ children: [] })
        ]
    });
}

// ─── Footer ───────────────────────────────────────────────────────────────────
function buildFooter() {
    const none    = { style: BorderStyle.NONE,   size: 0, color: "FFFFFF" };
    const topLine = { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA", space: 0 };

    function emptyCell(w) {
        return new TableCell({
            width:   { size: w, type: WidthType.DXA },
            borders: { top: none, left: none, right: none, bottom: none },
            margins: { top: 0, bottom: 0, left: 0, right: 0 },
            children: [new Paragraph({ children: [] })]
        });
    }

    const pageCell = new TableCell({
        width:         { size: 1415, type: WidthType.DXA },
        borders:       { top: none, left: none, right: none, bottom: none },
        margins:       { top: 0, bottom: 0, left: 0, right: 0 },
        verticalAlign: VerticalAlign.BOTTOM,
        children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [
                new TextRun({ text: "Page ",   font: "Segoe UI", size: 18, color: "555555" }),
                new TextRun({ children: [PageNumber.CURRENT],     font: "Segoe UI", size: 18, color: "555555" }),
                new TextRun({ text: " of ",    font: "Segoe UI", size: 18, color: "555555" }),
                new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Segoe UI", size: 18, color: "555555" }),
            ]
        })]
    });

    return new Footer({
        children: [
            new Table({
                width:        { size: CONTENT_W, type: WidthType.DXA },
                columnWidths: [1418, 2977, 3260, 1415],
                borders:      { top: topLine, bottom: none, left: none, right: none, insideH: none, insideV: none },
                rows: [new TableRow({
                    height:   { value: 397, rule: "exact" },
                    children: [emptyCell(1418), emptyCell(2977), emptyCell(3260), pageCell]
                })]
            }),
            new Paragraph({ children: [] })
        ]
    });
}

// ─── TOC ──────────────────────────────────────────────────────────────────────
function buildTOC() {
    return new TableOfContents("Inhaltsverzeichnis", {
        hyperlink:          true,
        headingStyleRange:  "1-4",
        // valantic Vorlage nutzt berschrift1–4 statt Heading1–4
        stylesWithLevels: [
            { styleId: "berschrift1", level: 1 },
            { styleId: "berschrift2", level: 2 },
            { styleId: "berschrift3", level: 3 },
            { styleId: "berschrift4", level: 4 },
        ]
    });
}

// ─── Haupt-Parser ─────────────────────────────────────────────────────────────
function parseMarkdown(md) {
    const lines            = md.split('\n');
    const elements         = [];
    const numberingConfigs = [];
    let   bulletRef        = 0;
    let   numberRef        = 0;
    let   docTitle         = '';
    let   hasTOC           = false;

    function newBulletRef() {
        const ref = `bullet-${bulletRef++}`;
        numberingConfigs.push({ reference: ref, levels: [{
                level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
                style: { paragraph: { indent: { left: 360, hanging: 360 } } }
            }]});
        return ref;
    }

    function newNumberRef() {
        const ref = `number-${numberRef++}`;
        numberingConfigs.push({ reference: ref, levels: [{
                level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
                style: { paragraph: { indent: { left: 360, hanging: 360 } } }
            }]});
        return ref;
    }

    let i = 0;
    while (i < lines.length) {
        const line    = lines[i];
        const trimmed = line.trim();

        if (trimmed === '') { i++; continue; }

        if (/^---+$/.test(trimmed)) { elements.push(hrPara()); i++; continue; }

        if (/^\[TOC\]$/i.test(trimmed) || /^\{toc\}$/i.test(trimmed)) {
            elements.push(buildTOC());
            elements.push(spacer(120));
            hasTOC = true;
            i++;
            continue;
        }

        const hMatch = trimmed.match(/^(#{1,4})\s+(.+)$/);
        if (hMatch) {
            const depth = hMatch[1].length;
            const text  = hMatch[2];
            if (!docTitle && depth === 1) docTitle = text;
            elements.push(hPara(text, depth));
            i++;
            continue;
        }

        if (trimmed.startsWith('```')) {
            i++;
            while (i < lines.length && !lines[i].trim().startsWith('```')) {
                elements.push(codePara(lines[i]));
                i++;
            }
            i++;
            continue;
        }

        if (trimmed.startsWith('>')) {
            const text = trimmed.replace(/^>\s?/, '').trim();
            elements.push(quotePara(text));
            i++;
            continue;
        }

        if (trimmed.startsWith('|')) {
            const headerCells = parseTableRow(trimmed);
            i++;
            if (i < lines.length && isAlignRow(lines[i])) i++;
            const dataRows = [];
            while (i < lines.length && lines[i].trim().startsWith('|')) {
                dataRows.push(parseTableRow(lines[i].trim()));
                i++;
            }
            elements.push(buildTable(headerCells, dataRows));
            elements.push(spacer(80));
            continue;
        }

        if (/^[-*]\s/.test(trimmed)) {
            const ref = newBulletRef();
            while (i < lines.length && /^[-*]\s/.test(lines[i].trim())) {
                elements.push(bodyPara(lines[i].trim().replace(/^[-*]\s/, ''), { numbering: { reference: ref, level: 0 }, before: 40, after: 40 }));
                i++;
            }
            continue;
        }

        if (/^\d+\.\s/.test(trimmed)) {
            const ref = newNumberRef();
            while (i < lines.length && /^\d+\.\s/.test(lines[i].trim())) {
                elements.push(bodyPara(lines[i].trim().replace(/^\d+\.\s/, ''), { numbering: { reference: ref, level: 0 }, before: 40, after: 40 }));
                i++;
            }
            continue;
        }

        // Normaler Absatz (mehrzeilig zusammenfassen)
        let paraText = trimmed;
        i++;
        while (
            i < lines.length &&
            lines[i].trim() !== '' &&
            !/^#{1,5}\s/.test(lines[i].trim()) &&
            !lines[i].trim().startsWith('|') &&
            !lines[i].trim().startsWith('>') &&
            !lines[i].trim().startsWith('```') &&
            !/^[-*]\s/.test(lines[i].trim()) &&
            !/^\d+\.\s/.test(lines[i].trim()) &&
            !/^---+$/.test(lines[i].trim()) &&
            !/^\[TOC\]$/i.test(lines[i].trim())
            ) {
            paraText += ' ' + lines[i].trim();
            i++;
        }
        elements.push(bodyPara(paraText, { before: 80, after: 80 }));
    }

    return { elements, numberingConfigs, docTitle, hasTOC };
}

// ─── Document bauen ───────────────────────────────────────────────────────────
function buildDocument(elements, numberingConfigs, docTitle) {
    return new Document({
        features:  { updateFields: true },
        numbering: { config: numberingConfigs },
        styles: {
            default: { document: { run: { font: "Segoe UI", size: SZ.body } } },
            paragraphStyles: [
                // Normal — Textfarbe aus valantic Vorlage
                { id: "Normal", name: "Normal",
                    run:       { font: "Segoe UI", size: SZ.body, color: "100C2A" },
                    paragraph: { spacing: { before: 60, after: 60 } } },

                // berschrift1 — sz=44, bold, spacing before=800, pageBreak vor Kapitel
                { id: "berschrift1", name: "heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
                    run:       { font: "Segoe UI Semibold", size: 44, bold: true, color: "100C2A" },
                    paragraph: { spacing: { before: 800, line: 240 }, keepNext: true, keepLines: true, pageBreakBefore: true, outlineLevel: 0 } },

                // berschrift2 — sz=32, spacing before=600
                { id: "berschrift2", name: "heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
                    run:       { font: "Segoe UI Semibold", size: 32, color: "100C2A" },
                    paragraph: { spacing: { before: 600, line: 240 }, keepNext: true, keepLines: true, outlineLevel: 1 } },

                // berschrift3 — sz=28, spacing before=600
                { id: "berschrift3", name: "heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
                    run:       { font: "Segoe UI Semibold", size: 28, color: "100C2A" },
                    paragraph: { spacing: { before: 600, line: 240 }, keepNext: true, keepLines: true, outlineLevel: 2 } },

                // berschrift4 — gleiche Größe wie body, spacing before=600
                { id: "berschrift4", name: "heading 4", basedOn: "Normal", next: "Normal", quickFormat: true,
                    run:       { font: "Segoe UI Semibold", size: SZ.body, color: "100C2A" },
                    paragraph: { spacing: { before: 600, line: 240 }, keepNext: true, keepLines: true, outlineLevel: 3 } },

                // Quellcode — Consolas, grauer Hintergrund F2F2F2
                { id: "Quellcode", name: "Quellcode", basedOn: "Normal", quickFormat: true,
                    run:       { font: "Consolas", size: SZ.body, color: "444444" },
                    paragraph: {
                        spacing: { before: 30, after: 30 },
                        indent:  { left: 284, right: 284 },
                        border:  {
                            top:    { style: BorderStyle.SINGLE, size: 8, color: "F2F2F2", space: 5 },
                            bottom: { style: BorderStyle.SINGLE, size: 8, color: "F2F2F2", space: 5 },
                            left:   { style: BorderStyle.SINGLE, size: 8, color: "F2F2F2", space: 5 },
                            right:  { style: BorderStyle.SINGLE, size: 8, color: "F2F2F2", space: 5 },
                        },
                        shading: { fill: "F2F2F2", type: ShadingType.CLEAR }
                    }
                },

                // Infobox — blauer Rahmen/Hintergrund C2D2F1
                { id: "Infobox", name: "Infobox", basedOn: "Normal", quickFormat: true,
                    run:       { font: "Segoe UI", size: SZ.body, color: "100C2A" },
                    paragraph: {
                        indent:  { left: 284, right: 284 },
                        border:  {
                            top:    { style: BorderStyle.SINGLE, size: 8, color: "C2D2F1", space: 5 },
                            bottom: { style: BorderStyle.SINGLE, size: 8, color: "C2D2F1", space: 5 },
                            left:   { style: BorderStyle.SINGLE, size: 8, color: "C2D2F1", space: 5 },
                            right:  { style: BorderStyle.SINGLE, size: 8, color: "C2D2F1", space: 5 },
                        },
                        shading: { fill: "C2D2F1", type: ShadingType.CLEAR }
                    }
                },

                // Zwischenüberschrift — Segoe UI Semibold, spacing before=400
                { id: "Zwischenberschrift", name: "Zwischenüberschrift", basedOn: "Normal", next: "Normal", quickFormat: true,
                    run:       { font: "Segoe UI Semibold", size: SZ.body, color: "100C2A" },
                    paragraph: { spacing: { before: 400, line: 240 }, keepNext: true } },
            ]
        },
        sections: [{
            properties: {
                page: {
                    size:   { width: PAGE_W, height: PAGE_H },
                    margin: { top: MARGIN_TOP, right: MARGIN_SIDE, bottom: MARGIN_BOT, left: MARGIN_SIDE, header: HEADER_H, footer: 0 }
                }
            },
            headers: { default: buildHeader(docTitle) },
            footers: { default: buildFooter() },
            children: elements
        }]
    });
}

// ─── Main ─────────────────────────────────────────────────────────────────────
const md = fs.readFileSync(inputPath, 'utf-8');
console.log(`📄 Lese:   ${inputPath}`);
if (hasLogo) console.log(`🖼  Logo:   ${logoPath}`);

const { elements, numberingConfigs, docTitle, hasTOC } = parseMarkdown(md);
const doc = buildDocument(elements, numberingConfigs, docTitle);

Packer.toBuffer(doc).then(buf => {
    fs.writeFileSync(outputPath, buf);
    console.log(`✅ Fertig: ${outputPath}`);
    if (hasTOC) {
        console.log(`ℹ  TOC: In Word Strg+A → F9 drücken um Inhaltsverzeichnis zu aktualisieren.`);
    }
}).catch(err => {
    console.error('Fehler:', err.message);
    process.exit(1);
});
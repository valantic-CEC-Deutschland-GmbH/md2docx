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
 *   *kursiv* / _kursiv_ → Italic, ***fett+kursiv***
 *   `inline code`   → Consolas 11pt
 *   [Text](url)     → Hyperlink (http/https/mailto), relative Links als Text
 *   \* \_ \| \`     → Escapes, Zeichen wird wörtlich ausgegeben
 *   ```codeblock``` → Consolas eingerückt
 *   - / * / + Liste → Bullet (verschachtelt bis 3 Ebenen, mehrzeilige Punkte)
 *   1. Liste        → Nummeriert (auch gemischt mit Bullets)
 *   > Blockquote    → Blauer linker Rand
 *   | Tabelle |     → CI-Tabelle, Spaltenbreite nach Inhalt, kleinere Schrift ab 5 Spalten
 *   ---             → Trennlinie
 *   [TOC]           → Inhaltsverzeichnis (Headings 1-4)
 */

'use strict';

const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    Header, Footer, HeadingLevel, BorderStyle, WidthType, ShadingType,
    AlignmentType, LevelFormat, VerticalAlign, ImageRun, TableOfContents,
    PageNumber, ExternalHyperlink
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
const LINK_CLR  = "0563C1";

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
// Escapes (\* \_ \| \` \[ \] \\) werden vor dem Parsen durch Platzhalter aus dem
// Private-Use-Bereich ersetzt und danach wiederhergestellt.
const ESC_RX = /\\([\\`*_\[\]|#])/g;
const escapeMd   = s => s.replace(ESC_RX, (_, c) => String.fromCharCode(0xE000 + c.charCodeAt(0)));
const unescapeMd = s => s.replace(/[-]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xE000));

// Reihenfolge: Code, Link, fett+kursiv, fett, kursiv (*), kursiv (_)
const INLINE_RX = new RegExp([
    '`([^`]+?)`',                                         // 1 code
    '\\[([^\\]]+)\\]\\(([^)\\s]+)\\)',                    // 2 linktext, 3 url
    '\\*\\*\\*(.+?)\\*\\*\\*',                            // 4 fett+kursiv
    '\\*\\*(.+?)\\*\\*',                                  // 5 fett
    '(?<![\\w*])\\*(?![\\s*])(.+?)(?<![\\s*])\\*(?![\\w*])', // 6 kursiv *
    '(?<![\\w_])_(?![\\s_])(.+?)(?<![\\s_])_(?![\\w_])',     // 7 kursiv _
].join('|'), 'g');

function parseInlineEscaped(text, base) {
    const runs = [];
    const rx   = new RegExp(INLINE_RX.source, 'g');
    let last   = 0, m;
    while ((m = rx.exec(text)) !== null) {
        if (m.index > last) runs.push({ ...base, text: text.slice(last, m.index) });
        if      (m[1] !== undefined) runs.push({ ...base, font: "Consolas", text: m[1] });
        else if (m[2] !== undefined) runs.push(...parseInlineEscaped(m[2], { ...base, link: unescapeMd(m[3]) }));
        else if (m[4] !== undefined) runs.push(...parseInlineEscaped(m[4], { ...base, bold: true, italics: true }));
        else if (m[5] !== undefined) runs.push(...parseInlineEscaped(m[5], { ...base, bold: true }));
        else if (m[6] !== undefined) runs.push(...parseInlineEscaped(m[6], { ...base, italics: true }));
        else                         runs.push(...parseInlineEscaped(m[7], { ...base, italics: true }));
        last = m.index + m[0].length;
    }
    if (last < text.length) runs.push({ ...base, text: text.slice(last) });
    return runs;
}

function parseInline(text, base = {}) {
    const runs = parseInlineEscaped(escapeMd(text), base).map(r => ({ ...r, text: unescapeMd(r.text) }));
    return runs.length ? runs : [{ ...base, text: unescapeMd(escapeMd(text)) }];
}

// Text ohne Markdown-Formatierung (für Überschriften, Header, Spaltenbreiten)
function plainText(text) {
    return parseInline(text).map(r => r.text).join('');
}

const isExternalLink = url => /^(https?:|mailto:)/i.test(url || '');

// Baut TextRuns; aufeinanderfolgende Runs mit gleichem externen Link werden zu
// einem anklickbaren Hyperlink zusammengefasst. Relative Links (z. B. auf andere
// .md-Dateien) werden als normaler Text ausgegeben.
// opts.boldFont: Schrift für fetten Text (Standard: Segoe UI Semibold)
function toRuns(text, base = {}, opts = {}) {
    const boldFont = opts.boldFont || "Segoe UI Semibold";
    const out = [];
    let linkGroup = null;
    for (const r of parseInline(text, base)) {
        const link = isExternalLink(r.link) ? r.link : null;
        const run  = new TextRun({
            font:      r.font || (r.bold ? boldFont : "Segoe UI"),
            size:      r.size || SZ.body,
            bold:      !!r.bold,
            italics:   !!r.italics,
            color:     link ? LINK_CLR : (r.color || undefined),
            underline: link ? {} : undefined,
            text:      r.text
        });
        if (link) {
            if (!linkGroup || linkGroup.link !== link) {
                linkGroup = { link, children: [] };
                out.push(linkGroup);
            }
            linkGroup.children.push(run);
        } else {
            linkGroup = null;
            out.push(run);
        }
    }
    return out.map(x => x instanceof TextRun ? x : new ExternalHyperlink({ link: x.link, children: x.children }));
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
        children: [new TextRun({ text: plainText(text), font: HEADING_FONT, size: HEADING_SIZE[level], bold: level === 1 })]
    });
}

function bodyPara(text, opts = {}) {
    return new Paragraph({
        spacing:   { before: opts.before || 60, after: opts.after || 60 },
        numbering:  opts.numbering,
        indent:     opts.indent,
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
    // Zellen nur an nicht maskierten "|" trennen (\| bleibt Teil des Zellinhalts)
    return line.trim().replace(/^\|/, '').replace(/(?<!\\)\|$/, '').split(/(?<!\\)\|/).map(c => c.trim());
}

function isAlignRow(line) {
    return /^\|[\s|:-]+\|$/.test(line.trim());
}

// Schriftgröße in Tabellen: bei vielen Spalten kleiner, damit der Text lesbar umbricht
function tableFontSize(n) {
    if (n >= 10) return 16;  // 8 pt
    if (n >= 6) return 18;   // 9 pt
    if (n === 5) return 20;  // 10 pt
    return SZ.body;          // 11 pt
}

// Spaltenbreiten nach Inhalt: schmale Spalten (z. B. "#", Zahlen) werden schmal,
// Spalten mit langem Text bekommen mehr Platz. Das längste Wort einer Spalte
// bestimmt die Mindestbreite, damit Wörter nicht mitten im Wort umbrechen.
function columnWidths(headerCells, dataRows, fontSize) {
    const n = headerCells.length;
    const charW = fontSize * 5.5;            // ungefähre Zeichenbreite in DXA (Segoe UI)
    const pad   = 200;                        // Zellenränder links + rechts
    const minW  = [], wantW = [];
    for (let c = 0; c < n; c++) {
        const texts   = [headerCells[c], ...dataRows.map(r => r[c] || '')].map(plainText);
        const longest = Math.max(...texts.map(t => t.length), 1);
        const word    = Math.max(...texts.flatMap(t => t.split(/\s+/)).map(w => w.length), 1);
        minW.push(Math.min(word, 18) * charW + pad);
        wantW.push(Math.min(longest, 80) * charW + pad);
    }
    const sumMin  = minW.reduce((a, b) => a + b, 0);
    const sumWant = wantW.reduce((a, b) => a + b, 0);
    let colW;
    if (sumWant <= CONTENT_W) {
        // alles passt: überschüssigen Platz proportional verteilen
        colW = wantW.map(w => w * CONTENT_W / sumWant);
    } else if (sumMin >= CONTENT_W) {
        colW = minW.map(w => w * CONTENT_W / sumMin);
    } else {
        // Mindestbreite garantieren, Rest nach Bedarf verteilen
        const rest  = CONTENT_W - sumMin;
        const extra = wantW.map((w, i) => w - minW[i]);
        const sumEx = extra.reduce((a, b) => a + b, 0) || 1;
        colW = minW.map((w, i) => w + rest * extra[i] / sumEx);
    }
    colW = colW.map(Math.floor);
    colW[n - 1] += CONTENT_W - colW.reduce((a, b) => a + b, 0);
    return colW;
}

function buildTable(headerCells, dataRows) {
    const n        = headerCells.length;
    const fontSize = tableFontSize(n);
    const colW     = columnWidths(headerCells, dataRows, fontSize);

    const none     = { style: BorderStyle.NONE,   size: 0, color: "FFFFFF" };
    const solidBdr = { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" };

    function makeCell(text, isHdr, colIdx, altRow) {
        const runs = toRuns(text, { size: fontSize, bold: isHdr || undefined }, { boldFont: "Segoe UI" });
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
            children: [new TextRun({ text: plainText(titleText), font: "Segoe UI", size: 18, color: "555555" })]
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

    const BULLET_CHARS = ["•", "◦", "▪"];
    const levelIndent  = lvl => ({ left: 360 * (lvl + 1), hanging: 360 });

    function newBulletRef() {
        const ref = `bullet-${bulletRef++}`;
        numberingConfigs.push({ reference: ref, levels: [0, 1, 2].map(lvl => ({
                level: lvl, format: LevelFormat.BULLET, text: BULLET_CHARS[lvl], alignment: AlignmentType.LEFT,
                style: { paragraph: { indent: levelIndent(lvl) } }
            }))});
        return ref;
    }

    function newNumberRef() {
        const ref = `number-${numberRef++}`;
        numberingConfigs.push({ reference: ref, levels: [0, 1, 2].map(lvl => ({
                level: lvl, format: LevelFormat.DECIMAL, text: `%${lvl + 1}.`, alignment: AlignmentType.LEFT,
                style: { paragraph: { indent: levelIndent(lvl) } }
            }))});
        return ref;
    }

    // Erkennt Zeilen, die einen neuen Block beginnen (beendet Absätze und Listenpunkte)
    function isBlockStart(t) {
        return /^#{1,6}\s/.test(t) || t.startsWith('|') || t.startsWith('>') || t.startsWith('```') ||
               /^---+$/.test(t) || /^\[TOC\]$/i.test(t) || /^\{toc\}$/i.test(t);
    }

    const LIST_RX = /^(\s*)([-*+]|\d+\.)\s+(.*)$/;

    // Liest eine komplette Liste ab Zeile i: mehrzeilige Punkte, Verschachtelung
    // (Einrückung 2+ Leerzeichen bzw. Tab = eine Ebene, max. 3 Ebenen) und
    // gemischte Listen (z. B. Nummern mit Unterpunkten). Eine Liste läuft über
    // Leerzeilen weiter, solange danach wieder ein Listenpunkt folgt.
    function parseList(i) {
        const refs   = {};              // "bullet-0", "number-1" → reference (pro Liste und Ebene)
        const indents = [];             // Einrückung (Spalten) je Ebene
        // Gleicher Typ wie die oberste Ebene → dieselbe Referenz (Word nummeriert
        // Unterebenen dann pro Oberpunkt neu); anderer Typ → eigene Referenz.
        const refFor = (ordered, lvl) => {
            const type = ordered ? 'n' : 'b';
            const key  = refs[`${type}0`] ? `${type}0` : `${type}${lvl}`;
            if (!refs[key]) refs[key] = ordered ? newNumberRef() : newBulletRef();
            return refs[key];
        };
        while (i < lines.length) {
            const m = lines[i].replace(/\t/g, '    ').match(LIST_RX);
            if (!m) break;
            const col = m[1].length;
            while (indents.length && col < indents[indents.length - 1]) indents.pop();
            if (!indents.length || col > indents[indents.length - 1]) indents.push(col);
            const lvl     = Math.min(indents.length - 1, 2);
            const ordered = /\d/.test(m[2]);
            let text = m[3];
            i++;
            // Folgezeilen des Listenpunkts (eingerückt oder "lazy") anhängen
            while (i < lines.length) {
                const t = lines[i].trim();
                if (t === '' || isBlockStart(t) || LIST_RX.test(lines[i].replace(/\t/g, '    '))) break;
                text += ' ' + t;
                i++;
            }
            elements.push(bodyPara(text, { numbering: { reference: refFor(ordered, lvl), level: lvl }, before: 40, after: 40 }));
            // Leerzeile(n) innerhalb der Liste überspringen, wenn danach ein Listenpunkt folgt
            let j = i;
            while (j < lines.length && lines[j].trim() === '') j++;
            if (j > i && j < lines.length && LIST_RX.test(lines[j].replace(/\t/g, '    '))) i = j;
        }
        return i;
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

        if (LIST_RX.test(line.replace(/\t/g, '    '))) {
            i = parseList(i);
            continue;
        }

        // Normaler Absatz (mehrzeilig zusammenfassen)
        let paraText = trimmed;
        i++;
        while (
            i < lines.length &&
            lines[i].trim() !== '' &&
            !isBlockStart(lines[i].trim()) &&
            !LIST_RX.test(lines[i].replace(/\t/g, '    '))
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
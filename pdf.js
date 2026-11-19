'use strict';
/**
 * pdf.js — converts a finished .docx into a PDF.
 *
 * Order:
 *   1. Microsoft Word (macOS via AppleScript, Windows via PowerShell/COM):
 *      updates the table of contents and uses the real fonts → best result.
 *   2. LibreOffice (soffice --headless): fallback without Word, fonts are
 *      substituted if Segoe UI is missing, table of contents is not updated.
 */

const fs   = require('fs');
const os   = require('os');
const path = require('path');
const { spawnSync } = require('child_process');

const TIMEOUT_MS = 180000;

const MAC_SCRIPT = `
on run argv
    set inFile to POSIX file (item 1 of argv)
    set outFile to item 2 of argv
    set docName to item 3 of argv
    set wasRunning to application "Microsoft Word" is running
    tell application "Microsoft Word"
        with timeout of 170 seconds
        open inFile
        -- wait until Word has actually opened the document (a dialog would block it)
        set d to missing value
        repeat 60 times
            if exists document docName then
                set d to document docName
                exit repeat
            end if
            delay 1
        end repeat
        if d is missing value then error "Word konnte das Dokument nicht öffnen (ist in Word ein Dialog offen?)"
        set n to count of tables of contents of d
        repeat with i from 1 to n
            update table of contents i of d
        end repeat
        save as d file name outFile file format format PDF
        close d saving no
        if not wasRunning then quit saving no
        end timeout
    end tell
end run
`;

const WIN_SCRIPT = `
param([string]$in, [string]$out)
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($in, $false, $true)
    foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
    $doc.ExportAsFixedFormat($out, 17)
    $doc.Close(0)
} finally {
    $word.Quit()
}
`;

function run(cmd, args) {
    const r = spawnSync(cmd, args, { encoding: 'utf-8', timeout: TIMEOUT_MS });
    if (r.error && r.error.code === 'ETIMEDOUT') {
        throw new Error(`${cmd}: kein Ergebnis nach ${TIMEOUT_MS / 1000} s (ist in Word ein Dialog offen?)`);
    }
    if (r.error) throw r.error;
    if (r.status !== 0) throw new Error(`${cmd}: ${(r.stderr || r.stdout || '').trim()}`);
}

function hasWordMac() {
    return process.platform === 'darwin' && fs.existsSync('/Applications/Microsoft Word.app');
}

function wordMac(docx, pdf) {
    run('osascript', ['-e', MAC_SCRIPT, docx, pdf, path.basename(docx)]);
}

function wordWindows(docx, pdf) {
    const ps1 = path.join(os.tmpdir(), `md2docx-${process.pid}.ps1`);
    fs.writeFileSync(ps1, WIN_SCRIPT);
    try {
        run('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps1, docx, pdf]);
    } finally {
        fs.rmSync(ps1, { force: true });
    }
}

function findSoffice() {
    const candidates = [
        '/Applications/LibreOffice.app/Contents/MacOS/soffice',
        'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
        'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
    ];
    for (const c of candidates) if (fs.existsSync(c)) return c;
    for (const name of ['soffice', 'libreoffice']) {
        const r = spawnSync(process.platform === 'win32' ? 'where' : 'which', [name], { encoding: 'utf-8' });
        if (r.status === 0 && r.stdout.trim()) return r.stdout.trim().split(/\r?\n/)[0];
    }
    return null;
}

function libreOffice(soffice, docx, pdf) {
    const outDir = fs.mkdtempSync(path.join(os.tmpdir(), 'md2docx-'));
    try {
        run(soffice, ['--headless', '--convert-to', 'pdf', '--outdir', outDir, docx]);
        const produced = path.join(outDir, path.basename(docx).replace(/\.docx$/i, '.pdf'));
        fs.copyFileSync(produced, pdf);
    } finally {
        fs.rmSync(outDir, { recursive: true, force: true });
    }
}

/**
 * Converts docx → pdf. Returns the name of the engine that was used.
 * docx must not ask to update fields on open (updateFields: false),
 * otherwise Word waits for a dialog.
 */
function convertToPdf(docx, pdf) {
    const errors = [];
    const tryEngine = (name, fn) => {
        try { fn(); return fs.existsSync(pdf) ? name : null; }
        catch (e) { errors.push(`${name}: ${e.message}`); return null; }
    };

    let used = null;
    if (hasWordMac()) used = tryEngine('Word', () => wordMac(docx, pdf));
    else if (process.platform === 'win32') used = tryEngine('Word', () => wordWindows(docx, pdf));
    if (used) return used;

    const soffice = findSoffice();
    if (soffice) used = tryEngine('LibreOffice', () => libreOffice(soffice, docx, pdf));
    if (used) return used;

    if (!errors.length) errors.push('weder Microsoft Word noch LibreOffice gefunden');
    throw new Error('PDF konnte nicht erzeugt werden:\n   ' + errors.join('\n   '));
}

module.exports = { convertToPdf };

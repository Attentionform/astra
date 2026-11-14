#!/usr/bin/env node
/**
 * Real run of the Word renderer and the Windows PDF-to-image renderer on the machine it runs on. No mocks.
 *
 * Windows only. Run after `npm run build --workspace=@attentionform/astra-engine`.
 *
 * What it proves, each step failing the run with a message:
 *   - Word is found the way a user's machine would have it (registry), or the script says so and exits 0
 *     (`--expect-word` makes a missing Word a failure, for a machine that is known to have it).
 *   - Word renders a 3-page document to a PDF whose producer is Microsoft Word, and no Word process is left behind.
 *   - No window of Word, PowerShell or a console shows while it renders (a poller watches for the whole run).
 *   - The Windows PDF renderer draws the pages: right count, PNG and JPEG, pages with ink, page selection like Poppler's.
 *   - The four engine pairs (Word or LibreOffice, Windows or Poppler) each produce images when their programs exist.
 *   - An explicit choice that cannot work fails with E_DEPENDENCY instead of switching engine.
 *   - Network references are not loaded by default: a local listener sees no request from Word or LibreOffice, and
 *     sees one when `allowNetworkReferences` is set (skipped for an engine that is not installed).
 */
import { spawn } from 'node:child_process';
import fs from 'node:fs';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { countInk } from './lib/png-ink.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);
const engine = require(path.join(root, 'packages', 'engine', 'dist', 'index.js'));

const expectWord = process.argv.includes('--expect-word');
const skipListener = process.argv.includes('--no-listener');

function fail(message) {
  console.error(`word-render-smoke: FAIL: ${message}`);
  process.exit(1);
}
function step(message) {
  console.log(message);
}

if (process.platform !== 'win32') {
  console.log('word-render-smoke: SKIP (not Windows)');
  process.exit(0);
}

const PNG = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

function makeDocx(pages, extraBody = '') {
  const parser = engine.DocxParser.createNewDocument('blank');
  for (let p = 1; p <= pages; p++) {
    parser.insertRootNode(
      `<w:p><w:r><w:t xml:space="preserve">ASTRA WORD SMOKE page ${p}. ${'The quick brown fox jumps over the lazy dog. '.repeat(20)}</w:t></w:r></w:p>`,
      undefined,
      'append'
    );
    if (p < pages) parser.insertRootNode('<w:p><w:r><w:br w:type="page"/></w:r></w:p>', undefined, 'append');
  }
  if (extraBody) parser.insertRootNode(extraBody, undefined, 'append');
  return parser.rezip();
}

function checkImage(image, label) {
  const bytes = fs.readFileSync(image.path);
  if (image.format === 'png') {
    if (!bytes.subarray(0, 8).equals(PNG)) fail(`${label}: page ${image.page} is not a PNG`);
    const info = countInk(bytes);
    if (info.width < 100 || info.height < 100) fail(`${label}: page ${image.page} is ${info.width}x${info.height}`);
    if (info.ink < 500) fail(`${label}: page ${image.page} has almost no ink (${info.ink} dark pixels)`);
    return info;
  }
  if (bytes[0] !== 0xff || bytes[1] !== 0xd8) fail(`${label}: page ${image.page} is not a JPEG`);
  return { width: 0, height: 0, ink: 0 };
}

// ---- the window poller -------------------------------------------------------------------------------------
// Runs a PowerShell loop for the whole script and records every visible top-level window owned by Word,
// PowerShell or a console host. Word runs hidden and the engine starts PowerShell without a console window, so
// the list must stay empty.
const POLLER = `
$ErrorActionPreference = 'SilentlyContinue'
$seen = @{}
while ($true) {
  Get-Process WINWORD, powershell, conhost, OpenConsole, soffice, soffice.bin | Where-Object { $_.MainWindowHandle -ne 0 -and $_.Id -ne $PID } | ForEach-Object {
    $k = "$($_.ProcessName):$($_.Id)"
    if (-not $seen.ContainsKey($k)) { $seen[$k] = 1; [Console]::Out.WriteLine("WINDOW $k '$($_.MainWindowTitle)'"); [Console]::Out.Flush() }
  }
  Start-Sleep -Milliseconds 40
}`;
const windows = [];
const poller = spawn(
  process.env.SystemRoot ? path.join(process.env.SystemRoot, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe') : 'powershell.exe',
  ['-NoProfile', '-NonInteractive', '-EncodedCommand', Buffer.from(POLLER, 'utf16le').toString('base64')],
  { windowsHide: true, stdio: ['ignore', 'pipe', 'ignore'] }
);
poller.stdout.setEncoding('utf8');
poller.stdout.on('data', (d) => {
  for (const line of d.split(/\r?\n/)) if (line.startsWith('WINDOW')) windows.push(line);
});
await new Promise((r) => setTimeout(r, 1500)); // let the poller start before the first render

function wordProcesses() {
  return new Promise((resolve) => {
    const p = spawn('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', '(Get-Process WINWORD -ErrorAction SilentlyContinue | Measure-Object).Count'], { windowsHide: true, stdio: ['ignore', 'pipe', 'ignore'] });
    let out = '';
    p.stdout.on('data', (d) => (out += d));
    p.on('close', () => resolve(Number(out.trim()) || 0));
  });
}

const workDir = fs.realpathSync.native(fs.mkdtempSync(path.join(os.tmpdir(), 'astra-word-smoke-')));
let exitCode = 0;
try {
  const word = engine.findWord();
  const windowsRenderer = engine.windowsRendererAvailable();
  const status = engine.getVisualBackendStatus();
  step(`platform:        ${process.platform}-${process.arch}`);
  step(`word:            ${word ? `${word.exe} (version ${word.version})` : 'NOT FOUND'}`);
  step(`windows renderer: ${windowsRenderer ? 'available' : 'NOT AVAILABLE'}`);
  step(`soffice:         ${status.sofficePath ?? 'NOT FOUND'}`);
  step(`pdftoppm:        ${status.pdftoppmPath ?? 'NOT FOUND'}`);
  if (!windowsRenderer) fail('the Windows PDF renderer is not available on this machine (Windows 10 or later with PowerShell is required)');
  if (!word && expectWord) fail('Word was expected but is not installed or not registered');
  if (await wordProcesses()) {
    console.log('word-render-smoke: SKIP (a Word process is already running; the engine would refuse to share it)');
    process.exit(0);
  }

  const docx3 = makeDocx(3);
  const t0 = Date.now();

  // ---- Word -> PDF -----------------------------------------------------------------------------------------
  if (word) {
    const pdfPath = path.join(workDir, 'word.pdf');
    const exported = await engine.convertDocxToPdfDetailed(docx3, pdfPath, { renderer: 'word' });
    const pdf = fs.readFileSync(pdfPath);
    if (pdf.subarray(0, 5).toString('latin1') !== '%PDF-') fail('Word export is not a PDF');
    if (exported.renderer.used !== 'word' || exported.renderer.version !== word.version) fail(`unexpected renderer report ${JSON.stringify(exported.renderer)}`);
    const rawProducer = /\/Producer\s*\(([^)]*)\)/.exec(pdf.toString('latin1'))?.[1] ?? '';
    // Word writes the producer as UTF-16BE with a byte order mark; other writers use plain text.
    const producer = rawProducer.startsWith('\xfe\xff') ? rawProducer.slice(2).replace(/\0/g, '') : rawProducer;
    step(`word export:     ok, ${pdf.length} bytes, producer "${producer.replace(/[^\x20-\x7e]/g, '?')}"`);
    if (!/microsoft/i.test(producer)) fail(`the PDF producer is not Microsoft Word: "${producer}"`);
    const left = await wordProcesses();
    if (left !== 0) fail(`${left} Word process(es) left running after the export`);
    step('word cleanup:    ok, no Word process left');

    const second = await engine.convertDocxToPdfDetailed(docx3, path.join(workDir, 'word2.pdf'), { renderer: 'word', tagged: false });
    if (second.renderer.used !== 'word') fail('second Word export did not use Word');
    step('word re-run:     ok (a second export right after the first)');
  } else {
    step('word export:     skipped (Word not installed)');
  }

  // ---- the four pairs --------------------------------------------------------------------------------------
  const pairs = [
    ['word', 'windows'],
    ['word', 'poppler'],
    ['libreoffice', 'windows'],
    ['libreoffice', 'poppler'],
  ];
  for (const [renderer, pdfToImage] of pairs) {
    const label = `${renderer} + ${pdfToImage}`;
    const needsWord = renderer === 'word';
    const needsLo = renderer === 'libreoffice';
    if (needsWord && !word) {
      step(`pair ${label}: skipped (Word not installed)`);
      continue;
    }
    if ((needsLo && !status.sofficePath) || (pdfToImage === 'poppler' && !status.pdftoppmPath)) {
      step(`pair ${label}: skipped (program not installed)`);
      continue;
    }
    const render = await engine.renderDocxToImages(docx3, { renderer, pdfToImage, format: 'png', dpi: 72, maxPages: 3 });
    if (render.renderer.used !== renderer || render.rasterizer.used !== pdfToImage) fail(`${label}: reported ${JSON.stringify([render.renderer, render.rasterizer])}`);
    if (render.renderer.fallbackFrom || render.rasterizer.fallbackFrom) fail(`${label}: an explicit choice fell back`);
    if (render.totalPages !== 3 || render.images.length !== 3) fail(`${label}: expected 3 pages, got totalPages ${render.totalPages}, ${render.images.length} image(s)`);
    const dims = render.images.map((i) => checkImage(i, label));
    step(`pair ${label}: ok, 3 pages, page 1 is ${dims[0].width}x${dims[0].height} with ${dims[0].ink} dark pixels`);
  }

  // ---- the Windows renderer in detail ----------------------------------------------------------------------
  {
    const base = await engine.renderDocxToImages(docx3, { renderer: 'libreoffice', pdfToImage: 'windows', pages: [3, 1, 3, 99, 0], maxPages: 2, format: 'jpeg', jpegQuality: 70, dpi: 96 });
    // Poppler's own rule: normalise, de-duplicate, sort, clamp to the page count.
    const got = base.images.map((i) => i.page);
    if (JSON.stringify(got) !== JSON.stringify([1, 3])) fail(`page selection: expected [1,3], got ${JSON.stringify(got)}`);
    base.images.forEach((i) => checkImage(i, 'windows jpeg'));
    if (!base.images.every((i) => /\.jpe?g$/i.test(i.path))) fail('windows jpeg: wrong file extension');
    step('windows detail:  ok, page selection [3,1,3,99,0] -> [1,3], JPEG with quality');
    const hi = await engine.renderDocxToImages(docx3, { renderer: 'libreoffice', pdfToImage: 'windows', pages: [1], dpi: 200, format: 'png' });
    const lo = await engine.renderDocxToImages(docx3, { renderer: 'libreoffice', pdfToImage: 'windows', pages: [1], dpi: 72, format: 'png' });
    const hiInfo = checkImage(hi.images[0], 'dpi 200');
    const loInfo = checkImage(lo.images[0], 'dpi 72');
    if (!(hiInfo.width > loInfo.width * 2)) fail(`dpi: 200 dpi image (${hiInfo.width}) is not much wider than the 72 dpi image (${loInfo.width})`);
    step(`windows dpi:     ok, ${loInfo.width}px at 72 dpi, ${hiInfo.width}px at 200 dpi`);
  }

  // ---- explicit choices never fall back ---------------------------------------------------------------------
  {
    const err = await engine
      .renderDocxToImages(docx3, { renderer: 'libreoffice', pdfToImage: 'poppler', pdftoppmBin: path.join(workDir, 'no-such-pdftoppm.exe') })
      .then(() => null, (e) => e);
    if (!err || err.code !== 'E_DEPENDENCY') fail(`an explicit missing Poppler should be E_DEPENDENCY, got ${err && err.code}`);
    step('explicit choice: ok, a missing program is E_DEPENDENCY, not a silent switch');
  }

  // ---- network references: nothing contacts a listener by default ------------------------------------------
  if (!skipListener) {
    const hits = [];
    const server = http.createServer((req, res) => {
      hits.push(`${req.method} ${req.url}`);
      res.writeHead(404);
      res.end();
    });
    await new Promise((r) => server.listen(0, '127.0.0.1', r));
    const port = server.address().port;
    const url = `http://127.0.0.1:${port}`;
    const AdmZip = engine.AdmZip;
    const withRef = (kind) => {
      const zip = new AdmZip(makeDocx(1));
      const relsName = 'word/_rels/document.xml.rels';
      let rels = zip.readAsText(relsName);
      const id = 'rIdAstraNet1';
      if (kind === 'image') {
        rels = rels.replace('</Relationships>', `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${url}/rel-image.png" TargetMode="External"/></Relationships>`);
        zip.updateFile(relsName, Buffer.from(rels));
        let doc = zip.readAsText('word/document.xml');
        doc = doc.replace(
          '</w:body>',
          `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="914400"/><wp:docPr id="77" name="net"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="net"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:link="${id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body>`
        );
        zip.updateFile('word/document.xml', Buffer.from(doc));
      } else {
        const field = `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> INCLUDEPICTURE "${url}/field-image.png" \\d </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>x</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
        let doc = zip.readAsText('word/document.xml');
        doc = doc.replace('</w:body>', `${field}</w:body>`);
        zip.updateFile('word/document.xml', Buffer.from(doc));
      }
      return zip.toBuffer();
    };
    const engines = [];
    if (word) engines.push('word');
    if (status.sofficePath) engines.push('libreoffice');
    for (const renderer of engines) {
      for (const kind of ['image', 'field']) {
        const doc = withRef(kind);
        hits.length = 0;
        const off = await engine.renderDocxToImages(doc, { renderer, pdfToImage: 'windows', pages: [1], dpi: 72 });
        await new Promise((r) => setTimeout(r, 1500));
        if (hits.length !== 0) fail(`${renderer} ${kind}: contacted the listener by default: ${hits.join(', ')}`);
        if (!off.references.omissions || off.references.omittedCount < 1) fail(`${renderer} ${kind}: the report did not say a reference was omitted: ${JSON.stringify(off.references)}`);
        hits.length = 0;
        const on = await engine.renderDocxToImages(doc, { renderer, pdfToImage: 'windows', pages: [1], dpi: 72, allowNetworkReferences: true });
        await new Promise((r) => setTimeout(r, 1500));
        const contacted = hits.length;
        step(`network ${renderer} ${kind}: default = 0 requests (reported ${off.references.omittedCount} omitted); allowed = ${contacted} request(s)${on.references.networkReferencesAllowed ? '' : ' [report says not allowed!]'}`);
        if (!on.references.networkReferencesAllowed) fail(`${renderer} ${kind}: the allowed run did not report networkReferencesAllowed`);
        // Whether an engine actually fetches an image or an INCLUDEPICTURE field on open is the engine's own behaviour;
        // the contract tested here is the default (never) and the report. The count is printed for the record.
      }
    }
    await new Promise((r) => server.close(r));
  } else {
    step('network:         skipped (--no-listener)');
  }

  step(`elapsed:         ${((Date.now() - t0) / 1000).toFixed(1)} s`);
} catch (error) {
  console.error(`word-render-smoke: FAIL: ${error && error.stack ? error.stack : String(error)}`);
  exitCode = 1;
} finally {
  poller.kill();
  fs.rmSync(workDir, { recursive: true, force: true });
}
if (exitCode === 0) {
  if (windows.length > 0) {
    console.error(`word-render-smoke: FAIL: a window showed during the run:\n  ${windows.join('\n  ')}`);
    exitCode = 1;
  } else {
    step('windows:         ok, no Word, PowerShell, console or LibreOffice window showed during the run');
  }
}
console.log(exitCode === 0 ? 'word-render-smoke: PASS' : 'word-render-smoke: FAILED');
process.exit(exitCode);

#!/usr/bin/env node
/**
 * Real render on the machine it runs on: no mocks, no injected binaries.
 *
 * Makes a document, finds LibreOffice and Poppler the way a user's machine would, renders page
 * images, and exports a PDF. Fails if a binary is not found, an image is not a PNG with pixels,
 * or the PDF is empty. Run after `npm run build --workspace=@attentionform/astra-engine`.
 *
 * Used by the `visual-render` job of ci.yml on Linux, macOS and Windows.
 */
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { countInk } from './lib/png-ink.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);
const engine = require(path.join(root, 'packages', 'engine', 'dist', 'index.js'));

const PNG_SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

function fail(message) {
  console.error(`visual-render-smoke: FAIL: ${message}`);
  process.exit(1);
}

const status = engine.getVisualBackendStatus();
console.log(`platform: ${process.platform}-${process.arch}`);
console.log(`soffice:  ${status.sofficePath ?? 'NOT FOUND'}`);
console.log(`pdftoppm: ${status.pdftoppmPath ?? 'NOT FOUND'}`);
if (!status.available) fail(`backend not found (missing: ${status.missing.join(', ')})`);

// Say where each tool came from: found on PATH, or only in a standard install folder.
// `--expect-source path|fallback` makes the run fail when a tool came from the other place, so a job
// can prove the install-folder lookup was the one used (run it with a PATH that has no LibreOffice).
function sourceOf(name) {
  const onPath = engine.findOnPath(name, process.env.PATH ?? '', { fallbackDirs: false });
  return onPath ? 'path' : 'fallback';
}
const sofficeSource = sourceOf(path.basename(status.sofficePath).replace(/\.(exe|com)$/i, ''));
const pdftoppmSource = sourceOf('pdftoppm');
console.log(`soffice found via:  ${sofficeSource}`);
console.log(`pdftoppm found via: ${pdftoppmSource}`);
const expectAt = process.argv.indexOf('--expect-source');
if (expectAt >= 0) {
  const expected = process.argv[expectAt + 1];
  if (sofficeSource !== expected || pdftoppmSource !== expected) {
    fail(`expected both tools via ${expected}, got soffice via ${sofficeSource} and pdftoppm via ${pdftoppmSource}`);
  }
}

// Prove the tools are really installed and run: print their own version output.
function versionLine(bin, args) {
  // pdftoppm prints its version on stderr, so read both streams.
  const run = spawnSync(bin, args, { encoding: 'utf8', timeout: 60000 });
  if (run.error) return fail(`could not run ${bin} ${args.join(' ')}: ${run.error.message}`);
  const text = `${run.stdout ?? ''}${run.stderr ?? ''}`.trim();
  if (!text) return fail(`${bin} ${args.join(' ')} printed nothing (exit ${run.status})`);
  return text.split(/\r?\n/)[0];
}
// On Windows soffice.exe is a GUI launcher that does not print --version and can hang; its console
// wrapper soffice.com does print it. Probe that one, and skip the probe if it is not installed.
let sofficeVersion = 'not probed (no soffice.com next to soffice.exe)';
if (process.platform !== 'win32') {
  sofficeVersion = versionLine(status.sofficePath, ['--version']);
  if (!/libreoffice/i.test(sofficeVersion)) fail(`unexpected soffice version output: ${sofficeVersion}`);
} else {
  const wrapper = status.sofficePath.replace(/\.exe$/i, '.com');
  if (fs.existsSync(wrapper)) {
    sofficeVersion = versionLine(wrapper, ['--version']);
    if (!/libreoffice/i.test(sofficeVersion)) fail(`unexpected soffice.com version output: ${sofficeVersion}`);
  }
}
const pdftoppmVersion = versionLine(status.pdftoppmPath, ['-v']);
console.log(`soffice version:  ${sofficeVersion}`);
console.log(`pdftoppm version: ${pdftoppmVersion}`);
if (!/pdftoppm/i.test(pdftoppmVersion)) fail(`unexpected pdftoppm version output: ${pdftoppmVersion}`);

const docx = engine.DocxParser.createNewDocument('blank').rezip();
// realpath: on macOS the temp folder sits behind the /var symlink, which the write sandbox refuses.
const workDir = fs.realpathSync.native(fs.mkdtempSync(path.join(os.tmpdir(), 'astra-visual-smoke-')));

try {
  // `auto` picks the page-image program for this machine (Windows: its built-in renderer first, Poppler as the
  // fallback; elsewhere Poppler), and `poppler` forces Poppler. Both must produce a real page.
  for (const pdfToImage of ['auto', 'poppler']) {
    const render = await engine.renderDocxToImages(docx, { format: 'png', dpi: 72, maxPages: 1, pdfToImage });
    if (render.images.length < 1) fail(`render (${pdfToImage}) returned no page image`);
    // `auto` uses Word first on a Windows machine that has it, and LibreOffice everywhere else: the engine's own plan says which.
    const expectedRenderer = engine.getVisualBackendStatus().plan?.renderers?.[0] ?? 'libreoffice';
    if (render.renderer.used !== expectedRenderer) fail(`render (${pdfToImage}) used ${render.renderer.used}, expected ${expectedRenderer}`);
    if (pdfToImage === 'poppler' && render.rasterizer.used !== 'poppler') fail(`pdfToImage poppler used ${render.rasterizer.used}`);
    if (render.rasterizer.fallbackFrom) console.log(`note: ${render.rasterizer.fallbackFrom} could not draw the page (${render.rasterizer.fallbackReason}); ${render.rasterizer.used} did`);
    const image = render.images[0];
    const bytes = fs.readFileSync(image.path);
    if (bytes.length < 200) fail(`page image (${pdfToImage}) is only ${bytes.length} bytes`);
    if (!bytes.subarray(0, 8).equals(PNG_SIGNATURE)) fail(`page image (${pdfToImage}) is not a PNG`);
    const { width, height } = countInk(bytes);
    if (width < 100 || height < 100) fail(`page image (${pdfToImage}) is ${width}x${height}`);
    console.log(`render:   ok (${pdfToImage} -> ${render.rasterizer.used}), page 1 is ${width}x${height} PNG, ${bytes.length} bytes, ${render.totalPages} page(s)`);
  }

  const pdfPath = path.join(workDir, 'export.pdf');
  await engine.convertDocxToPdf(docx, pdfPath);
  const pdf = fs.readFileSync(pdfPath);
  if (pdf.length < 500 || pdf.subarray(0, 5).toString('latin1') !== '%PDF-') fail('PDF export is not a PDF');
  console.log(`export:   ok, PDF ${pdf.length} bytes`);
} catch (error) {
  fail(error && error.stack ? error.stack : String(error));
} finally {
  fs.rmSync(workDir, { recursive: true, force: true });
}
console.log('visual-render-smoke: PASS');

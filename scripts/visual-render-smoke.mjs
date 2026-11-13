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

// Prove the tools are really installed and run: print their own version output.
function versionLine(bin, args) {
  // pdftoppm prints its version on stderr, so read both streams.
  const run = spawnSync(bin, args, { encoding: 'utf8', timeout: 120000 });
  if (run.error) return fail(`could not run ${bin} ${args.join(' ')}: ${run.error.message}`);
  const text = `${run.stdout ?? ''}${run.stderr ?? ''}`.trim();
  if (!text) return fail(`${bin} ${args.join(' ')} printed nothing (exit ${run.status})`);
  return text.split(/\r?\n/)[0];
}
const sofficeVersion = versionLine(status.sofficePath, ['--version']);
const pdftoppmVersion = versionLine(status.pdftoppmPath, ['-v']);
console.log(`soffice version:  ${sofficeVersion}`);
console.log(`pdftoppm version: ${pdftoppmVersion}`);
if (!/libreoffice/i.test(sofficeVersion)) fail(`unexpected soffice version output: ${sofficeVersion}`);
if (!/pdftoppm/i.test(pdftoppmVersion)) fail(`unexpected pdftoppm version output: ${pdftoppmVersion}`);

const docx = engine.DocxParser.createNewDocument('blank').rezip();
const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-visual-smoke-'));

try {
  const render = await engine.renderDocxToImages(docx, { format: 'png', dpi: 72, maxPages: 1 });
  if (render.images.length < 1) fail('render returned no page image');
  const image = render.images[0];
  const bytes = fs.readFileSync(image.path);
  if (bytes.length < 200) fail(`page image is only ${bytes.length} bytes`);
  if (!bytes.subarray(0, 8).equals(PNG_SIGNATURE)) fail('page image is not a PNG');
  const width = bytes.readUInt32BE(16);
  const height = bytes.readUInt32BE(20);
  if (width < 100 || height < 100) fail(`page image is ${width}x${height}`);
  console.log(`render:   ok, page 1 is ${width}x${height} PNG, ${bytes.length} bytes, ${render.totalPages} page(s)`);

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

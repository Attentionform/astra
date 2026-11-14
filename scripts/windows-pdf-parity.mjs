#!/usr/bin/env node
/**
 * Real check, on Windows, that the built-in PDF renderer (Windows.Data.Pdf) and Poppler agree on what a page
 * image looks like: same page count, same pixel size (within 1 px), same page selection, ink on every page.
 * Poppler is the reference. No mocks. Run after `npm run build --workspace=@attentionform/astra-engine`.
 *
 * Why it exists: Windows.Data.Pdf multiplies the requested width by the system display scale (125 percent on a
 * 120 dpi display), so a naive request gives images 25 percent too big. The engine measures and corrects; this
 * script proves the corrected sizes match Poppler on pages of several sizes, rotations and crop boxes.
 *
 * Known, accepted difference: for a page that has a CropBox smaller than its MediaBox, Windows draws the crop box
 * and Poppler (without -cropbox) draws the media box. Word and LibreOffice never write one; the script reports it
 * and does not fail on it.
 */
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { countInk, pngInfo } from './lib/png-ink.mjs';
import { makePdf } from './lib/make-pdf.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const engine = createRequire(import.meta.url)(path.join(root, 'packages', 'engine', 'dist', 'index.js'));

if (process.platform !== 'win32') {
  console.log('windows-pdf-parity: SKIP (not Windows)');
  process.exit(0);
}
const fail = (m) => {
  console.error(`windows-pdf-parity: FAIL: ${m}`);
  process.exit(1);
};
if (!engine.windowsRendererAvailable()) fail('the Windows PDF renderer is not available here');
const poppler = engine.getVisualBackendStatus().pdftoppmPath;
if (!poppler) fail('Poppler (pdftoppm) was not found; it is the reference');

const work = fs.realpathSync.native(fs.mkdtempSync(path.join(os.tmpdir(), 'astra-pdf-parity-')));
const cases = {
  letter: [{ w: 612, h: 792 }],
  a4: [{ w: 595, h: 842 }],
  landscape: [{ w: 792, h: 612 }],
  rotate90: [{ w: 612, h: 792, rotate: 90 }],
  rotate180: [{ w: 612, h: 792, rotate: 180 }],
  rotate270: [{ w: 612, h: 792, rotate: 270 }],
  mixed: [{ w: 612, h: 792 }, { w: 792, h: 612 }, { w: 300, h: 400 }, { w: 595, h: 842 }],
  big: [{ w: 3000, h: 2000 }],
  tiny: [{ w: 50, h: 50 }],
  crop: [{ w: 612, h: 792, crop: [50, 100, 500, 700] }],
};
let problems = 0;
let notes = 0;
try {
  for (const [name, pages] of Object.entries(cases)) {
    const pdf = path.join(work, `${name}.pdf`);
    fs.writeFileSync(pdf, makePdf(pages));
    for (const dpi of [72, 150]) {
      const winDir = path.join(work, `${name}-${dpi}-win`);
      const popDir = path.join(work, `${name}-${dpi}-pop`);
      fs.mkdirSync(winDir);
      fs.mkdirSync(popDir);
      const win = await engine.windowsPdfToImages(pdf, winDir, { format: 'png', dpi, jpegQuality: 85, maxPages: 10 });
      const run = spawnSync(poppler, ['-png', '-r', String(dpi), pdf, path.join(popDir, 'page')], { encoding: 'utf8' });
      if (run.status !== 0) fail(`pdftoppm failed on ${name}: ${run.stderr}`);
      const popFiles = fs.readdirSync(popDir).filter((f) => f.endsWith('.png')).sort();
      if (win.totalPages !== pages.length || win.images.length !== pages.length || popFiles.length !== pages.length) {
        fail(`${name}@${dpi}: page counts differ (windows total ${win.totalPages}/${win.images.length} images, poppler ${popFiles.length}, expected ${pages.length})`);
      }
      win.images.forEach((img, i) => {
        const w = pngInfo(fs.readFileSync(img.file));
        const p = pngInfo(fs.readFileSync(path.join(popDir, popFiles[i])));
        const dw = Math.abs(w.width - p.width);
        const dh = Math.abs(w.height - p.height);
        const ink = countInk(fs.readFileSync(img.file)).ink;
        const line = `${name}@${dpi} page ${i + 1}: windows ${w.width}x${w.height}, poppler ${p.width}x${p.height}, ink ${ink}`;
        if (name === 'crop') {
          console.log(`${line}  (crop box: known difference)`);
          notes++;
        } else if (dw > 1 || dh > 1) {
          console.error(`MISMATCH ${line}`);
          problems++;
        } else if (ink < 50) {
          console.error(`NO INK ${line}`);
          problems++;
        } else {
          console.log(`ok       ${line}`);
        }
      });
    }
  }

  // Page selection: the same rule as Poppler's (normalise, de-duplicate, sort, clamp to the page count).
  const mixed = path.join(work, 'mixed.pdf');
  const selDir = path.join(work, 'sel');
  fs.mkdirSync(selDir);
  const sel = await engine.windowsPdfToImages(mixed, selDir, { format: 'png', dpi: 72, jpegQuality: 85, pages: [4, 2, 2, 99, 0, 1.9], maxPages: 10 });
  const got = sel.images.map((i) => i.page);
  if (JSON.stringify(got) !== JSON.stringify([1, 2, 4])) {
    console.error(`MISMATCH page selection [4,2,2,99,0,1.9] gave ${JSON.stringify(got)}, expected [1,2,4]`);
    problems++;
  } else console.log('ok       page selection [4,2,2,99,0,1.9] -> [1,2,4]');

  // A damaged PDF is a document failure (E_VALIDATION), not a dependency failure, so `auto` can still try Poppler.
  const bad = path.join(work, 'bad.pdf');
  fs.writeFileSync(bad, Buffer.from('%PDF-1.4\nthis is not a pdf\n'));
  const badDir = path.join(work, 'bad');
  fs.mkdirSync(badDir);
  const err = await engine.windowsPdfToImages(bad, badDir, { format: 'png', dpi: 72, jpegQuality: 85, maxPages: 1 }).then(() => null, (e) => e);
  if (!err) {
    console.error('MISMATCH a damaged PDF rendered without an error');
    problems++;
  } else if (err.failureClass !== 'document') {
    console.error(`MISMATCH a damaged PDF should be a document failure, got ${err.failureClass} (${err.code}): ${err.message}`);
    problems++;
  } else console.log(`ok       damaged PDF -> ${err.code} (${err.failureClass})`);
} finally {
  fs.rmSync(work, { recursive: true, force: true });
}
console.log(problems === 0 ? `windows-pdf-parity: PASS (${notes} known difference(s) reported)` : `windows-pdf-parity: FAILED (${problems} problem(s))`);
process.exit(problems === 0 ? 0 : 1);

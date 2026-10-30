/**
 * Live binary probes — slow file, skipped unless explicitly enabled.
 *
 * File tag is the `.slow.test.ts` suffix plus the opt-in gate below.
 * Default runs skip every case (no process spawns, suite stays fast).
 * Run live with: ASTRA_RUN_SLOW=1 vitest run packages/engine/tests/live-libreoffice.slow.test.ts
 */
import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { DocxParser } from '../src/parser';
import { convertDocxToPdf, hasPdfBackend } from '../src/export/pdf';
import { hasPdfBackend as hasVisualBackend } from '../src/visual/backend';
import { renderDocxToImages, getVisualArtifactsDir } from '../src/visual/renderer';

const RUN_LIVE = process.env.ASTRA_RUN_SLOW === '1';

function blankDocx(): Buffer {
  return DocxParser.createNewDocument('blank').rezip();
}

describe.runIf(RUN_LIVE)('live binaries (opt-in, slow)', () => {
  it('pdf export renders a real file when the backend exists', async () => {
    if (!hasPdfBackend()) return;
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-live-pdf-'));
    try {
      const outPdf = path.join(dir, 'live.pdf');
      const resPath = await convertDocxToPdf(blankDocx(), outPdf);
      expect(resPath).toBe(outPdf);
      expect(fs.existsSync(outPdf)).toBe(true);
      expect(fs.statSync(outPdf).size).toBeGreaterThan(1000);
      const head = Buffer.alloc(5);
      const fd = fs.openSync(outPdf, 'r');
      try {
        fs.readSync(fd, head, 0, 5, 0);
      } finally {
        fs.closeSync(fd);
      }
      expect(head.toString('ascii')).toBe('%PDF-');
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  }, 60000);

  it('visual render produces page images when the backend exists', async () => {
    const status = hasVisualBackend();
    if (!status.available) return;
    const outDir = getVisualArtifactsDir(`live-${Date.now()}`);
    const res = await renderDocxToImages(blankDocx(), { format: 'png', outDir });
    expect(res.ok).toBe(true);
    expect(res.images.length).toBeGreaterThanOrEqual(1);
    for (const img of res.images) {
      expect(fs.existsSync(img.path)).toBe(true);
      expect(img.bytes).toBeGreaterThan(0);
    }
  }, 60000);
});

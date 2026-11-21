/**
 * The pages that LibreOffice draws have the size that was asked for: every named paper in both orientations, a custom size, the largest
 * page (22 in) and the smallest (1 in) a request may make. Opt-in: ASTRA_RUN_SLOW=1 (the CI workflow sets it).
 */
import { describe, expect, it } from 'vitest';
import { spawnSync } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { DocxParser } from '../src/parser';
import { convertDocxToPdf, hasPdfBackend } from '../src/export/pdf';
import { PAPERS } from '../src/page';

const RUN = process.env.ASTRA_RUN_SLOW === '1';
const has = (cmd: string): boolean => spawnSync(cmd, ['-v'], { stdio: 'ignore' }).error === undefined;

/** The first page size that pdfinfo reports, in points. */
function pageSizePt(pdf: string): { w: number; h: number } {
  const out = spawnSync('pdfinfo', [pdf], { encoding: 'utf8' }).stdout;
  const m = /Page size:\s+([\d.]+) x ([\d.]+) pts/.exec(out);
  if (!m) throw new Error(`no page size in pdfinfo output: ${out}`);
  return { w: Number(m[1]), h: Number(m[2]) };
}

describe.runIf(RUN && hasPdfBackend() && has('pdfinfo'))('page sizes drawn by LibreOffice (opt-in)', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-size-matrix-'));
  const cases: Array<{ label: string; page: object; w: number; h: number }> = [];
  for (const p of PAPERS) {
    cases.push({ label: `${p.name} portrait`, page: { size: p.name }, w: p.width / 20, h: p.height / 20 });
    cases.push({ label: `${p.name} landscape`, page: { size: p.name, orientation: 'landscape' }, w: p.height / 20, h: p.width / 20 });
  }
  cases.push({ label: 'custom 150x200mm', page: { size: '150x200mm' }, w: 425.2, h: 566.9 });
  cases.push({ label: 'largest 22x22in', page: { size: '22x22in', margins: '1in' }, w: 1584, h: 1584 });
  cases.push({ label: 'smallest 1x1in with the minimum text area', page: { size: '1x1in', margins: { top: '0.25in', bottom: '0.25in', left: '0.125in', right: '0.125in', header: '0', footer: '0' } }, w: 72, h: 72 });
  for (const c of cases) {
    it(c.label, async () => {
      const pdf = path.join(dir, `${c.label.replace(/\W+/g, '-')}.pdf`);
      await convertDocxToPdf(DocxParser.createNewDocument('blank', { page: c.page } as never).rezip(), pdf);
      const got = pageSizePt(pdf);
      expect(Math.abs(got.w - c.w)).toBeLessThan(1.5);
      expect(Math.abs(got.h - c.h)).toBeLessThan(1.5);
    }, 180_000);
  }
});

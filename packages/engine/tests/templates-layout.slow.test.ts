/**
 * Page count and last-page content of every template on four papers, rendered by LibreOffice with the metric-compatible fonts
 * (Carlito for Calibri, Caladea for Cambria). Opt-in: ASTRA_RUN_SLOW=1 (the CI workflow sets it). It guards the regression where a template
 * ends with a page that holds only the footer, and that a template's page count changes without anyone noticing.
 * The counts were checked against Microsoft Word 16 as well: they are the same.
 */
import { describe, expect, it } from 'vitest';
import { spawnSync } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { DocxParser } from '../src/parser';
import { convertDocxToPdf, hasPdfBackend } from '../src/export/pdf';

const RUN = process.env.ASTRA_RUN_SLOW === '1';
const has = (cmd: string): boolean => spawnSync(cmd, ['-v'], { stdio: 'ignore' }).error === undefined;
const fontsOk = (): boolean => {
  const r = spawnSync('fc-list', [], { encoding: 'utf8' });
  return r.status === 0 && /Carlito/i.test(r.stdout) && /Caladea/i.test(r.stdout);
};

const EXPECTED: Record<string, Record<string, number>> = {
  Letter: { invoice: 1, letter: 1, memo: 1, newsletter: 1, novel: 2, proposal: 3, report: 3, 'research-paper': 1, 'technical-spec': 2 },
  A4: { invoice: 1, letter: 1, memo: 1, newsletter: 1, novel: 2, proposal: 3, report: 3, 'research-paper': 1, 'technical-spec': 2 },
  A5: { invoice: 2, letter: 1, memo: 1, newsletter: 1, novel: 2, proposal: 4, report: 4, 'research-paper': 2, 'technical-spec': 2 },
  Legal: { invoice: 1, letter: 1, memo: 1, newsletter: 1, novel: 2, proposal: 3, report: 3, 'research-paper': 1, 'technical-spec': 1 },
};

describe.runIf(RUN && hasPdfBackend() && has('pdftotext') && fontsOk())('template layout (LibreOffice, opt-in)', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-tpl-layout-'));
  for (const [size, counts] of Object.entries(EXPECTED)) {
    for (const [name, pages] of Object.entries(counts)) {
      it(`${name} on ${size}: ${pages} page(s), and the last one holds body content`, async () => {
        const pdf = path.join(dir, `${name}-${size}.pdf`);
        await convertDocxToPdf(DocxParser.createNewDocument(name, { page: { size } }).rezip(), pdf);
        const count = /Pages:\s+(\d+)/.exec(spawnSync('pdfinfo', [pdf], { encoding: 'utf8' }).stdout)?.[1];
        expect(Number(count)).toBe(pages);
        const last = spawnSync('pdftotext', ['-f', String(pages), '-l', String(pages), '-layout', pdf, '-'], { encoding: 'utf8' }).stdout;
        // What is left after the running footer ("Page 3 of 3") and header lines is body text.
        const body = last.split('\n').map((l) => l.trim()).filter((l) => l !== '' && !/^Page \d+ of \d+$/.test(l) && !/^\[Surname\] \/ \[Title\] \/ \d+$/.test(l));
        expect(body.length, `last page of ${name} on ${size} has only furniture`).toBeGreaterThan(1);
      }, 180_000);
    }
  }
});

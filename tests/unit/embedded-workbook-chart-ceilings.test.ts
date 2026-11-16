/**
 * A chart data update compares the stored workbook with the one ASTra would generate. The stored workbook is a zip
 * from an untrusted file: when it breaks the size or ratio ceilings none of its entries may be inflated.
 */
import { describe, expect, it, vi } from 'vitest';
import AdmZip from 'adm-zip';

const inflated = vi.hoisted(() => [] as string[]);

// Records every entry that is inflated through the engine's zip wrapper.
vi.mock('../../packages/engine/src/adm-zip-safe', async (importOriginal) => {
  const mod = await importOriginal<{ default: new (...args: unknown[]) => { getEntries(): Array<{ entryName: string; getData: (...a: unknown[]) => Buffer }> } }>();
  const Real = mod.default;
  function Recording(this: unknown, ...args: unknown[]) {
    const zip = new Real(...args);
    for (const entry of zip.getEntries()) {
      const original = entry.getData.bind(entry);
      entry.getData = (...a: unknown[]) => { inflated.push(entry.entryName); return original(...a); };
    }
    return zip;
  }
  return { ...mod, default: Recording };
});

import { isGeneratedWorkbook } from '../../packages/engine/src/operations/chart-workbook';

/** An xlsx-looking zip whose one big entry is 30 MB of blanks (about 30 KB packed: a ratio of about 1000). */
function bombWorkbook(): Buffer {
  const zip = new AdmZip();
  zip.addFile('[Content_Types].xml', Buffer.from('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'));
  zip.addFile('xl/worksheets/sheet1.xml', Buffer.alloc(30 * 1024 * 1024, 0x20));
  return zip.toBuffer();
}
const stored = { categories: ['a'], series: [{ name: 'A', values: [1] }], scatter: false };

describe('chart data update reading an embedded workbook', () => {
  it('does not inflate any entry of a workbook that breaks the ratio ceiling', () => {
    const bomb = bombWorkbook();
    expect(bomb.length).toBeLessThan(200 * 1024);
    inflated.length = 0;
    expect(isGeneratedWorkbook(bomb, stored)).toBe(false);
    expect(inflated).not.toContain('xl/worksheets/sheet1.xml');
  });
  it('the recording works: an ordinary workbook is inflated by the same call', () => {
    const book = new AdmZip();
    book.addFile('[Content_Types].xml', Buffer.from('<Types/>'));
    book.addFile('xl/worksheets/sheet1.xml', Buffer.from('<worksheet/>'));
    inflated.length = 0;
    isGeneratedWorkbook(book.toBuffer(), stored);
    expect(inflated).toContain('xl/worksheets/sheet1.xml');
  });
});

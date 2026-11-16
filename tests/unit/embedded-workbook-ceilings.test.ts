/**
 * An embedded workbook is a zip inside the package. The outer ceilings do not see what it holds, so its own declared
 * sizes and ratios are checked before any entry is inflated: by a chart data update and by sanitize.
 */
import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { createBlankDocxBuffer, sanitizeDocxBuffer } from '@attentionform/astra-engine';

/** An xlsx-looking zip whose one big entry is 30 MB of zeros (about 30 KB packed: a ratio of about 1000). */
function bombWorkbook(): Buffer {
  const zip = new AdmZip();
  zip.addFile('[Content_Types].xml', Buffer.from('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'));
  zip.addFile('xl/worksheets/sheet1.xml', Buffer.alloc(30 * 1024 * 1024, 0x20));
  return zip.toBuffer();
}

describe('sanitize reading an embedded workbook', () => {
  it('refuses a workbook that breaks the ceilings, before inflating it', () => {
    const zip = new AdmZip(createBlankDocxBuffer('blank'));
    zip.addFile('word/embeddings/Microsoft_Excel_Worksheet1.xlsx', bombWorkbook());
    expect(() => sanitizeDocxBuffer(zip.toBuffer(), { purgePII: true })).toThrow(/zip bomb ceiling/);
  });
  it('still redacts an e-mail address in an ordinary embedded workbook', () => {
    const book = new AdmZip();
    book.addFile('[Content_Types].xml', Buffer.from('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'));
    book.addFile('xl/sharedStrings.xml', Buffer.from('<sst><si><t>write to jane.doe@example.com</t></si></sst>'));
    const zip = new AdmZip(createBlankDocxBuffer('blank'));
    zip.addFile('word/embeddings/Microsoft_Excel_Worksheet1.xlsx', book.toBuffer());
    const out = sanitizeDocxBuffer(zip.toBuffer(), { purgePII: true });
    const inner = new AdmZip(new AdmZip(out.buffer).getEntry('word/embeddings/Microsoft_Excel_Worksheet1.xlsx')!.getData());
    expect(inner.readAsText('xl/sharedStrings.xml')).not.toContain('jane.doe@example.com');
  });
});

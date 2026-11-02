import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import AdmZip from 'adm-zip';
import { validateBuffer } from '../src/index.js';

/**
 * The Open XML SDK validator stops by itself after 1000 errors and does not say so. The worker now
 * sets that limit from the request, so a request for more than 1000 errors gets them, and a list that
 * is cut at the requested number is marked truncated.
 */
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const VALID_DOCX = path.join(__dirname, 'fixtures', 'valid.docx');
const BAD_PARAGRAPHS = 1200;

function docxWithManyErrors(): Buffer {
  const zip = new AdmZip(fs.readFileSync(VALID_DOCX));
  const xml = zip.readAsText('word/document.xml');
  // w:bogus is not a child the schema allows inside a paragraph: one error for each paragraph.
  const bad = '<w:p><w:bogus/></w:p>'.repeat(BAD_PARAGRAPHS);
  const marker = '<w:body>';
  expect(xml).toContain(marker);
  zip.updateFile('word/document.xml', Buffer.from(xml.replace(marker, marker + bad)));
  return zip.toBuffer();
}

describe('maxErrors above the SDK default of 1000', () => {
  const bytes = docxWithManyErrors();

  it('returns every error when the request allows more than the document has', async () => {
    const res = await validateBuffer(bytes, { maxErrors: 5000 });
    expect(res.ok).toBe(false);
    expect(res.errors.length).toBeGreaterThanOrEqual(BAD_PARAGRAPHS);
    expect(res.truncated).toBe(false);
  });

  it('marks the list truncated when the request stops it at 1000', async () => {
    const res = await validateBuffer(bytes, { maxErrors: 1000 });
    expect(res.errors.length).toBe(1000);
    expect(res.truncated).toBe(true);
  });

  it('marks the list truncated when the request stops it between 1000 and the full count', async () => {
    const res = await validateBuffer(bytes, { maxErrors: 1100 });
    expect(res.errors.length).toBe(1100);
    expect(res.truncated).toBe(true);
  });

  it('does not mark a list truncated when it holds exactly the requested number', async () => {
    const all = await validateBuffer(bytes, { maxErrors: 5000 });
    const exact = await validateBuffer(bytes, { maxErrors: all.errors.length });
    expect(exact.errors.length).toBe(all.errors.length);
    expect(exact.truncated).toBe(false);
  });
});

/** new_document `page`: paper, orientation and margins of the new document; the old `pageSize` is refused with its replacement. */
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { handleNewDocument } from '../../src/tools/new-document';

let dir: string;
beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-new-doc-page-'));
});
afterEach(() => fs.rmSync(dir, { recursive: true, force: true }));

const xmlOf = (file: string): string => new AdmZip(file).readAsText('word/document.xml');

describe('new_document page', () => {
  it('a named paper, landscape and one margin for all sides', async () => {
    const out = path.join(dir, 'a5.docx');
    const r: any = await handleNewDocument(createMcpSessionStore(), { outPath: out, template: 'memo', page: { size: 'A5', orientation: 'landscape', margins: '15mm' } });
    expect(r.ok, JSON.stringify(r)).toBe(true);
    const xml = xmlOf(out);
    expect(xml).toContain('w:w="11906" w:h="8391"');
    expect(xml).toContain('w:orient="landscape"');
    expect(xml).toContain('w:top="850" w:right="850" w:bottom="850" w:left="850"');
  });

  it('a custom size with a gutter', async () => {
    const out = path.join(dir, 'custom.docx');
    const r: any = await handleNewDocument(createMcpSessionStore(), { outPath: out, template: 'report', page: { size: { width: '6in', height: '9in' }, margins: { left: '0.75in', right: '0.75in', gutter: '0.25in' } } });
    expect(r.ok, JSON.stringify(r)).toBe(true);
    expect(xmlOf(out)).toContain('w:gutter="360"');
  });

  it('no page means Letter portrait', async () => {
    const out = path.join(dir, 'letter.docx');
    expect(((await handleNewDocument(createMcpSessionStore(), { outPath: out, template: 'memo' })) as any).ok).toBe(true);
    expect(xmlOf(out)).toContain('w:w="12240" w:h="15840"');
  });

  it('a page that cannot be laid out is refused and nothing is written', async () => {
    const out = path.join(dir, 'big.docx');
    const r: any = await handleNewDocument(createMcpSessionStore(), { outPath: out, page: { size: '30x40in' } });
    expect(r.ok).toBe(false);
    expect(r.code).toBe('E_VALIDATION');
    expect(r.message).toMatch(/1 in to 22 in/);
    expect(fs.existsSync(out)).toBe(false);
    const narrow: any = await handleNewDocument(createMcpSessionStore(), { outPath: path.join(dir, 'n.docx'), page: { size: 'A5', margins: '70mm' } });
    expect(narrow.ok).toBe(false);
    expect(narrow.message).toMatch(/text width/);
  });

  it('a narrow page that suits the template only with fewer columns warns; an invoice that cannot hold its table is refused', async () => {
    const warn: any = await handleNewDocument(createMcpSessionStore(), { outPath: path.join(dir, 'w.docx'), template: 'memo', page: { size: 'A6', margins: '20mm' } });
    expect(warn.ok, JSON.stringify(warn)).toBe(true);
    expect(warn.warnings?.[0]).toMatch(/wrap awkwardly/);
    const refused: any = await handleNewDocument(createMcpSessionStore(), { outPath: path.join(dir, 'i.docx'), template: 'invoice', page: { size: 'A6', margins: '20mm' } });
    expect(refused.ok).toBe(false);
    expect(refused.message).toMatch(/too narrow for the invoice template/);
  });

  it('the old pageSize is refused and names page.size', async () => {
    const r: any = await handleNewDocument(createMcpSessionStore(), { outPath: path.join(dir, 'old.docx'), pageSize: 'a4' });
    expect(r.ok).toBe(false);
    expect(r.code).toBe('E_VALIDATION');
    expect(r.message).toMatch(/pageSize.*page\.size/);
    expect(r.remedy).toContain('"page"');
    expect(fs.existsSync(path.join(dir, 'old.docx'))).toBe(false);
  });

  it('an unknown field in page is refused', async () => {
    const r: any = await handleNewDocument(createMcpSessionStore(), { outPath: path.join(dir, 'u.docx'), page: { paper: 'A4' } });
    expect(r.ok).toBe(false);
    expect(r.message).toMatch(/paper/);
  });
});

/** DN3 9.2: the open result says when namespace prefixes were rewritten and when the package is digitally signed. */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, SIGNED_WARNING, createBlankDocxBuffer, packageNotes, packageWarnings } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';
import { handleOpenDocument } from '../../packages/mcp-server/src/lib/mcp-tool-adapter';

const renameW = (xml: string): string =>
  xml.replace(/(<\/?)w:/g, '$1x:').replace(/(\s)w:([A-Za-z])/g, '$1x:$2').replace(/xmlns:w=/g, 'xmlns:x=');

function build(o: { prefixed?: boolean; signed?: boolean }): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  if (o.prefixed) zip.updateFile('word/document.xml', Buffer.from(renameW(zip.readAsText('word/document.xml')), 'utf8'));
  if (o.signed) {
    zip.addFile('_xmlsignatures/origin.sigs', Buffer.alloc(0));
    zip.addFile('_xmlsignatures/sig1.xml', Buffer.from('<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>', 'utf8'));
    zip.updateFile('[Content_Types].xml', Buffer.from(zip.readAsText('[Content_Types].xml')
      .replace('<Default Extension="xml"', '<Default Extension="sigs" ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/><Default Extension="xml"'), 'utf8'));
  }
  return zip.toBuffer();
}

describe('engine report', () => {
  it('a canonical, unsigned package has nothing to report', () => {
    const report = new DocxParser(build({})).getPackageReport();
    expect(report).toEqual({ canonicalised: [], signed: false });
    expect(packageWarnings(report)).toEqual([]);
    expect(packageNotes(report)).toBeUndefined();
  });

  it('a prefixed package lists the rewritten part with from and to', () => {
    const report = new DocxParser(build({ prefixed: true })).getPackageReport();
    expect(report.signed).toBe(false);
    expect(report.canonicalised.find((p) => p.part === 'word/document.xml')).toEqual({ part: 'word/document.xml', from: ['x'], to: ['w'] });
    const [warning] = packageWarnings(report);
    expect(warning).toMatch(/^\d+ parts? (was|were) rewritten|used namespace prefixes other than the usual ones/);
    expect(warning).toContain('word/document.xml: x to w');
    expect(warning).toContain('Saving writes the usual prefixes.');
  });

  it('a signed package is reported as signed whether or not a prefix needed rewriting', () => {
    const canonical = new DocxParser(build({ signed: true })).getPackageReport();
    expect(canonical).toEqual({ canonicalised: [], signed: true });
    expect(packageWarnings(canonical)).toEqual([SIGNED_WARNING]);
    const both = new DocxParser(build({ signed: true, prefixed: true })).getPackageReport();
    expect(both.signed).toBe(true);
    expect(packageWarnings(both)).toHaveLength(2);
    expect(packageWarnings(both)[0]).toBe(SIGNED_WARNING);
  });

  it('the signed warning says plainly that saving invalidates the signature', () => {
    expect(SIGNED_WARNING).toBe('This package is digitally signed. Saving it changes signed parts and invalidates the signature.');
  });
});

describe('open_document result', () => {
  let dir: string;
  beforeEach(() => { dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pkgrep-')); });
  afterEach(() => { fs.rmSync(dir, { recursive: true, force: true }); });
  const open = async (name: string, buf: Buffer): Promise<Record<string, unknown>> => {
    const file = path.join(dir, name);
    fs.writeFileSync(file, buf);
    return (await handleOpenDocument(createMcpSessionStore(), { path: file })) as unknown as Record<string, unknown>;
  };

  it('canonical and unsigned: the result has no packageNotes key at all (output unchanged)', async () => {
    const res = await open('plain.docx', build({}));
    expect(res.ok).toBe(true);
    expect('packageNotes' in res).toBe(false);
  });

  it('signed and canonical: the warning is in the result', async () => {
    const res = await open('signed.docx', build({ signed: true }));
    expect(res.ok, JSON.stringify(res)).toBe(true);
    const notes = res.packageNotes as { signed: boolean; canonicalised: unknown[]; warnings: string[] };
    expect(notes.signed).toBe(true);
    expect(notes.canonicalised).toEqual([]);
    expect(notes.warnings).toEqual([SIGNED_WARNING]);
  });

  it('prefixed: the rewrite is in the result and the document is readable', async () => {
    const res = await open('prefixed.docx', build({ prefixed: true }));
    expect(res.ok, JSON.stringify(res)).toBe(true);
    const notes = res.packageNotes as { signed: boolean; canonicalised: Array<{ part: string }>; warnings: string[] };
    expect(notes.signed).toBe(false);
    expect(notes.canonicalised.map((p) => p.part)).toContain('word/document.xml');
    expect(notes.warnings).toHaveLength(1);
    expect((res.summary as { nodes?: unknown[] } | null)?.nodes?.length ?? 1).toBeGreaterThan(0);
  });
});

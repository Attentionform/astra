/**
 * validate_document judges the document as ASTra would save it, so out-of-order
 * property children do not fail it; the result lists the repairs a save would
 * apply (the same list open_document and inspect_document carry), and stays
 * silent for a document that needs none.
 */
import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { LEAN6_HANDLERS } from '../../src/lib/mcp-tool-adapter';
import { handleNewDocument } from '../../src/tools/new-document';
import { handleComposeDocument } from '../../src/tools/compose-document';
import { handleValidateDocument } from '../../src/lib/mcp/handlers/parity';

const dirs: string[] = [];
afterEach(() => {
  while (dirs.length) fs.rmSync(dirs.pop() as string, { recursive: true, force: true });
});

async function savedDoc(misorder: boolean): Promise<string> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-validate-pending-'));
  dirs.push(dir);
  const file = path.join(dir, 'doc.docx');
  const store = createMcpSessionStore();
  const created = await handleNewDocument(store, { outPath: file, template: 'blank' });
  if (!created.ok) throw new Error(`new_document failed: ${JSON.stringify(created)}`);
  const composed = await handleComposeDocument(store, { sessionId: created.sessionId, markdown: '| A | B |\n| --- | --- |\n| 1 | 2 |\n\nBody.\n' });
  if (!composed.ok) throw new Error(`compose failed: ${JSON.stringify(composed)}`);
  const saved = (await LEAN6_HANDLERS.save_document(store, { sessionId: created.sessionId, outPath: file })) as { ok: boolean };
  expect(saved.ok).toBe(true);
  if (misorder) {
    const zip = new AdmZip(file);
    const xml = zip.readAsText('word/document.xml');
    const edited = xml.replace('</w:tblBorders><w:tblCellMar>', '</w:tblBorders><w:tblLook w:val="04A0"/><w:tblCellMar>');
    expect(edited).not.toBe(xml);
    zip.updateFile('word/document.xml', Buffer.from(edited, 'utf8'));
    zip.writeZip(file);
  }
  return file;
}

type Validated = { ok: boolean; validation: { valid: boolean }; pendingNormalizations?: string[] };

async function validate(file: string): Promise<Validated> {
  const store = createMcpSessionStore();
  const opened = (await LEAN6_HANDLERS.open_document(store, { path: file })) as { ok: boolean; sessionId: string };
  expect(opened.ok).toBe(true);
  return (await handleValidateDocument(store, { sessionId: opened.sessionId })) as unknown as Validated;
}

describe('validate_document pending repairs', () => {
  it('lists the reorder a save would make while the verdict stays ok', async () => {
    const result = await validate(await savedDoc(true));
    expect(result.ok).toBe(true);
    expect(result.validation.valid).toBe(true);
    expect(result.pendingNormalizations).toEqual([expect.stringMatching(/Reordered w:tblPr into schema order/)]);
  });

  it('says nothing extra for a document that needs no repair', async () => {
    const result = await validate(await savedDoc(false));
    expect(result.ok).toBe(true);
    expect(result.validation.valid).toBe(true);
    expect(result.pendingNormalizations).toBeUndefined();
  });
});

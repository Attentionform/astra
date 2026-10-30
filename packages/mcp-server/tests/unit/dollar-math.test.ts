import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { AdmZip, DocxParser } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { ComposeDocumentSchema, handleComposeDocument } from '../../src/tools/compose-document';
import { ImportDocumentSchema } from '../../src/lib/mcp/schemas';
import { handleImportDocument } from '../../src/lib/mcp/handlers/parity';

const createdDirs: string[] = [];
afterEach(() => {
  for (const dir of createdDirs.splice(0)) fs.rmSync(dir, { recursive: true, force: true });
});

async function openBlankSession() {
  const store = createMcpSessionStore();
  const handle = await store.openBuffer(DocxParser.createNewDocument('blank').rezip(), 'memory.docx');
  return { store, sessionId: handle.sessionId };
}

function mathCount(store: ReturnType<typeof createMcpSessionStore>, sessionId: string): number {
  const xml = new AdmZip(store.get(sessionId).parser.rezip()).readAsText('word/document.xml');
  return (xml.match(/<m:oMath>/g) ?? []).length;
}

const SOURCE = 'We have $x^2$ here, and \\(y_1\\) there. It costs $5 and $10.\n';

describe('dollarMath option', () => {
  it('is accepted by both schemas as a boolean and rejects other types', () => {
    expect(ComposeDocumentSchema.safeParse({ sessionId: 's', markdown: 'x', dollarMath: true }).success).toBe(true);
    expect(ComposeDocumentSchema.safeParse({ sessionId: 's', markdown: 'x', dollarMath: 'yes' }).success).toBe(false);
    expect(ImportDocumentSchema.safeParse({ sessionId: 's', path: 'a.md', dollarMath: true }).success).toBe(true);
    expect(ImportDocumentSchema.safeParse({ sessionId: 's', path: 'a.md', dollarMath: 1 }).success).toBe(false);
  });

  it('compose_document reads $...$ only when asked (parenthesis math always works)', async () => {
    const off = await openBlankSession();
    expect((await handleComposeDocument(off.store, { sessionId: off.sessionId, markdown: SOURCE })).ok).toBe(true);
    expect(mathCount(off.store, off.sessionId)).toBe(1);
    const on = await openBlankSession();
    expect((await handleComposeDocument(on.store, { sessionId: on.sessionId, markdown: SOURCE, dollarMath: true })).ok).toBe(true);
    expect(mathCount(on.store, on.sessionId)).toBe(2);
  });

  it('import_document forwards the option', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-dollar-'));
    createdDirs.push(dir);
    const mdPath = path.join(dir, 'm.md');
    fs.writeFileSync(mdPath, SOURCE, 'utf8');
    const off = await openBlankSession();
    expect((await handleImportDocument(off.store, { sessionId: off.sessionId, path: mdPath })).ok).toBe(true);
    expect(mathCount(off.store, off.sessionId)).toBe(1);
    const on = await openBlankSession();
    expect((await handleImportDocument(on.store, { sessionId: on.sessionId, path: mdPath, dollarMath: true })).ok).toBe(true);
    expect(mathCount(on.store, on.sessionId)).toBe(2);
  });
});

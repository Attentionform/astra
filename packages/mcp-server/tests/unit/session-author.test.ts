/** open_document / new_document `author`: the session's default author for tracked changes and comments. */
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { createBlankDocxBuffer } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { LEAN6_HANDLERS } from '../../src/lib/mcp-tool-adapter';
import { handleNewDocument } from '../../src/tools/new-document';
import { OpenDocumentSchema } from '../../src/lib/mcp/schemas';

let dir: string;
beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-session-author-'));
  delete process.env.ASTRA_AUTHOR;
});
afterEach(() => fs.rmSync(dir, { recursive: true, force: true }));

describe('session default author', () => {
  it('new_document sets it', async () => {
    const store = createMcpSessionStore();
    const created: any = await handleNewDocument(store, { outPath: path.join(dir, 'a.docx'), template: 'blank', author: 'Session Owner' });
    expect(created.ok).toBe(true);
    const handle = store.get(created.sessionId) as any;
    expect(handle.ast.settings.defaultAuthor).toBe('Session Owner');
  });

  it('new_document without it keeps ASTra', async () => {
    const store = createMcpSessionStore();
    const created: any = await handleNewDocument(store, { outPath: path.join(dir, 'b.docx'), template: 'blank' });
    expect((store.get(created.sessionId) as any).ast.settings.defaultAuthor).toBe('ASTra');
  });

  it('open_document accepts and applies it, and rejects an empty name', async () => {
    const file = path.join(dir, 'c.docx');
    fs.writeFileSync(file, createBlankDocxBuffer('blank'));
    expect(OpenDocumentSchema.safeParse({ path: file, author: '   ' }).success).toBe(false);
    const store = createMcpSessionStore();
    const opened: any = await (LEAN6_HANDLERS as any).open_document(store, { path: file, author: 'Opener' });
    expect(opened.ok).toBe(true);
    expect((store.get(opened.sessionId) as any).ast.settings.defaultAuthor).toBe('Opener');
  });
});

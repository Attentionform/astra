import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { DocxParser } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { LEAN6_HANDLERS } from '../../src/lib/mcp-tool-adapter';
import { handleNewDocument } from '../../src/tools/new-document';

const files: string[] = [];
afterEach(() => {
  while (files.length > 0) {
    try {
      fs.unlinkSync(files.pop() as string);
    } catch {
      /* best-effort */
    }
  }
});

function newFile(template: string): string {
  const file = path.join(os.tmpdir(), `astra-cover-seed-${template}-${process.pid}-${Date.now()}-${files.length}.docx`);
  fs.writeFileSync(file, DocxParser.createNewDocument(template).rezip());
  files.push(file);
  return file;
}

describe('coverPage on a document opened from a template file keeps its content', () => {
  it('a report opened from a file is not treated as a fresh template', async () => {
    const store = createMcpSessionStore();
    const opened = await LEAN6_HANDLERS.open_document(store, { path: newFile('report') });
    if (!opened.ok) throw new Error(JSON.stringify(opened));
    const out = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId: opened.sessionId,
      operations: [{ action: 'coverPage', title: 'Annual Report' }],
    } as never);
    expect(out.ok).toBe(true);
    if (!out.ok) return;
    expect(out.results[0].removedSeedNodes).toBeUndefined();
    const inspect = await LEAN6_HANDLERS.inspect_document(store, { sessionId: opened.sessionId } as never);
    expect(JSON.stringify(inspect)).toContain('Summarise the purpose');
  });

  it('a blank file that is still one empty paragraph has nothing to keep: the cover replaces it', async () => {
    const store = createMcpSessionStore();
    const opened = await LEAN6_HANDLERS.open_document(store, { path: newFile('blank') });
    if (!opened.ok) throw new Error(JSON.stringify(opened));
    const out = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId: opened.sessionId,
      operations: [{ action: 'coverPage', title: 'Annual Report' }],
    } as never);
    expect(out.ok).toBe(true);
    if (!out.ok) return;
    expect(out.results[0].removedSeedNodes).toHaveLength(1);
  });
});

describe('coverPage on a document made by new_document', () => {
  it('blank: the empty paragraph is removed and listed in removedSeedNodes', async () => {
    const store = createMcpSessionStore();
    const outPath = path.join(os.tmpdir(), `astra-cover-new-blank-${process.pid}-${Date.now()}.docx`);
    files.push(outPath);
    const made = await handleNewDocument(store, { outPath, template: 'blank' });
    if (!made.ok) throw new Error(JSON.stringify(made));
    const out = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId: made.sessionId,
      operations: [{ action: 'coverPage', title: 'Annual Report' }],
    } as never);
    expect(out.ok).toBe(true);
    if (!out.ok) return;
    expect(out.results[0].removedSeedNodes).toHaveLength(1);
    const after = String(((await LEAN6_HANDLERS.inspect_document(store, { sessionId: made.sessionId } as never)) as { toonText?: string }).toonText);
    expect(after).toContain('Annual Report');
  });

  for (const template of ['report', 'memo']) {
    it(`${template}: a designed template keeps its content; the cover is added in front`, async () => {
      const store = createMcpSessionStore();
      const outPath = path.join(os.tmpdir(), `astra-cover-new-${template}-${process.pid}-${Date.now()}.docx`);
      files.push(outPath);
      const made = await handleNewDocument(store, { outPath, template });
      if (!made.ok) throw new Error(JSON.stringify(made));
      const out = await LEAN6_HANDLERS.execute_batch(store, {
        sessionId: made.sessionId,
        operations: [{ action: 'coverPage', title: 'Annual Report' }],
      } as never);
      expect(out.ok).toBe(true);
      if (!out.ok) return;
      expect(out.results[0].removedSeedNodes).toBeUndefined();
      const after = String(((await LEAN6_HANDLERS.inspect_document(store, { sessionId: made.sessionId } as never)) as { toonText?: string }).toonText);
      expect(after).toContain('Annual Report');
    });
  }
});

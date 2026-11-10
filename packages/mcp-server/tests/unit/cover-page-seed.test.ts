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
  it('a document opened from a file is not treated as a fresh template', async () => {
    const store = createMcpSessionStore();
    const opened = await LEAN6_HANDLERS.open_document(store, { path: newFile('blank') });
    if (!opened.ok) throw new Error(JSON.stringify(opened));
    const out = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId: opened.sessionId,
      operations: [{ action: 'coverPage', title: 'Annual Report' }],
    } as never);
    expect(out.ok).toBe(true);
    if (!out.ok) return;
    expect(out.results[0].removedSeedNodes).toBeUndefined();
    const inspect = await LEAN6_HANDLERS.inspect_document(store, { sessionId: opened.sessionId } as never);
    expect(JSON.stringify(inspect)).toContain('Blank Document');
  });
});

describe('coverPage on a document made by new_document replaces the placeholder text', () => {
  for (const template of ['blank', 'report', 'memo']) {
    it(`${template}: the placeholder nodes are removed and listed in removedSeedNodes`, async () => {
      const store = createMcpSessionStore();
      const outPath = path.join(os.tmpdir(), `astra-cover-new-${template}-${process.pid}-${Date.now()}.docx`);
      files.push(outPath);
      const made = await handleNewDocument(store, { outPath, template });
      if (!made.ok) throw new Error(JSON.stringify(made));
      const before = await LEAN6_HANDLERS.inspect_document(store, { sessionId: made.sessionId } as never);
      const seedCount = (before as { rowCount: number }).rowCount;
      const out = await LEAN6_HANDLERS.execute_batch(store, {
        sessionId: made.sessionId,
        operations: [{ action: 'coverPage', title: 'Annual Report' }],
      } as never);
      expect(out.ok).toBe(true);
      if (!out.ok) return;
      expect(out.results[0].removedSeedNodes).toHaveLength(seedCount);
      const after = String(((await LEAN6_HANDLERS.inspect_document(store, { sessionId: made.sessionId } as never)) as { toonText?: string }).toonText);
      expect(after).toContain('Annual Report');
      expect(after).not.toContain('Blank Document');
      expect(after).not.toContain('Q3 Executive');
    });
  }
});

import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { createBlankDocxBuffer } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { LEAN6_HANDLERS } from '../../src/lib/mcp-tool-adapter';
import { handleExportDocument } from '../../src/tools/export';

const tmpDirs: string[] = [];

async function makeSession(
  store: ReturnType<typeof createMcpSessionStore>
): Promise<{ sessionId: string; file: string }> {
  const dir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'astra-export-pin-'));
  tmpDirs.push(dir);
  const file = path.join(dir, 'doc.docx');
  await fs.promises.writeFile(file, createBlankDocxBuffer('blank'));
  const opened = await LEAN6_HANDLERS.open_document(store, { path: file });
  expect(opened.ok).toBe(true);
  if (!opened.ok) throw new Error('open failed in test setup');
  return { sessionId: opened.sessionId, file };
}

async function reopenSession(
  store: ReturnType<typeof createMcpSessionStore>,
  file: string
): Promise<string> {
  const opened = await LEAN6_HANDLERS.open_document(store, { path: file });
  expect(opened.ok).toBe(true);
  if (!opened.ok) throw new Error('reopen failed in test setup');
  return opened.sessionId;
}

async function firstNodeId(
  store: ReturnType<typeof createMcpSessionStore>,
  sessionId: string
): Promise<string> {
  const first = await LEAN6_HANDLERS.inspect_document(store, { sessionId, offset: 0, limit: 1 });
  expect(first.ok).toBe(true);
  if (!first.ok) throw new Error('inspect failed in test setup');
  const id = first.toonText
    .split('\n')
    .map((l) => l.split(/\s+/)[0])
    .filter((v) => !!v && !['>', 'ID'].includes(v) && !/^---/.test(v))[0];
  expect(typeof id).toBe('string');
  return id;
}

afterEach(async () => {
  while (tmpDirs.length > 0) {
    const d = tmpDirs.pop() as string;
    try {
      await fs.promises.rm(d, { recursive: true, force: true });
    } catch {
      /* best-effort */
    }
  }
});

describe('MCP HTML export fidelity pins', () => {
  it('HTML export of a footnoted document carries the real footnote body', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store);
    const anchor = await firstNodeId(store, sessionId);
    const body = 'Vindication lineage parcel 7XQ.';
    const batch = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [{ action: 'footnote', anchor, text: body }],
    });
    expect(batch.ok).toBe(true);

    const exported = await handleExportDocument(store, { sessionId, format: 'html' });
    expect(exported.ok).toBe(true);
    if (!exported.ok) return;
    expect(exported.content).toContain(body);
    expect(exported.content).not.toContain('see document');
  });

  it('HTML export renders callouts as div.callout, never a data table', async () => {
    const store = createMcpSessionStore();
    const { sessionId } = await makeSession(store);
    const anchor = await firstNodeId(store, sessionId);
    const batch = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [
        { action: 'callout', anchor, position: 'after', title: 'Field Memo', text: 'Keep the north gate shut.', theme: 'indigo' },
      ],
    });
    expect(batch.ok).toBe(true);

    const exported = await handleExportDocument(store, { sessionId, format: 'html' });
    expect(exported.ok).toBe(true);
    if (!exported.ok) return;
    expect(exported.content).toContain('<div class="callout"><h4>Field Memo</h4><p>Keep the north gate shut.</p></div>');
    expect(exported.content).not.toContain('<table class="doc-table">');
  });

  it('HTML export renders a persisted/reopened callout as div.callout', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store);
    const anchor = await firstNodeId(store, sessionId);
    const batch = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [
        { action: 'callout', anchor, position: 'after', title: 'North Gate', text: 'Keep the north gate shut.', theme: 'amber' },
      ],
    });
    expect(batch.ok).toBe(true);
    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId });
    expect(saved.ok).toBe(true);
    const closed = await LEAN6_HANDLERS.close_document(store, { sessionId });
    expect(closed.ok).toBe(true);

    const reopened = await reopenSession(store, file);
    const exported = await handleExportDocument(store, { sessionId: reopened, format: 'html' });
    expect(exported.ok).toBe(true);
    if (!exported.ok) return;
    expect(exported.content).toContain('<div class="callout"><h4>North Gate</h4><p>Keep the north gate shut.</p></div>');
    expect(exported.content).not.toContain('<table class="doc-table">');
  });

  it('HTML export keeps a persisted/reopened plain data table as a table', async () => {
    const store = createMcpSessionStore();
    const { sessionId, file } = await makeSession(store);
    const anchor = await firstNodeId(store, sessionId);
    const batch = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [
        { action: 'insert', type: 'table', anchor, position: 'after', headers: ['A', 'B'], rows: [['1', '2']] },
      ],
    });
    expect(batch.ok).toBe(true);
    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId });
    expect(saved.ok).toBe(true);
    const closed = await LEAN6_HANDLERS.close_document(store, { sessionId });
    expect(closed.ok).toBe(true);

    const reopened = await reopenSession(store, file);
    const exported = await handleExportDocument(store, { sessionId: reopened, format: 'html' });
    expect(exported.ok).toBe(true);
    if (!exported.ok) return;
    expect(exported.content).toContain('<table class="doc-table">');
    expect(exported.content).not.toContain('<div class="callout">');
  });
});

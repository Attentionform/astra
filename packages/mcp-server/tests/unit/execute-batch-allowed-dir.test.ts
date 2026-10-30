import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { DocxParser, IMAGE_FILE_MAX_BYTES } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { LEAN6_HANDLERS } from '../../src/lib/mcp-tool-adapter';

const TINY_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

function imageDir(): { dir: string; session: { store: ReturnType<typeof createMcpSessionStore>; sessionId: string } | null } {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-mcp-batch-'));
  fs.writeFileSync(path.join(dir, 'tiny.png'), Buffer.from(TINY_PNG_BASE64, 'base64'));
  return { dir, session: null };
}

async function openDocIn(dir: string) {
  const store = createMcpSessionStore();
  const buf = DocxParser.createNewDocument('blank').rezip();
  const docPath = path.join(dir, 'doc.docx');
  fs.writeFileSync(docPath, buf);
  const handle = await store.openDocument(docPath);
  return { store, sessionId: handle.sessionId };
}

function imageOp(anchor: string, extra?: Record<string, unknown>) {
  return { action: 'image', operation: 'insert', anchor, filePath: 'tiny.png', ...(extra ?? {}) };
}

describe('execute_batch allowedDir jail', () => {
  it('inserts by path via the session anchor (no explicit allowedDir)', async () => {
    const { dir } = imageDir();
    const { store, sessionId } = await openDocIn(dir);
    const inspect = await LEAN6_HANDLERS.inspect_document(store, { sessionId, fidelity: 'compact' });
    expect(inspect.ok).toBe(true);
    const anchor = (inspect as { toonText: string }).toonText.match(/([0-9A-F]{8})/)?.[1] ?? 'end';
    const res = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [imageOp(anchor === 'end' ? 'end' : anchor, { altText: 'Tiny dot' })],
    });
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error(`batch failed: ${JSON.stringify(res)}`);
    expect(res.appliedCount).toBe(1);
  });

  it('inserts by path via an explicit allowedDir', async () => {
    const { dir } = imageDir();
    const other = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-mcp-batch-'));
    const { store, sessionId } = await openDocIn(other);
    const res = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [{ action: 'image', operation: 'insert', anchor: 'end', filePath: 'tiny.png', altText: 'Tiny dot' }],
      allowedDir: dir,
    });
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error(`batch failed: ${JSON.stringify(res)}`);
    expect(res.appliedCount).toBe(1);
  });

  it('refuses an explicit allowedDir outside the roots with E_SANDBOX', async () => {
    const { dir } = imageDir();
    const { store, sessionId } = await openDocIn(dir);
    const res = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [{ action: 'image', operation: 'insert', anchor: 'end', filePath: 'tiny.png' }],
      allowedDir: '/definitely-not-an-allowed-root-xyz',
    });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('expected refusal');
    expect(res.code).toBe('E_SANDBOX');
  });

  it('fails closed with E_SANDBOX when the session has no filesystem anchor', async () => {
    const store = createMcpSessionStore();
    const buf = DocxParser.createNewDocument('blank').rezip();
    const handle = await store.openBuffer(buf, 'memory.docx');
    const res = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId: handle.sessionId,
      operations: [{ action: 'image', operation: 'insert', anchor: 'end', filePath: 'tiny.png' }],
    });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('expected refusal');
    expect(res.code).toBe('E_SANDBOX');
  });

  it('maps outside/missing/oversize filePath refs to E_SANDBOX/E_IO/E_LIMIT', async () => {
    const { dir } = imageDir();
    const big = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-mcp-batch-'));
    fs.writeFileSync(path.join(big, 'big.png'), Buffer.alloc(IMAGE_FILE_MAX_BYTES + 1));
    const { store, sessionId } = await openDocIn(dir);
    const cases: Array<[Record<string, unknown> | undefined, string, string]> = [
      [undefined, '../escape.png', 'E_SANDBOX'],
      [undefined, 'nope.png', 'E_IO'],
      [{ allowedDir: big } as unknown as Record<string, unknown>, 'big.png', 'E_LIMIT'],
    ];
    for (const [extra, filePath, code] of cases) {
      const res = await LEAN6_HANDLERS.execute_batch(store, {
        sessionId,
        operations: [{ action: 'image', operation: 'insert', anchor: 'end', filePath }],
        ...(extra ?? {}),
      });
      expect(res.ok).toBe(false);
      if (res.ok) throw new Error(`expected refusal for ${filePath}`);
      expect(res.code).toBe(code);
    }
  });

  it('leaves batches without filePath ops untouched (explicit allowedDir inert)', async () => {
    const { dir } = imageDir();
    const { store, sessionId } = await openDocIn(dir);
    const res = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'plain' }],
      allowedDir: '/definitely-not-an-allowed-root-xyz',
    });
    expect(res.ok).toBe(true);
  });
});

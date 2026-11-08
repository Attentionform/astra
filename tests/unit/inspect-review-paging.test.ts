import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { createBlankDocxBuffer } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';
import { handleInspectDocument } from '../../packages/mcp-server/src/lib/mcp-tool-adapter';

async function open(paragraphs: number, revisions: number) {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  const body = Array.from({ length: paragraphs }, (_, i) => `<w:p><w:r><w:t>para${String(i).padStart(3, '0')}</w:t></w:r></w:p>`).join('')
    + `<w:p>${Array.from({ length: revisions }, (_, i) => `<w:ins w:id="${i + 1}" w:author="A"><w:r><w:t>rev${i}</w:t></w:r></w:ins>`).join('')}</w:p>`;
  zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`)));
  const store = createMcpSessionStore();
  const handle = await store.openBuffer(zip.toBuffer(), 'paging.docx');
  return { store, sessionId: handle.sessionId };
}
const records = (out: unknown) => ((out as { revisions: { revisions: { id: number }[] } }).revisions.revisions);

describe('FX-13: review records and nodes page independently', () => {
  it('reviewOffset pages revisions and leaves the node window where offset puts it', async () => {
    const { store, sessionId } = await open(120, 55);
    const nodes = await handleInspectDocument(store, { sessionId, include: ['revisions'], offset: 10, limit: 5 }) as Record<string, any>;
    expect(nodes.ok).toBe(true);
    expect(nodes.offset).toBe(10);
    expect(nodes.toonText).toContain('para010');
    expect(nodes.toonText).not.toContain('para050');
    expect(records(nodes)).toHaveLength(50); // default review window, not moved by offset
    expect(nodes.revisions.nextOffset).toBe(50);

    const second = await handleInspectDocument(store, { sessionId, include: ['revisions'], reviewOffset: 50, offset: 10, limit: 5 }) as Record<string, any>;
    expect(second.ok).toBe(true);
    expect(records(second)).toHaveLength(5);
    expect(second.revisions.total).toBe(55);
    expect(second.revisions.nextOffset).toBeUndefined();
    expect(second.offset).toBe(10);
    expect(second.toonText).toContain('para010');
  });

  it('a 1-paragraph document pages 55 revisions with reviewOffset', async () => {
    const { store, sessionId } = await open(1, 55);
    const first = await handleInspectDocument(store, { sessionId, include: ['revisions'] }) as Record<string, any>;
    expect(first).toMatchObject({ ok: true, revisions: { total: 55, nextOffset: 50 } });
    const second = await handleInspectDocument(store, { sessionId, include: ['revisions'], reviewOffset: 50 }) as Record<string, any>;
    expect(records(second)).toHaveLength(5);
  });

  it('reviewLimit narrows the page and is clamped to 50', async () => {
    const { store, sessionId } = await open(1, 55);
    const small = await handleInspectDocument(store, { sessionId, include: ['revisions'], reviewLimit: 7 }) as Record<string, any>;
    expect(records(small)).toHaveLength(7);
    expect(small.revisions.nextOffset).toBe(7);
    const big = await handleInspectDocument(store, { sessionId, include: ['revisions'], reviewLimit: 200 }) as Record<string, any>;
    expect(records(big)).toHaveLength(50);
  });

  it('offset and limit do not move the review window', async () => {
    const { store, sessionId } = await open(1, 55);
    const out = await handleInspectDocument(store, { sessionId, include: ['revisions'], offset: 0, limit: 3 }) as Record<string, any>;
    expect(records(out)[0].id).toBe(1);
    expect(records(out)).toHaveLength(50);
  });
});

describe('FX-13: guards', () => {
  it('a node offset past the end is rejected even when include is set', async () => {
    const { store, sessionId } = await open(1, 3);
    for (const include of [undefined, ['revisions']]) {
      const out = await handleInspectDocument(store, { sessionId, offset: 500, ...(include ? { include } : {}) }) as Record<string, any>;
      expect(out.ok).toBe(false);
      expect(JSON.stringify(out)).toMatch(/offset 500 is beyond totalNodes/);
    }
  });

  it('revisionId without include revisions is rejected, with it is accepted', async () => {
    const { store, sessionId } = await open(1, 3);
    const bad = await handleInspectDocument(store, { sessionId, revisionId: 3 }) as Record<string, any>;
    expect(bad.ok).toBe(false);
    expect(JSON.stringify(bad)).toMatch(/revisionId needs include/);
    const wrong = await handleInspectDocument(store, { sessionId, revisionId: 3, include: ['comments'] }) as Record<string, any>;
    expect(wrong.ok).toBe(false);
    const good = await handleInspectDocument(store, { sessionId, revisionId: 3, include: ['revisions'] }) as Record<string, any>;
    expect(good.ok).toBe(true);
  });

  it('commentId without include comments is rejected', async () => {
    const { store, sessionId } = await open(1, 1);
    const bad = await handleInspectDocument(store, { sessionId, commentId: 0 }) as Record<string, any>;
    expect(bad.ok).toBe(false);
    expect(JSON.stringify(bad)).toMatch(/commentId needs include/);
  });

  it('reviewOffset and reviewLimit out of range are rejected by the schema', async () => {
    const { store, sessionId } = await open(1, 1);
    expect(((await handleInspectDocument(store, { sessionId, include: ['revisions'], reviewOffset: -1 })) as { ok: boolean }).ok).toBe(false);
    expect(((await handleInspectDocument(store, { sessionId, include: ['revisions'], reviewLimit: 0 })) as { ok: boolean }).ok).toBe(false);
  });
});

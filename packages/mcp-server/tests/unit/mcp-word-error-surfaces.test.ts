/**
 * Word-level span discoverability + error-envelope pins: execute_batch docs
 * name the match span capability; a match miss keeps its inspect-and-retry
 * remedy through the MCP envelope (never the stale node-id hint), while a
 * genuine lowercase anchor miss keeps the stale-id hint (exclusion
 * semantics: only the "not found in node" word-miss shape is excluded).
 */
import { describe, it, expect } from 'vitest';
import { DocxParser } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { LEAN6_DESCRIPTIONS, LEAN6_HANDLERS } from '../../src/lib/mcp-tool-adapter';

function blankBuf(): Buffer {
  return DocxParser.createNewDocument('blank').rezip();
}

async function firstNodeId(store: ReturnType<typeof createMcpSessionStore>, sessionId: string): Promise<string> {
  const first = await LEAN6_HANDLERS.inspect_document(store, { sessionId, offset: 0, limit: 1 });
  expect(first.ok).toBe(true);
  if (!first.ok) throw new Error('inspect failed in test setup');
  const toon = String((first as unknown as Record<string, unknown>).toonText ?? '');
  const id = toon
    .split('\n')
    .map((l) => l.split(/\s+/)[0])
    .filter((t) => !!t && !['>', 'ID'].includes(t) && !/^---/.test(t))[0];
  expect(typeof id).toBe('string');
  return id;
}

describe('execute_batch docs name word-level span styling', () => {
  it('description mentions match spans', () => {
    expect(LEAN6_DESCRIPTIONS.execute_batch).toContain('word-level span');
    expect(LEAN6_DESCRIPTIONS.execute_batch).toContain('matchAll');
    expect(LEAN6_DESCRIPTIONS.execute_batch).toContain('matchCase');
    expect(LEAN6_DESCRIPTIONS.execute_batch).toContain('E_NOT_FOUND');
  });
});

describe('match miss keeps its remedy through the MCP envelope', () => {
  it('word miss remedy survives, stale-id hint does not fire', async () => {
    const store = createMcpSessionStore();
    const h = await store.openBuffer(blankBuf(), 'word-miss.docx');
    const target = await firstNodeId(store, h.sessionId);
    const bad = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId: h.sessionId,
      operations: [{ action: 'style', target, match: 'zz-absent-needle', bold: true }],
    });
    expect(bad.ok).toBe(false);
    if (!bad.ok) {
      expect(bad.code).toBe('E_NOT_FOUND');
      expect(String(bad.remedy)).toContain('exact matching substring');
      expect(String(bad.remedy)).not.toContain('go stale');
    }
    await LEAN6_HANDLERS.close_document(store, { sessionId: h.sessionId });
  });

  it('lowercase anchor miss keeps the stale-id hint through the MCP envelope', async () => {
    const store = createMcpSessionStore();
    const h = await store.openBuffer(blankBuf(), 'anchor-miss.docx');
    const bad = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId: h.sessionId,
      operations: [{ action: 'insert', type: 'paragraph', anchor: 'p_no_such_node', position: 'after', text: 'x' }],
    });
    expect(bad.ok).toBe(false);
    if (!bad.ok) {
      expect(String(bad.remedy)).toContain('go stale');
    }
    await LEAN6_HANDLERS.close_document(store, { sessionId: h.sessionId });
  });
});

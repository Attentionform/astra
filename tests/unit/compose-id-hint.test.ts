import { describe, it, expect } from 'vitest';
import { COMPOSE_ID_PATTERN, createBlankDocxBuffer } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';
import { handleExecuteBatch } from '../../packages/mcp-server/src/lib/mcp-tool-adapter';
import { handleComposeDocument } from '../../packages/mcp-server/src/tools/compose-document';

/**
 * A missing id that compose_document could have minted gets a hint that names the ways out and does not claim
 * a cause. The hint reads the id from the failure, so it never hides a "not found" about another node.
 */
describe('compose id pattern', () => {
  it('matches the ids compose mints, including sequences of five digits and more', () => {
    for (const id of ['md_0001', 'md_0002', 'md2_0001', 'md3_0123', 'md_9999', 'md_10000', 'md_123456']) {
      expect(COMPOSE_ID_PATTERN.test(id), id).toBe(true);
    }
  });

  it('does not match other ids', () => {
    for (const id of ['md_abc', 'mdx_0001', 'p_001', 't_001', 'md_12', 'md_', '_0001', 'xmd_0001']) {
      expect(COMPOSE_ID_PATTERN.test(id), id).toBe(false);
    }
  });
});

async function composedSession() {
  const store = createMcpSessionStore();
  const handle = await store.openBuffer(createBlankDocxBuffer('blank'), 'compose-hint.docx');
  const composed = await handleComposeDocument(store, {
    sessionId: handle.sessionId,
    markdown: 'First paragraph.\n\nSecond paragraph.\n',
  });
  if (!composed.ok) throw new Error(`compose failed: ${JSON.stringify(composed)}`);
  return { store, handle, composed };
}

function failureText(out: unknown): string {
  return JSON.stringify(out);
}

describe('hint for a missing compose id', () => {
  it('a compose id that is not in the document gets the hint, worded as a possibility', async () => {
    const { store, handle } = await composedSession();
    const out = await handleExecuteBatch(store, {
      sessionId: handle.sessionId,
      operations: [{ action: 'replace', target: 'md_0099', text: 'x' }],
    });
    expect(out.ok).toBe(false);
    const text = failureText(out);
    expect(text).toContain('looks like an id from compose_document');
    expect(text).toContain('It may be');
    expect(text).toContain('persistentIds');
    expect(text).toContain('search_content');
  });

  it('a five-digit compose id gets the hint', async () => {
    const { store, handle } = await composedSession();
    const out = await handleExecuteBatch(store, {
      sessionId: handle.sessionId,
      operations: [{ action: 'replace', target: 'md_10000', text: 'x' }],
    });
    expect(failureText(out)).toContain('looks like an id from compose_document');
  });

  it('another missing id keeps the generic hint', async () => {
    const { store, handle } = await composedSession();
    const out = await handleExecuteBatch(store, {
      sessionId: handle.sessionId,
      operations: [{ action: 'replace', target: 'p_999', text: 'x' }],
    });
    const text = failureText(out);
    expect(text).not.toContain('looks like an id from compose_document');
    expect(text).toContain('session-scoped and go stale after reopen/edit');
  });

  it('a failure about a table id is not masked, even when the batch also uses an md id that exists', async () => {
    const { store, handle } = await composedSession();
    const out = await handleExecuteBatch(store, {
      sessionId: handle.sessionId,
      operations: [
        { action: 'replace', target: 'md_0001', text: 'Still here.' },
        { action: 'table', target: 't_009', operation: 'updateCell', rowIndex: 1, colIndex: 1, text: 'x' },
      ],
    });
    expect(out.ok).toBe(false);
    expect(failureText(out)).not.toContain('looks like an id from compose_document');
  });

  it('after the document is reopened, the old compose id fails with the hint and the persistent id works', async () => {
    const { store, handle, composed } = await composedSession();
    const ids = (composed as { persistentIds?: Record<string, string> }).persistentIds ?? {};
    const persistent = ids['md_0002'];
    expect(persistent, 'compose response must map md_0002').toBeTruthy();

    // A reopen: the saved bytes opened as a new session. Compose ids belong to the first session.
    const reopened = await store.openBuffer(handle.parser.rezip(), 'reopened.docx');
    const stale = await handleExecuteBatch(store, {
      sessionId: reopened.sessionId,
      operations: [{ action: 'replace', target: 'md_0002', text: 'Changed.' }],
    });
    expect(stale.ok).toBe(false);
    expect(failureText(stale)).toContain('persistentIds');

    const advised = await handleExecuteBatch(store, {
      sessionId: reopened.sessionId,
      operations: [{ action: 'replace', target: persistent, text: 'Changed.' }],
    });
    expect(advised.ok).toBe(true);
  });

  it('a deleted compose id gets the hint, which lists deletion as one possible cause', async () => {
    const { store, handle } = await composedSession();
    const deleted = await handleExecuteBatch(store, {
      sessionId: handle.sessionId,
      operations: [{ action: 'delete', target: 'md_0002' }],
    });
    expect(deleted.ok).toBe(true);
    const again = await handleExecuteBatch(store, {
      sessionId: handle.sessionId,
      operations: [{ action: 'replace', target: 'md_0002', text: 'x' }],
    });
    expect(again.ok).toBe(false);
    expect(failureText(again)).toContain('a node that was deleted');
  });

  it('control: a compose id that exists in the same session still works after the compose', async () => {
    const { store, handle } = await composedSession();
    const out = await handleExecuteBatch(store, {
      sessionId: handle.sessionId,
      operations: [{ action: 'replace', target: 'md_0002', text: 'Edited in place.' }],
    });
    expect(out.ok).toBe(true);
  });
});

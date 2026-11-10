import { describe, it, expect, vi } from 'vitest';
import { DocxParser, OperationEngine } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { handleComposeDocument } from '../../src/tools/compose-document';
import { composeDocument } from '../../../engine/src/compose/markdown';

async function openBlank(): Promise<{ store: ReturnType<typeof createMcpSessionStore>; sessionId: string }> {
  const store = createMcpSessionStore();
  const handle = await store.openBuffer(DocxParser.createNewDocument('blank').rezip(), 'memory.docx');
  return { store, sessionId: handle.sessionId };
}

const SEED_TEXTS = new Set(['Blank Document', 'Start typing your content here...']);

/** Text of every body node in document order, without the two placeholder paragraphs of the blank template. */
function bodyTexts(store: ReturnType<typeof createMcpSessionStore>, sessionId: string): string[] {
  const out: string[] = [];
  for (const section of store.get(sessionId).ast.sections) {
    for (const node of section.nodes) {
      const text = String(node.metadata?.text ?? '').trim();
      if (!SEED_TEXTS.has(text)) out.push(text);
    }
  }
  return out;
}

describe('compose_document more than once per session', () => {
  it('keeps md_ ids for the first compose and gives later composes their own prefix', async () => {
    const { store, sessionId } = await openBlank();
    const first = await handleComposeDocument(store, { sessionId, markdown: '# One\n\nFirst body.\n' });
    if (!first.ok) throw new Error(JSON.stringify(first));
    expect(first.createdNodeIds.every((id) => id.startsWith('md_'))).toBe(true);

    const second = await handleComposeDocument(store, { sessionId, markdown: '## Two\n\nSecond body.\n' });
    if (!second.ok) throw new Error(JSON.stringify(second));
    expect(second.createdNodeIds.length).toBeGreaterThan(0);
    expect(second.createdNodeIds.every((id) => id.startsWith('md2_'))).toBe(true);

    const third = await handleComposeDocument(store, { sessionId, markdown: 'Third body.\n' });
    if (!third.ok) throw new Error(JSON.stringify(third));
    expect(third.createdNodeIds.every((id) => id.startsWith('md3_'))).toBe(true);

    expect(bodyTexts(store, sessionId)).toEqual(['One', 'First body.', 'Two', 'Second body.', 'Third body.']);
    // persistentIds is per call: only the ids this call created.
    for (const id of Object.keys(second.persistentIds ?? {})) expect(id.startsWith('md2_')).toBe(true);
    expect(second.validation.valid).toBe(true);
  });

  it('places a later compose after a node, before a node, at start and at end', async () => {
    const { store, sessionId } = await openBlank();
    const first = await handleComposeDocument(store, { sessionId, markdown: '# Head\n\nAlpha.\n\nOmega.\n' });
    if (!first.ok) throw new Error(JSON.stringify(first));
    const [head, alpha] = first.createdNodeIds;

    const mid = await handleComposeDocument(store, { sessionId, markdown: 'Mid one.\n\nMid two.\n', anchor: alpha, position: 'after' });
    expect(mid.ok).toBe(true);
    const start = await handleComposeDocument(store, { sessionId, markdown: 'Very first.\n', anchor: 'start', position: 'before' });
    expect(start.ok).toBe(true);
    const before = await handleComposeDocument(store, { sessionId, markdown: 'Before head.\n', anchor: head, position: 'before' });
    expect(before.ok).toBe(true);
    const end = await handleComposeDocument(store, { sessionId, markdown: 'The end.\n', anchor: 'end' });
    expect(end.ok).toBe(true);

    expect(bodyTexts(store, sessionId)).toEqual([
      'Very first.', 'Before head.', 'Head', 'Alpha.', 'Mid one.', 'Mid two.', 'Omega.', 'The end.',
    ]);
  });

  it('keeps the order of a block that follows an alert when it goes in the middle', async () => {
    const { store, sessionId } = await openBlank();
    const first = await handleComposeDocument(store, { sessionId, markdown: 'Top.\n\nBottom.\n' });
    if (!first.ok) throw new Error(JSON.stringify(first));
    const top = first.createdNodeIds[0];
    const mid = await handleComposeDocument(store, { sessionId, markdown: 'Before note.\n\n> [!NOTE]\n> Inside note.\n\nAfter note.\n', anchor: top });
    if (!mid.ok) throw new Error(JSON.stringify(mid));
    expect(bodyTexts(store, sessionId)).toEqual(['Top.', 'Before note.', 'Inside note.', 'After note.', 'Bottom.']);
  });

  it('refuses template after the first compose, a position without an anchor, and an unknown anchor', async () => {
    const { store, sessionId } = await openBlank();
    // A template check on the first compose of a pristine blank still works.
    const first = await handleComposeDocument(store, { sessionId, markdown: '# One\n', template: 'blank' });
    expect(first.ok).toBe(true);
    const again = await handleComposeDocument(store, { sessionId, markdown: 'More.\n', template: 'blank' });
    expect(again.ok).toBe(false);
    if (again.ok) throw new Error('unexpected');
    expect(again.message).toMatch(/template is accepted only on the first compose/);

    const noAnchor = await handleComposeDocument(store, { sessionId, markdown: 'More.\n', position: 'before' });
    expect(noAnchor.ok).toBe(false);
    if (noAnchor.ok) throw new Error('unexpected');
    expect(noAnchor.message).toMatch(/position needs an anchor/);

    const unknown = await handleComposeDocument(store, { sessionId, markdown: 'More.\n', anchor: 'p_nope' });
    expect(unknown.ok).toBe(false);
    if (unknown.ok) throw new Error('unexpected');
    expect(unknown.message).toMatch(/not a node of this document/);
  });

  it('leaves the document unchanged when a later compose fails', async () => {
    const { store, sessionId } = await openBlank();
    const first = await handleComposeDocument(store, { sessionId, markdown: '# One\n\nBody.\n' });
    expect(first.ok).toBe(true);
    const before = bodyTexts(store, sessionId);
    const revisionBefore = store.get(sessionId).revision;
    const bad = await handleComposeDocument(store, { sessionId, markdown: 'Fine.\n\n> | a |\n> | - |\n> | 1 |\n' });
    expect(bad.ok).toBe(false);
    expect(bodyTexts(store, sessionId)).toEqual(before);
    expect(store.get(sessionId).revision).toBe(revisionBefore);
  });
});

describe('a failed chunk rewinds the chunks before it', () => {
  it('restores the document when the second chunk fails', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const idsBefore = (await parser.getAST()).sections.flatMap((s) => s.nodes.map((n) => n.id));
    const real = OperationEngine.prototype.executeBatch;
    let calls = 0;
    const spy = vi.spyOn(OperationEngine.prototype, 'executeBatch').mockImplementation(async function (this: OperationEngine, batch) {
      calls += 1;
      if (calls === 2) {
        return {
          success: false,
          appliedCount: 0,
          errors: [{ operationIndex: 0, error: 'forced failure in the second chunk' }],
          rollbackTriggered: true,
          skippedIds: [],
          warnings: [],
          revisionsPending: 0,
          possiblyStaleBodyRefs: [],
        } as never;
      }
      return real.call(this, batch);
    });
    try {
      const result = await composeDocument(parser, '# One\n\nTwo.\n\nThree.\n\nFour.\n', { chunkSize: 2 });
      expect(calls).toBe(2);
      expect(result.success).toBe(false);
      expect(result.appliedCount).toBe(0);
    } finally {
      spy.mockRestore();
    }
    const idsAfter = (await parser.getAST()).sections.flatMap((s) => s.nodes.map((n) => n.id));
    expect(idsAfter).toEqual(idsBefore);
  });
});

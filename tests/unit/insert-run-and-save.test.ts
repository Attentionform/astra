/**
 * A run of consecutive inserts shares one scan for used paragraph ids, and a save writes the main
 * document once unless a revision date had to change. Both must leave the output exactly as it was.
 */
import { describe, it, expect, vi } from 'vitest';
import crypto from 'node:crypto';
import { DocxParser, OperationEngine } from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

type Op = Record<string, unknown>;

/** First 16 hex digits of the SHA-256 of the document XML. */
const digest = (xml: string): string => crypto.createHash('sha256').update(xml).digest('hex').slice(0, 16);

// Digests of the document XML written by the unpatched v0.1.0 engine for the same operations, so these
// tests fail if the shared scan ever changes an id or the order of anything.
const V010_MIXED = 'a3c82046ad1d2d4f';
const V010_AFTER_ROLLBACK = 'fcaf85b5cbaf75ff';
const V010_IDENTICAL_INSERTS = '8e37997bb5313bea';

async function session(paragraphs: number) {
  const parser = DocxParser.createNewDocument('blank');
  const ast = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  const first = ast.sections.flatMap((s) => s.nodes).find((n) => n.type === 'paragraph');
  if (!first) throw new Error('no paragraph in the blank document');
  const ops: Op[] = [];
  for (let i = 0; i < paragraphs; i++) {
    ops.push({ action: 'insert', type: 'paragraph', anchor: first.id, position: 'after', text: `base ${i}`, id: `base_${i}` });
  }
  await engine.executeBatch({ operations: ops as never });
  return { parser, engine, anchor: first.id };
}

function mixed(anchor: string): Op[] {
  return [
    { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'same text', id: 'n1' },
    { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'another', id: 'n2' },
    { action: 'delete', target: 'n1' },
    { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'same text', id: 'n1' },
    { action: 'replace', target: 'n2', text: 'changed' },
    { action: 'insert', type: 'paragraph', anchor, position: 'before', text: 'last', id: 'n3' },
    { action: 'insert', type: 'paragraph', anchor, position: 'before', text: 'last', id: 'n4' },
  ];
}

describe('run of inserts', () => {
  it('gives the document the v0.1.0 engine gave, with a delete and a re-insert of the same text in between', async () => {
    const s = await session(30);
    await s.engine.executeBatch({ operations: mixed(s.anchor) as never });
    expect(digest(s.parser.readDocumentXml())).toBe(V010_MIXED);
  });

  it('gives identical text inserted three times different paragraph ids, as v0.1.0 did', async () => {
    const s = await session(5);
    const ops: Op[] = ['x', 'y', 'z'].map((id) => ({ action: 'insert', type: 'paragraph', anchor: s.anchor, position: 'after', text: 'same', id }));
    await s.engine.executeBatch({ operations: ops as never });
    const xml = s.parser.readDocumentXml();
    const ids = [...xml.matchAll(/w14:paraId="([0-9A-F]{8})"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
    expect(digest(xml)).toBe(V010_IDENTICAL_INSERTS);
  });

  it('after a rolled-back batch the next inserts give what v0.1.0 gave', async () => {
    const s = await session(10);
    const failing: Op[] = [
      { action: 'insert', type: 'paragraph', anchor: s.anchor, position: 'after', text: 'doomed', id: 'd1' },
      { action: 'replace', target: 'no-such-node', text: 'boom' },
    ];
    await s.engine.executeBatch({ operations: failing as never }).catch(() => undefined);
    await s.engine.executeBatch({ operations: mixed(s.anchor) as never });
    expect(digest(s.parser.readDocumentXml())).toBe(V010_AFTER_ROLLBACK);
  });

  it('scans the paragraphs once for a run of inserts', async () => {
    const s = await session(40);
    const dom = s.parser.getDocumentDom();
    const original = dom.getElementsByTagName.bind(dom);
    let paragraphScans = 0;
    dom.getElementsByTagName = ((name: string) => {
      if (name === 'w:p') paragraphScans++;
      return original(name);
    }) as typeof dom.getElementsByTagName;
    const ops: Op[] = Array.from({ length: 20 }, (_, i) => ({ action: 'insert', type: 'paragraph', anchor: s.anchor, position: 'after', text: `run ${i}`, id: `run_${i}` }));
    await s.engine.executeBatch({ operations: ops as never });
    expect(paragraphScans).toBeLessThan(20);
  });
});

describe('save', () => {
  it('writes the main document once when no revision date is pinned', async () => {
    const s = await session(3);
    const spy = vi.spyOn(s.parser, 'readDocumentXml');
    s.parser.rezip({ reproducible: false });
    expect(spy).toHaveBeenCalledTimes(1);
  });
});

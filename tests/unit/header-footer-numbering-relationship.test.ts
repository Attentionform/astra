import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import {
  DocxParser,
  createBlankDocxBuffer,
  composeDocument,
  executeHeaderFooter,
} from '@attentionform/astra-engine';

const NUMBERING = 'relationships/numbering';

function rels(buf: Buffer): Array<{ id: string; type: string; target: string }> {
  const xml = new AdmZip(buf).getEntry('word/_rels/document.xml.rels')!.getData().toString('utf8');
  return [...xml.matchAll(/<Relationship\b[^>]*>/g)].map(m => ({
    id: /Id="([^"]*)"/.exec(m[0])![1],
    type: /Type="([^"]*)"/.exec(m[0])![1],
    target: /Target="([^"]*)"/.exec(m[0])![1],
  }));
}

// Mirrors the MCP session: the session AST is a separate object from the one
// composeDocument builds, its relationships are copied from the fresh AST, and
// a bookkeeping rezip then writes the numbering rel into the package only.
async function build(afterCompose: (ast: any, parser: DocxParser) => void) {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const session: any = await parser.getAST();
  await composeDocument(parser, '# T\n\n- one\n- two\n');
  const fresh: any = await parser.getAST();
  for (const k of Object.keys(session.relationships)) delete session.relationships[k];
  Object.assign(session.relationships, fresh.relationships);
  parser.rezip();
  afterCompose(session, parser);
  return rels(parser.rezip());
}

describe('header/footer allocation never overwrites the numbering relationship', () => {
  it('keeps numbering.xml reachable after a footer is set on a doc with a list', async () => {
    const list = await build((ast, p) =>
      executeHeaderFooter({ action: 'headerFooter', operation: 'set', targetType: 'footer', text: 'f' } as any, ast, p));
    const ids = list.map(r => r.id);
    expect(new Set(ids).size).toBe(ids.length);
    expect(list.some(r => r.type.endsWith(NUMBERING) && r.target === 'numbering.xml')).toBe(true);
    expect(list.some(r => r.type.endsWith('/footer'))).toBe(true);
  });

  it('keeps it with both header and footer', async () => {
    const list = await build((ast, p) => {
      executeHeaderFooter({ action: 'headerFooter', operation: 'set', targetType: 'header', text: 'h' } as any, ast, p);
      executeHeaderFooter({ action: 'headerFooter', operation: 'set', targetType: 'footer', text: 'f' } as any, ast, p);
    });
    expect(new Set(list.map(r => r.id)).size).toBe(list.length);
    expect(list.filter(r => r.type.endsWith(NUMBERING))).toHaveLength(1);
  });

  it('control: a list without header/footer keeps its numbering rel', async () => {
    const list = await build(() => {});
    expect(list.filter(r => r.type.endsWith(NUMBERING))).toHaveLength(1);
  });
});

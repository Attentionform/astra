import { describe, it, expect, vi } from 'vitest';
import {
  DocxParser,
  createBlankDocxBuffer,
  executeInsert,
} from '@attentionform/astra-engine';

async function freshDoc(): Promise<{ parser: DocxParser; ast: any; anchor: string }> {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const anchor = ast.sections[0].nodes[0].id as string;
  return { parser, ast, anchor };
}

function orderedIds(ast: any): string[] {
  return ast.sections.flatMap((s: any) => s.nodes).map((n: any) => n.id);
}

function domOrderIds(ast: any): string[] {
  const first = ast.sections[0].nodes[0] as any;
  const parent = first.xml.parentNode as Element;
  const out: string[] = [];
  for (let i = 0; i < parent.childNodes.length; i++) {
    const el = parent.childNodes[i] as unknown as { nodeType?: number; getAttribute?: (k: string) => string | null };
    if (el.nodeType === 1 && el.getAttribute) {
      const id = el.getAttribute('data-astra-id');
      if (id) out.push(id);
    }
  }
  return out;
}

describe('insert wave 1e: order, strict xml, no silent drops', () => {
  it('three same-anchor position:after inserts preserve call order (FIFO, not LIFO)', async () => {
    const { ast, anchor } = await freshDoc();
    await executeInsert({
      action: 'insert', type: 'paragraph', anchor, position: 'after',
      text: 'first', id: 'w1e-first',
    } as any, ast, undefined as any);
    await executeInsert({
      action: 'insert', type: 'paragraph', anchor, position: 'after',
      text: 'second', id: 'w1e-second',
    } as any, ast, undefined as any);
    await executeInsert({
      action: 'insert', type: 'paragraph', anchor, position: 'after',
      text: 'third', id: 'w1e-third',
    } as any, ast, undefined as any);

    const ids = orderedIds(ast);
    const iAnchor = ids.indexOf(anchor);
    expect(ids.slice(iAnchor + 1, iAnchor + 4)).toEqual(['w1e-first', 'w1e-second', 'w1e-third']);

    const dom = domOrderIds(ast);
    const dAnchor = dom.indexOf(anchor);
    expect(dom.slice(dAnchor + 1, dAnchor + 4)).toEqual(['w1e-first', 'w1e-second', 'w1e-third']);
  });

  it('truncated / unclosed XML rejects with E_VALIDATION (never an auto-repaired empty paragraph)', async () => {
    const truncated = [
      '<w:p><w:r><w:t>hello</w:t></w:r>',
      '<w:p><w:r><w:t>hello',
      '<w:p><w:r></w:p>',
      '<w:p>',
    ];
    for (const xml of truncated) {
      const { ast, anchor } = await freshDoc();
      const before = orderedIds(ast).length;
      await expect(executeInsert({
        action: 'insert', type: 'xml', anchor, position: 'after', xml, id: 'w1e-trunc',
      } as any, ast, undefined as any)).rejects.toThrow(/E_VALIDATION/);
      expect(orderedIds(ast).length).toBe(before);
    }
  });

  it('multi-root xml fragments import ALL roots or reject loudly (never a silent drop)', async () => {
    const { ast, anchor } = await freshDoc();
    const before = orderedIds(ast).length;
    const xml =
      '<w:p><w:r><w:t>root one</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>root two</w:t></w:r></w:p>';
    let threw: unknown = null;
    try {
      await executeInsert({
        action: 'insert', type: 'xml', anchor, position: 'after', xml, id: 'w1e-multi',
      } as any, ast, undefined as any);
    } catch (err) {
      threw = err;
    }
    if (threw) {
      expect(String((threw as Error)?.message || threw)).toMatch(/E_VALIDATION/);
      expect(orderedIds(ast).length).toBe(before);
    } else {
      const after = ast.sections
        .flatMap((s: any) => s.nodes)
        .map((n: any) => (n.xml as Element).textContent);
      expect(after).toContain('root one');
      expect(after).toContain('root two');
    }
  });

  it('script payload warns-or-rejects (never silent success)', async () => {
    const warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => {});
    try {
      for (const xml of [
        '<script>alert(1)</script>',
        '<w:p><w:r><w:t>ok</w:t></w:r></w:p><script>alert(1)</script>',
        '<w:p><w:r><w:t>x</w:t></w:r><script>alert(1)</script></w:p>',
      ]) {
        const { ast, anchor } = await freshDoc();
        const before = orderedIds(ast).length;
        let threw: unknown = null;
        try {
          await executeInsert({
            action: 'insert', type: 'xml', anchor, position: 'after', xml, id: 'w1e-script',
          } as any, ast, undefined as any);
        } catch (err) {
          threw = err;
        }
        if (threw) {
          expect(String((threw as Error)?.message || threw)).toMatch(/E_VALIDATION/i);
          expect(orderedIds(ast).length).toBe(before);
        } else {
          expect(warnSpy).toHaveBeenCalled();
          const bodies = ast.sections
            .flatMap((s: any) => s.nodes)
            .map((n: any) => new XMLSerializer().serializeToString(n.xml as Element));
          expect(bodies.join('')).not.toContain('alert(1)');
        }
        warnSpy.mockClear();
      }
    } finally {
      warnSpy.mockRestore();
    }
  });
});

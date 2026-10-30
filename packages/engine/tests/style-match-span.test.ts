/**
 * Word-level span styling pins (`match` on styleOp).
 *
 * A style op without `match` styles every run in the node (legacy bytes).
 * With `match`, only the matched substring gains the run props: the target
 * run splits into pre/matched/post (or first-tail + whole middles +
 * last-head across runs), paragraph props stay paragraph-wide, absent
 * matches fail closed, and hyperlink/field boundary crosses refuse.
 */
import { describe, it, expect } from 'vitest';
import { XMLSerializer } from '@xmldom/xmldom';
import { DocxParser, createBlankDocxBuffer } from '../src/parser';
import { OperationEngine } from '../src/operation-engine';
import { executeStyle, findMatchSpans } from '../src/operations/style';
import type { StyleOperation } from '../src/operations/style';
import { validateOperationsStrict } from '../src/operation-schema';
import { ToonGenerator } from '../src/toon';
import type { Node } from '../src/ast';

const ser = new XMLSerializer();

async function freshDoc() {
  const parser = new DocxParser(createBlankDocxBuffer('Span Test'));
  const ast = await parser.getAST();
  return { parser, ast };
}

function lastNode(ast: { sections: Array<{ nodes: Node[] }> }): Node {
  const nodes = ast.sections[0].nodes;
  return nodes[nodes.length - 1];
}

function xmlOf(node: Node): string {
  return ser.serializeToString(node.xml);
}

function runText(run: Element): string {
  let out = '';
  const kids = run.childNodes;
  for (let i = 0; i < kids.length; i++) {
    const c = kids[i] as unknown as Element;
    if (!c || c.nodeType !== 1) continue;
    if (c.tagName === 'w:t') out += c.textContent || '';
    else if (c.tagName === 'w:tab') out += '\t';
    else if (c.tagName === 'w:br') out += '\n';
  }
  return out;
}

function runsOf(node: Node): Array<{ text: string; xml: string }> {
  const els = Array.from(node.xml.getElementsByTagName('w:r')) as unknown as Element[];
  return els.map((r) => ({ text: runText(r), xml: ser.serializeToString(r) }));
}

function paraText(node: Node): string {
  return runsOf(node).map((r) => r.text).join('');
}

async function insertParagraph(engine: OperationEngine, ast: { sections: Array<{ nodes: Node[] }> }, text: string): Promise<Node> {
  const anchor = lastNode(ast).id;
  await engine.execute({ action: 'insert', type: 'paragraph', anchor, position: 'after', text } as never);
  return lastNode(ast);
}

async function insertXml(engine: OperationEngine, ast: { sections: Array<{ nodes: Node[] }> }, xml: string): Promise<Node> {
  const anchor = lastNode(ast).id;
  await engine.execute({ action: 'insert', type: 'xml', anchor, position: 'after', xml } as never);
  return lastNode(ast);
}

function codeOf(err: unknown): unknown {
  return (err as { code?: unknown } | null | undefined)?.code;
}

describe('single-run three-way split', () => {
  it('styles only the matched word and preserves the edges', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertParagraph(engine, ast, 'Hello Riverbend world');
    const before = para.metadata.text;
    await engine.execute({ action: 'style', target: para.id, match: 'Riverbend', color: '047857', bold: true } as never);

    const runs = runsOf(para);
    expect(runs.map((r) => r.text)).toEqual(['Hello ', 'Riverbend', ' world']);
    expect(runs[1].xml).toContain('w:val="047857"');
    expect(runs[1].xml).toContain('<w:b/>');
    expect(runs[0].xml).not.toContain('w:color');
    expect(runs[0].xml).not.toContain('<w:b/>');
    expect(runs[2].xml).not.toContain('w:color');
    expect(runs[2].xml).not.toContain('<w:b/>');
    // Whitespace edges keep xml:space so Word does not trim them.
    expect(runs[0].xml).toContain('xml:space="preserve"');
    expect(runs[2].xml).toContain('xml:space="preserve"');
    // Adjacent text untouched: concatenated text and metadata agree.
    expect(paraText(para)).toBe('Hello Riverbend world');
    expect(para.metadata.text).toBe(before);
  });
});

describe('cross-run span', () => {
  it('splits first tail plus whole middles plus last head without coalescing', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertXml(
      engine,
      ast,
      '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Hello </w:t></w:r>' +
        '<w:r><w:rPr><w:i/></w:rPr><w:t>River</w:t></w:r>' +
        '<w:r><w:t>bend world</w:t></w:r></w:p>'
    );
    await engine.execute({ action: 'style', target: para.id, match: 'lo Riverbend', color: 'FF0000' } as never);

    const runs = runsOf(para);
    expect(runs.map((r) => r.text)).toEqual(['Hel', 'lo ', 'River', 'bend', ' world']);
    // First-run head keeps its own formatting only; tail gains the new color.
    expect(runs[0].xml).toContain('<w:b/>');
    expect(runs[0].xml).not.toContain('w:color');
    expect(runs[1].xml).toContain('<w:b/>');
    expect(runs[1].xml).toContain('w:val="FF0000"');
    // Middle runs keep their formatting (italic survives, not coalesced away).
    expect(runs[2].xml).toContain('<w:i/>');
    expect(runs[2].xml).toContain('w:val="FF0000"');
    expect(runs[3].xml).toContain('w:val="FF0000"');
    expect(runs[3].xml).not.toContain('<w:i/>');
    expect(runs[3].xml).not.toContain('<w:b/>');
    // Last-run tail untouched.
    expect(runs[4].xml).not.toContain('w:color');
    expect(paraText(para)).toBe('Hello Riverbend world');
  });
});

describe('nested match resolution', () => {
  it('honors a params-nested match instead of styling the whole paragraph', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertParagraph(engine, ast, 'aaa bbb aaa');
    const checked = validateOperationsStrict([
      { action: 'style', target: para.id, params: { match: 'bbb', bold: true } },
    ]);
    expect(checked.ok).toBe(true);
    await engine.execute({ action: 'style', target: para.id, params: { match: 'bbb', bold: true } } as never);

    const runs = runsOf(para);
    expect(runs.map((r) => r.text)).toEqual(['aaa ', 'bbb', ' aaa']);
    expect(runs[1].xml).toContain('<w:b/>');
    expect(runs[0].xml).not.toContain('<w:b/>');
    expect(runs[2].xml).not.toContain('<w:b/>');
  });

  it('honors a properties-nested match', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertParagraph(engine, ast, 'aaa bbb aaa');
    await engine.execute({ action: 'style', target: para.id, properties: { match: 'bbb', italic: true } } as never);

    const runs = runsOf(para);
    expect(runs.map((r) => r.text)).toEqual(['aaa ', 'bbb', ' aaa']);
    expect(runs[1].xml).toContain('<w:i/>');
    expect(runs[0].xml).not.toContain('<w:i/>');
    expect(runs[2].xml).not.toContain('<w:i/>');
  });

  it('lets a flat match win over a params-nested match', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertParagraph(engine, ast, 'aaa bbb');
    await engine.execute({ action: 'style', target: para.id, match: 'aaa', params: { match: 'bbb', bold: true } } as never);

    const runs = runsOf(para);
    expect(runs.map((r) => r.text)).toEqual(['aaa', ' bbb']);
    expect(runs[0].xml).toContain('<w:b/>');
    expect(runs[1].xml).not.toContain('<w:b/>');
  });
});

describe('occurrence and case selection', () => {
  it('styles only the first occurrence by default and all with matchAll', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertParagraph(engine, ast, 'cat hat cat');
    await engine.execute({ action: 'style', target: para.id, match: 'cat', bold: true } as never);

    let runs = runsOf(para);
    expect(runs.map((r) => r.text)).toEqual(['cat', ' hat cat']);
    expect(runs[0].xml).toContain('<w:b/>');
    expect(runs[1].xml).not.toContain('<w:b/>');

    await engine.execute({ action: 'style', target: para.id, match: 'cat', matchAll: true, bold: true, italic: true } as never);
    runs = runsOf(para);
    const bolds = runs.filter((r) => r.text === 'cat');
    expect(bolds.length).toBe(2);
    for (const b of bolds) {
      expect(b.xml).toContain('<w:b/>');
      expect(b.xml).toContain('<w:i/>');
    }
    expect(paraText(para)).toBe('cat hat cat');
  });

  it('matches case-insensitively with matchCase false', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertParagraph(engine, ast, 'Cat Nap');
    await engine.execute({ action: 'style', target: para.id, match: 'cat', matchCase: false, bold: true } as never);

    const runs = runsOf(para);
    expect(runs.map((r) => r.text)).toEqual(['Cat', ' Nap']);
    expect(runs[0].xml).toContain('<w:b/>');
    expect(runs[1].xml).not.toContain('<w:b/>');
  });

  it('stays case-sensitive by default', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertParagraph(engine, ast, 'Cat Nap');
    // Lowercase 'cat' never occurs case-sensitively, so the default
    // case-sensitive search fails closed instead of styling 'Cat'.
    const err = await engine.execute({ action: 'style', target: para.id, match: 'cat', bold: true } as never)
      .then(() => null, (e: unknown) => e);
    expect(err).not.toBe(null);
    expect(codeOf(err)).toBe('E_NOT_FOUND');
  });
});

describe('paragraph-wide versus span behavior', () => {
  it('applies paragraph props paragraph-wide while run props stay on the span', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertParagraph(engine, ast, 'Hello Riverbend world');
    await engine.execute({ action: 'style', target: para.id, match: 'Riverbend', alignment: 'center', color: '047857' } as never);

    expect(xmlOf(para)).toContain('<w:jc w:val="center"/>');
    const runs = runsOf(para);
    expect(runs.map((r) => r.text)).toEqual(['Hello ', 'Riverbend', ' world']);
    expect(runs[1].xml).toContain('w:val="047857"');
    expect(runs[0].xml).not.toContain('w:color');
    expect(runs[2].xml).not.toContain('w:color');
  });

  it('matches across a tab and keeps the tab element', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertXml(engine, ast, '<w:p><w:r><w:t>left</w:t><w:tab/><w:t>right</w:t></w:r></w:p>');
    await engine.execute({ action: 'style', target: para.id, match: 'left', bold: true } as never);

    const runs = runsOf(para);
    expect(runs.map((r) => r.text)).toEqual(['left', '\tright']);
    expect(runs[0].xml).toContain('<w:b/>');
    expect(runs[1].xml).not.toContain('<w:b/>');
    expect(runs[1].xml).toContain('<w:tab/>');
    expect(paraText(para)).toBe('left\tright');
  });
});

describe('fail-closed gates', () => {
  it('throws E_NOT_FOUND with a remedy on an absent match', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    await insertParagraph(engine, ast, 'abc def');

    const err = await engine.execute({ action: 'style', target: lastNode(ast).id, match: 'zzz', bold: true } as never)
      .then(() => null, (e: unknown) => e);
    expect(err).not.toBe(null);
    expect(codeOf(err)).toBe('E_NOT_FOUND');
    expect(String((err as Error).message)).toContain('inspect_document');

    const live = ast.sections[0].nodes.find((n) => n.metadata.text === 'abc def');
    expect(live).toBeDefined();
    expect(paraText(live as Node)).toBe('abc def');
    expect(xmlOf(live as Node)).not.toContain('<w:b/>');
  });

  it('refuses a match crossing a hyperlink boundary and leaves the paragraph alone', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    await insertXml(
      engine,
      ast,
      '<w:p><w:r><w:t>see </w:t></w:r>' +
        '<w:hyperlink w:anchor="b1" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>' +
        '<w:t>click here</w:t></w:r></w:hyperlink><w:r><w:t> now</w:t></w:r></w:p>'
    );
    const err = await engine.execute({ action: 'style', target: lastNode(ast).id, match: 'see click', bold: true } as never)
      .then(() => null, (e: unknown) => e);
    expect(err).not.toBe(null);
    expect(codeOf(err)).toBe('E_VALIDATION');
    expect(String((err as Error).message)).toContain('hyperlink');

    const live = ast.sections[0].nodes.find((n) => n.metadata.text === 'see click here now');
    expect(live).toBeDefined();
    expect(paraText(live as Node)).toBe('see click here now');
    expect(xmlOf(live as Node)).not.toContain('<w:b/>');
  });

  it('styles a match fully inside a hyperlink without touching the link shell', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertXml(
      engine,
      ast,
      '<w:p><w:r><w:t>see </w:t></w:r>' +
        '<w:hyperlink w:anchor="b1" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>' +
        '<w:t>click here</w:t></w:r></w:hyperlink><w:r><w:t> now</w:t></w:r></w:p>'
    );
    await engine.execute({ action: 'style', target: para.id, match: 'click', bold: true } as never);

    const xml = xmlOf(para);
    expect(xml).toContain('<w:hyperlink');
    const linkRuns = Array.from(para.xml.getElementsByTagName('w:hyperlink')[0].getElementsByTagName('w:r')) as unknown as Element[];
    const texts = linkRuns.map(runText);
    expect(texts).toEqual(['click', ' here']);
    expect(ser.serializeToString(linkRuns[0])).toContain('<w:b/>');
    expect(ser.serializeToString(linkRuns[0])).toContain('w:val="Hyperlink"');
    expect(ser.serializeToString(linkRuns[1])).not.toContain('<w:b/>');
    expect(paraText(para)).toBe('see click here now');
  });

  it('rejects match on a non-paragraph node', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const anchor = lastNode(ast).id;
    await engine.execute({ action: 'insert', type: 'table', position: 'after', anchor, headers: ['A'], rows: [['x']] } as never);
    const table = ast.sections[0].nodes.find((n) => n.type === 'table');
    expect(table).toBeDefined();

    const err = await engine.execute({ action: 'style', target: (table as Node).id, match: 'x', bold: true } as never)
      .then(() => null, (e: unknown) => e);
    expect(err).not.toBe(null);
    expect(codeOf(err)).toBe('E_VALIDATION');
  });

  it('rejects an empty match at the schema gate', () => {
    expect(() => validateOperationsStrict([{ action: 'style', target: 'p1', match: '' }])).toThrow(/E_VALIDATION/);
  });

  it('accepts match, matchAll, and matchCase at the schema gate', () => {
    const checked = validateOperationsStrict([
      { action: 'style', target: 'p1', match: 'Riverbend', matchAll: true, matchCase: false, bold: true },
    ]);
    expect(checked.ok).toBe(true);
  });
});

describe('failure atomicity (KILL-1)', () => {
  it('para-prop + match-miss leaves the paragraph byte-identical', async () => {
    const { ast } = await freshDoc();
    const engine = new OperationEngine(ast, undefined);
    const para = await insertParagraph(engine, ast, 'abc def');
    const before = xmlOf(para);
    const beforeVersion = para.version;

    // Direct op call (no engine snapshot rollback): the op itself must
    // refuse before touching a byte.
    const err = await Promise.resolve()
      .then(() => executeStyle({ action: 'style', target: para.id, match: 'zzz', alignment: 'center' } as never, ast))
      .then(() => null, (e: unknown) => e);
    expect(err).not.toBe(null);
    expect(codeOf(err)).toBe('E_NOT_FOUND');
    expect(xmlOf(para)).toBe(before);
    expect(xmlOf(para)).not.toContain('w:jc');
    expect(para.version).toBe(beforeVersion);
  });

  it('para-prop + run-prop + match-miss leaves zero bytes touched', async () => {
    const { ast } = await freshDoc();
    const engine = new OperationEngine(ast, undefined);
    const para = await insertParagraph(engine, ast, 'abc def');
    const before = xmlOf(para);
    const beforeVersion = para.version;

    const err = await Promise.resolve()
      .then(() => executeStyle({ action: 'style', target: para.id, match: 'zzz', alignment: 'center', bold: true } as never, ast))
      .then(() => null, (e: unknown) => e);
    expect(err).not.toBe(null);
    expect(codeOf(err)).toBe('E_NOT_FOUND');
    expect(xmlOf(para)).toBe(before);
    expect(xmlOf(para)).not.toContain('w:jc');
    expect(xmlOf(para)).not.toContain('<w:b/>');
    expect(para.version).toBe(beforeVersion);
  });

  it('para-prop + boundary-crossing match leaves zero bytes touched', async () => {
    const { ast } = await freshDoc();
    const engine = new OperationEngine(ast, undefined);
    const para = await insertXml(
      engine,
      ast,
      '<w:p><w:r><w:t>see </w:t></w:r>' +
        '<w:hyperlink w:anchor="b1" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>' +
        '<w:t>click here</w:t></w:r></w:hyperlink><w:r><w:t> now</w:t></w:r></w:p>'
    );
    const before = xmlOf(para);
    const beforeVersion = para.version;

    const err = await Promise.resolve()
      .then(() => executeStyle({ action: 'style', target: para.id, match: 'see click', alignment: 'center', bold: true } as never, ast))
      .then(() => null, (e: unknown) => e);
    expect(err).not.toBe(null);
    expect(codeOf(err)).toBe('E_VALIDATION');
    expect(xmlOf(para)).toBe(before);
    expect(xmlOf(para)).not.toContain('w:jc');
    expect(xmlOf(para)).not.toContain('<w:b/>');
    expect(para.version).toBe(beforeVersion);
  });

  it('multi-target batch validates every span before mutating any node', async () => {
    const { ast } = await freshDoc();
    const engine = new OperationEngine(ast, undefined);
    const p1 = await insertParagraph(engine, ast, 'alpha one');
    const p2 = await insertParagraph(engine, ast, 'beta two');
    const before1 = xmlOf(p1);
    const before2 = xmlOf(p2);
    const v1 = p1.version;
    const v2 = p2.version;

    // 'alpha' hits p1 but misses p2: the op must refuse with p1 untouched.
    const err = await Promise.resolve()
      .then(() => executeStyle({ action: 'style', target: [p1.id, p2.id], match: 'alpha', bold: true } as never, ast))
      .then(() => null, (e: unknown) => e);
    expect(err).not.toBe(null);
    expect(codeOf(err)).toBe('E_NOT_FOUND');
    expect(xmlOf(p1)).toBe(before1);
    expect(xmlOf(p2)).toBe(before2);
    expect(p1.version).toBe(v1);
    expect(p2.version).toBe(v2);
  });

  it('match on a non-paragraph node with para-props leaves zero bytes touched', async () => {
    const { ast } = await freshDoc();
    const engine = new OperationEngine(ast, undefined);
    const anchor = lastNode(ast).id;
    await engine.execute({ action: 'insert', type: 'table', position: 'after', anchor, headers: ['A'], rows: [['x']] } as never);
    const table = ast.sections[0].nodes.find((n) => n.type === 'table') as Node;
    const before = xmlOf(table);
    const beforeVersion = table.version;

    const err = await Promise.resolve()
      .then(() => executeStyle({ action: 'style', target: table.id, match: 'x', alignment: 'center', bold: true } as never, ast))
      .then(() => null, (e: unknown) => e);
    expect(err).not.toBe(null);
    expect(codeOf(err)).toBe('E_VALIDATION');
    expect(xmlOf(table)).toBe(before);
    expect(table.version).toBe(beforeVersion);
  });
});

describe('legacy preservation', () => {
  it('styles every run and splits nothing when match is omitted', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertXml(
      engine,
      ast,
      '<w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:t>world</w:t></w:r></w:p>'
    );
    await engine.execute({ action: 'style', target: para.id, bold: true, color: 'FF0000' } as never);

    const runs = runsOf(para);
    expect(runs.length).toBe(2);
    expect(runs.map((r) => r.text)).toEqual(['Hello ', 'world']);
    for (const r of runs) {
      expect(r.xml).toContain('<w:b/>');
      expect(r.xml).toContain('w:val="FF0000"');
    }
    expect(paraText(para)).toBe('Hello world');
  });
});

describe('props bag merge (flat-wins per key)', () => {
  it('merges flat run props with a params-nested match instead of dropping them', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertParagraph(engine, ast, 'aaa bbb');
    await engine.execute({ action: 'style', target: para.id, bold: true, params: { match: 'bbb' } } as never);

    const runs = runsOf(para);
    expect(runs.map((r) => r.text)).toEqual(['aaa ', 'bbb']);
    expect(runs[1].xml).toContain('<w:b/>');
    expect(runs[0].xml).not.toContain('<w:b/>');
  });

  it('merges a properties-bag run prop under a flat match', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertParagraph(engine, ast, 'aaa bbb');
    await engine.execute({ action: 'style', target: para.id, match: 'bbb', properties: { bold: true } } as never);

    const runs = runsOf(para);
    expect(runs.map((r) => r.text)).toEqual(['aaa ', 'bbb']);
    expect(runs[1].xml).toContain('<w:b/>');
    expect(runs[0].xml).not.toContain('<w:b/>');
  });

  it('lets flat color win over a properties-bag color on the same key', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertParagraph(engine, ast, 'aaa bbb');
    await engine.execute({ action: 'style', target: para.id, match: 'bbb', color: '00FF00', properties: { color: 'FF0000', match: 'bbb' } } as never);

    const runs = runsOf(para);
    expect(runs.map((r) => r.text)).toEqual(['aaa ', 'bbb']);
    expect(runs[1].xml).toContain('w:val="00FF00"');
    expect(runs[1].xml).not.toContain('w:val="FF0000"');
  });

  it('lets a properties-bag color win over a params-bag color on the same key', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertParagraph(engine, ast, 'aaa bbb');
    await engine.execute({
      action: 'style', target: para.id,
      properties: { color: '00FF00', match: 'bbb' }, params: { color: 'FF0000', match: 'bbb' },
    } as never);

    const runs = runsOf(para);
    expect(runs.map((r) => r.text)).toEqual(['aaa ', 'bbb']);
    expect(runs[1].xml).toContain('w:val="00FF00"');
    expect(runs[1].xml).not.toContain('w:val="FF0000"');
  });

  it('lets flat win when all three bags set the same key', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertParagraph(engine, ast, 'aaa bbb');
    await engine.execute({
      action: 'style', target: para.id, match: 'bbb', color: '0000FF',
      properties: { color: '00FF00', match: 'bbb' }, params: { color: 'FF0000', match: 'bbb' },
    } as never);

    const runs = runsOf(para);
    expect(runs.map((r) => r.text)).toEqual(['aaa ', 'bbb']);
    expect(runs[1].xml).toContain('w:val="0000FF"');
    expect(runs[1].xml).not.toContain('w:val="00FF00"');
    expect(runs[1].xml).not.toContain('w:val="FF0000"');
  });
});

describe('unread style-key locations refuse instead of silently ignoring', () => {
  it('refuses style keys nested under params.params before touching any byte', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertParagraph(engine, ast, 'aaa bbb');
    const before = xmlOf(para);
    const beforeVersion = para.version;

    const err = await Promise.resolve()
      .then(() => executeStyle({ action: 'style', target: para.id, color: 'FF0000', params: { params: { match: 'bbb' } } } as never, ast))
      .then(() => null, (e: unknown) => e);
    expect(err).not.toBe(null);
    expect(codeOf(err)).toBe('E_VALIDATION');
    expect(String((err as Error).message)).toContain('params.params');
    expect(String((err as Error).message)).toContain('silently ignored');
    const remedy = (err as { remedy?: string }).remedy ?? '';
    expect(remedy).toContain('"properties"');
    expect(remedy).toContain('"params"');
    expect(xmlOf(para)).toBe(before);
    expect(para.version).toBe(beforeVersion);
  });

  it('refuses array-wrapped style keys through the gate path (params.wrapper[0])', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertParagraph(engine, ast, 'aaa bbb');
    const op = { action: 'style', target: para.id, color: 'FF0000', params: { wrapper: [{ match: 'bbb' }] } };
    expect(validateOperationsStrict([op]).ok).toBe(true);
    const before = xmlOf(para);
    const res = await engine.executeBatch({ operations: [op] as never, sequential: false, rollbackOnError: false });
    expect(res.success).toBe(false);
    expect(String(res.errors[0].error)).toContain('params.wrapper[0]');
    expect(String(res.errors[0].error)).toContain('silently ignored');
    expect(xmlOf(para)).toBe(before);
  });

  it('refuses a 12-level deep unread chain through the gate path (no depth cap)', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertParagraph(engine, ast, 'aaa bbb');
    let wrapper: Record<string, unknown> = { match: 'bbb' };
    for (let i = 0; i < 12; i++) wrapper = { a: wrapper };
    const op = { action: 'style', target: para.id, color: 'FF0000', params: wrapper };
    expect(validateOperationsStrict([op]).ok).toBe(true);
    const before = xmlOf(para);
    const res = await engine.executeBatch({ operations: [op] as never, sequential: false, rollbackOnError: false });
    expect(res.success).toBe(false);
    expect(String(res.errors[0].error)).toContain('params.a.a.a');
    expect(xmlOf(para)).toBe(before);
  });

  it('refuses a null-valued style key (gate rejects; engine-direct also fails closed)', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertParagraph(engine, ast, 'aaa bbb');
    const op = { action: 'style', target: para.id, color: null, params: { match: 'bbb' } };
    expect(() => validateOperationsStrict([op])).toThrow(/E_VALIDATION/);
    const before = xmlOf(para);
    const beforeVersion = para.version;
    const err = await Promise.resolve()
      .then(() => executeStyle(op as never, ast))
      .then(() => null, (e: unknown) => e);
    expect(err).not.toBe(null);
    expect(codeOf(err)).toBe('E_VALIDATION');
    expect(String((err as Error).message)).toContain('"color" to null');
    expect(xmlOf(para)).toBe(before);
    expect(para.version).toBe(beforeVersion);
  });

  it('keeps the documented mixed location working: flat color + params-nested match', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertParagraph(engine, ast, 'aaa bbb');
    await engine.execute({ action: 'style', target: para.id, color: 'FF0000', params: { match: 'bbb' } } as never);

    const runs = runsOf(para);
    expect(runs.map((r) => r.text)).toEqual(['aaa ', 'bbb']);
    expect(runs[1].xml).toContain('w:val="FF0000"');
    expect(runs[0].xml).not.toContain('w:color');
  });

  it('keeps nested border configs in bags exempt (their keys are read, not unread)', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertParagraph(engine, ast, 'aaa bbb');
    await engine.execute({
      action: 'style', target: para.id, match: 'bbb',
      properties: { borderBottom: { color: 'FF0000', sz: 8 } },
    } as never);

    const xml = xmlOf(para);
    expect(xml).toContain('<w:pBdr>');
    expect(xml).toContain('w:color="FF0000"');
    expect(paraText(para)).toBe('aaa bbb');
  });
});

describe('surrogate-pair refusal', () => {
  it('refuses a lone lead surrogate in the needle and leaves zero bytes touched', async () => {
    const { ast } = await freshDoc();
    const engine = new OperationEngine(ast, undefined);
    const para = await insertParagraph(engine, ast, 'A😀B');
    const before = xmlOf(para);
    const beforeVersion = para.version;

    const err = await Promise.resolve()
      .then(() => executeStyle({ action: 'style', target: para.id, match: 'A\ud83d', bold: true } as never, ast))
      .then(() => null, (e: unknown) => e);
    expect(err).not.toBe(null);
    expect(codeOf(err)).toBe('E_VALIDATION');
    expect(String((err as Error).message)).toContain('surrogate');
    expect(xmlOf(para)).toBe(before);
    expect(para.version).toBe(beforeVersion);
  });

  it('refuses a span cutting a haystack pair and leaves zero bytes touched', async () => {
    const { ast } = await freshDoc();
    const engine = new OperationEngine(ast, undefined);
    const para = await insertParagraph(engine, ast, '😀X');
    const before = xmlOf(para);
    const beforeVersion = para.version;

    const err = await Promise.resolve()
      .then(() => executeStyle({ action: 'style', target: para.id, match: '\ude00X', bold: true } as never, ast))
      .then(() => null, (e: unknown) => e);
    expect(err).not.toBe(null);
    expect(codeOf(err)).toBe('E_VALIDATION');
    expect(String((err as Error).message)).toContain('surrogate');
    expect(xmlOf(para)).toBe(before);
    expect(para.version).toBe(beforeVersion);
  });

  it('still styles a match that is a whole surrogate pair', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertParagraph(engine, ast, 'A😀B');
    await engine.execute({ action: 'style', target: para.id, match: '😀', bold: true } as never);

    const runs = runsOf(para);
    expect(runs.map((r) => r.text)).toEqual(['A', '😀', 'B']);
    expect(runs[1].xml).toContain('<w:b/>');
    expect(runs[0].xml).not.toContain('<w:b/>');
    expect(runs[2].xml).not.toContain('<w:b/>');
  });
});

describe('defensive empty-needle clamp', () => {
  it('returns no spans for an empty needle instead of looping', () => {
    expect(findMatchSpans('abc', '', true, true)).toEqual([]);
    expect(findMatchSpans('abc', '', false, false)).toEqual([]);
    expect(findMatchSpans('', '', true, true)).toEqual([]);
  });
});

describe('para-only boundary harmonization', () => {
  it('refuses a cross-hyperlink match with para-only props and leaves zero bytes touched', async () => {
    const { ast } = await freshDoc();
    const engine = new OperationEngine(ast, undefined);
    const para = await insertXml(
      engine,
      ast,
      '<w:p><w:r><w:t>see </w:t></w:r>' +
        '<w:hyperlink w:anchor="b1" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>' +
        '<w:t>click here</w:t></w:r></w:hyperlink><w:r><w:t> now</w:t></w:r></w:p>'
    );
    const before = xmlOf(para);
    const beforeVersion = para.version;

    const err = await Promise.resolve()
      .then(() => executeStyle({ action: 'style', target: para.id, match: 'see click', alignment: 'center' } as never, ast))
      .then(() => null, (e: unknown) => e);
    expect(err).not.toBe(null);
    expect(codeOf(err)).toBe('E_VALIDATION');
    expect(xmlOf(para)).toBe(before);
    expect(xmlOf(para)).not.toContain('w:jc');
    expect(para.version).toBe(beforeVersion);
  });
});

describe('flat field typing', () => {
  it('exposes match and run props with precise types on StyleOperation', () => {
    const op: StyleOperation = { action: 'style', target: 'p1', match: 'x', matchAll: true, matchCase: false, bold: true };
    const m: string | undefined = op.match;
    const all: boolean | undefined = op.matchAll;
    const cs: boolean | undefined = op.matchCase;
    const b: boolean | undefined = op.bold;
    expect([m, all, cs, b]).toEqual(['x', true, false, true]);
  });
});

describe('round-trip and reporting', () => {
  it('survives a save and reopen with spans intact', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertParagraph(engine, ast, 'Hello Riverbend world');
    await engine.execute({ action: 'style', target: para.id, match: 'Riverbend', color: '047857', bold: true } as never);

    const buffer = parser.rezip();
    const reparsed = new DocxParser(buffer);
    const ast2 = await reparsed.getAST();
    const live = ast2.sections.flatMap((s) => s.nodes).find((n) => n.metadata.text === 'Hello Riverbend world');
    expect(live).toBeDefined();
    const runs = runsOf(live as Node);
    expect(runs.map((r) => r.text)).toEqual(['Hello ', 'Riverbend', ' world']);
    expect(runs[1].xml).toContain('w:val="047857"');
    expect(runs[1].xml).toContain('<w:b/>');
    expect(runs[0].xml).not.toContain('w:color');
    expect(runs[2].xml).not.toContain('w:color');
  });

  it('reflects the styled span in full-tier TOON detail', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const para = await insertParagraph(engine, ast, 'Hello Riverbend world');
    await engine.execute({ action: 'style', target: para.id, match: 'Riverbend', color: '047857', bold: true } as never);

    const gen = new ToonGenerator();
    const full = gen.toon(ast, { fidelity: 'full', sectionId: ast.sections[0].id });
    expect(full.tier).toBe('full');
    const row = (full.rows as Array<{ id: string; text: string; detail?: string; flags?: string }>).find((r) => r.text === 'Hello Riverbend world');
    expect(row).toBeDefined();
    expect(row?.detail ?? '').toContain('047857');
  });
});

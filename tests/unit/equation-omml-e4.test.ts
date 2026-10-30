/**
 * Phase E4 — LaTeX→OMML converter (ooxml/omml.ts) + export upgrades.
 *
 * - matrices (matrix/pmatrix/bmatrix/array → bare m:m, NOT delimiter-wrapped)
 *   convert structurally and validate through the real OOXML validator;
 * - new symbol coverage (full Greek, arrows, sets/logic) + \det lit runs;
 * - nesting depth bound is 6 (5- and 6-deep pass, 7-deep fails loudly);
 * - retired loud cases stay loud (\vec, \oint, align envs, stray &/\\);
 * - HTML export renders display equations as native MathML <math display>
 *   (well-formed by strict re-parse), inline math stays a span;
 * - Markdown export renders display equations as $$ blocks, inline keeps $.
 */
import { describe, it, expect } from 'vitest';
import { DOMParser } from '@xmldom/xmldom';
import {
  DocxParser,
  createBlankDocxBuffer,
  createMathElement,
  executeEquation,
  executeCallout,
  executeDelete,
  convertAstToHtml,
  convertAstToMarkdown,
} from '@attentionform/astra-engine';
import { validateDocxBuffer, assertValidDocx } from '../../packages/engine/src/ooxml/validator-gate';

async function fresh() {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast = await parser.getAST();
  return { parser, ast, anchor: ast.sections[0].nodes[0].id };
}

function xmlOf(parser: DocxParser, id: string): string {
  return parser.readNode(id);
}

function domOf(xml: string): Document {
  return new DOMParser().parseFromString(xml, 'text/xml');
}

function count(xml: string, tag: string): number {
  return domOf(xml).getElementsByTagName(tag).length;
}

/** Concatenated text of every m:t run in the paragraph. */
function allRunText(xml: string): string {
  const els = domOf(xml).getElementsByTagName('m:t');
  let out = '';
  for (let i = 0; i < els.length; i++) out += (els[i] as Element).textContent || '';
  return out;
}

async function insertEq(
  parser: DocxParser, ast: Awaited<ReturnType<DocxParser['getAST']>>, anchor: string, latex: string
): Promise<string> {
  const out = (await executeEquation(
    { action: 'equation', operation: 'insert', anchor, latex } as never,
    ast,
    parser
  )) as { id: string };
  return out.id;
}

describe('E4 matrices (bare m:m, ECMA-376 §22.1)', () => {
  const cases: Array<{ name: string; latex: string; rows: number; cells: number }> = [
    { name: 'pmatrix 2x2', latex: '\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}', rows: 2, cells: 4 },
    { name: 'matrix 1x3', latex: '\\begin{matrix} a & b & c \\end{matrix}', rows: 1, cells: 3 },
    { name: 'bmatrix 2x3 rectangular', latex: '\\begin{bmatrix} a & b & c \\\\ d & e & f \\end{bmatrix}', rows: 2, cells: 6 },
    { name: 'array 2x2', latex: '\\begin{array}{cc} a & b \\\\ c & d \\end{array}', rows: 2, cells: 4 },
    { name: 'vmatrix single column', latex: '\\begin{vmatrix} a \\\\ b \\\\ c \\end{vmatrix}', rows: 3, cells: 3 },
  ];

  it.each(cases.map((c) => [c.name, c] as const))('converts %s', async (_name, c) => {
    const { parser, ast, anchor } = await fresh();
    const id = await insertEq(parser, ast, anchor, c.latex);
    const xml = xmlOf(parser, id);
    expect(count(xml, 'm:m')).toBe(1);
    expect(count(xml, 'm:mr')).toBe(c.rows);
    expect(count(xml, 'm:e')).toBe(c.cells);
    // NOT delimiter-wrapped: pure matrix equations emit no m:d.
    expect(count(xml, 'm:d')).toBe(0);
    expect(allRunText(xml)).not.toContain('\\');
    expect(allRunText(xml)).not.toContain('begin');
  });

  it('matrix cells convert structurally (frac + sup inside)', async () => {
    const { parser, ast, anchor } = await fresh();
    const id = await insertEq(parser, ast, anchor, '\\begin{pmatrix} \\frac{1}{2} & x^2 \\\\ y & z \\end{pmatrix}');
    const xml = xmlOf(parser, id);
    expect(count(xml, 'm:m')).toBe(1);
    expect(count(xml, 'm:f')).toBe(1);
    expect(count(xml, 'm:sSup')).toBe(1);
    expect(allRunText(xml)).not.toContain('\\');
  });

  it('invisible delimiters emit explicit empty begChr/endChr (no Word "(" default)', async () => {
    const { parser, ast, anchor } = await fresh();
    const id = await insertEq(parser, ast, anchor, '\\left. \\frac{a}{b} \\right)');
    const xml = xmlOf(parser, id);
    expect(count(xml, 'm:d')).toBe(1);
    expect(count(xml, 'm:f')).toBe(1);
    // ECMA-376 defaults a missing begChr/endChr to "(" / ")", so the
    // invisible side must be explicit-empty, not absent.
    expect(xml).toContain('m:begChr');
    expect(xml).toMatch(/<m:begChr m:val=""\/>/);
    expect(xml).toContain('m:val=")"');
    expect(allRunText(xml)).not.toContain('\\');
  });

  it('matrix failures stay loud (nothing emitted)', async () => {
    const bad: Array<[string, string, RegExp]> = [
      ['empty cell', '\\begin{matrix} a & \\\\ c & d \\end{matrix}', /cell/],
      ['unclosed env', '\\begin{matrix} a & b', /begin/],
      ['mismatched end', '\\begin{matrix} a \\end{pmatrix}', /end/],
      ['ruled colspec', '\\begin{array}{c|c} a \\\\ b \\end{array}', /array/],
    ];
    for (const [, latex, namesConstruct] of bad) {
      const { parser, ast, anchor } = await fresh();
      const before = ast.sections[0].nodes.length;
      const docBefore = parser.readDocumentXml();
      let err: unknown = null;
      try {
        await executeEquation(
          { action: 'equation', operation: 'insert', anchor, latex } as never,
          ast,
          parser
        );
      } catch (e) {
        err = e;
      }
      expect(err, latex).not.toBeNull();
      const msg = (err as Error).message;
      expect(msg).toMatch(/E_VALIDATION/);
      expect(msg).toMatch(namesConstruct);
      expect(msg).toMatch(/position \d+/);
      expect(msg).toMatch(/Remedy:/);
      expect((err as unknown as { code?: string }).code).toBe('E_VALIDATION');
      expect(ast.sections[0].nodes.length).toBe(before);
      expect(parser.readDocumentXml()).toBe(docBefore);
    }
  });
});

describe('E4 symbols + det + depth bound', () => {
  it('new symbol coverage (Greek/arrows/sets/logic)', async () => {
    const { parser, ast, anchor } = await fresh();
    const latex =
      '\\delta \\lambda \\Omega \\Gamma \\varphi \\leftarrow \\Rightarrow \\leftrightarrow ' +
      '\\forall \\exists \\emptyset \\cup \\cap \\leq \\geq \\propto \\perp \\hbar \\det';
    const id = await insertEq(parser, ast, anchor, latex);
    const xml = xmlOf(parser, id);
    for (const sym of ['δ', 'λ', 'Ω', 'Γ', 'ϕ', '←', '⇒', '↔', '∀', '∃', '∅', '∪', '∩', '≤', '≥', '∝', '⊥', 'ℏ']) {
      expect(allRunText(xml)).toContain(sym);
    }
    expect(allRunText(xml)).not.toContain('\\');
  });

  it('\\det renders as a roman lit run', async () => {
    const { parser, ast, anchor } = await fresh();
    const id = await insertEq(parser, ast, anchor, '\\det A');
    const xml = xmlOf(parser, id);
    expect(xml).toContain('<m:lit');
    expect(allRunText(xml)).toContain('det');
  });

  it('depth 5 and 6 pass, depth 7 fails loudly', async () => {
    const { parser, ast, anchor } = await fresh();
    const d5 = '\\frac{a}{\\frac{b}{\\frac{c}{\\frac{d}{\\frac{e}{f}}}}}';
    const d6 = '\\frac{a}{\\frac{b}{\\frac{c}{\\frac{d}{\\frac{e}{\\frac{f}{g}}}}}}';
    const d7 = '\\frac{a}{\\frac{b}{\\frac{c}{\\frac{d}{\\frac{e}{\\frac{f}{\\frac{g}{h}}}}}}}';
    const id5 = await insertEq(parser, ast, anchor, d5);
    expect(count(xmlOf(parser, id5), 'm:f')).toBe(5);
    const id6 = await insertEq(parser, ast, anchor, d6);
    expect(count(xmlOf(parser, id6), 'm:f')).toBe(6);
    const before = ast.sections[0].nodes.length;
    try {
      await executeEquation({ action: 'equation', operation: 'insert', anchor, latex: d7 } as never, ast, parser);
      expect.unreachable('depth-7 input must fail');
    } catch (e) {
      expect((e as Error).message).toMatch(/E_VALIDATION/);
      expect((e as Error).message).toMatch(/nest|depth/i);
      expect((e as unknown as { code?: string }).code).toBe('E_VALIDATION');
    }
    expect(ast.sections[0].nodes.length).toBe(before);
  });

  it('retired loud cases stay loud', async () => {
    const bad: Array<[string, string, RegExp]> = [
      ['arrow-accent macro', '\\overrightarrow{x}', /\\overrightarrow/],
      ['exotic macro', '\\oint_{C} x dx', /\\oint/],
      ['alignment env', '\\begin{aligned} x \\end{aligned}', /aligned/],
      ['stray alignment tab', '\\alpha & \\beta', /&/],
    ];
    for (const [, latex, namesConstruct] of bad) {
      const { parser, ast, anchor } = await fresh();
      await expect(
        executeEquation({ action: 'equation', operation: 'insert', anchor, latex } as never, ast, parser)
      ).rejects.toThrow(namesConstruct);
    }
  });
});

describe('E4 validator proof (matrices + depth + symbols validate)', () => {
  it('combined E4 document passes the real OOXML validator', async () => {
    const { parser, ast, anchor } = await fresh();
    let at = anchor;
    const lats = [
      'x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}',
      '\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}',
      '\\begin{array}{cc} \\frac{1}{2} & x^2 \\\\ y_{i} & z \\end{array}',
      '\\sum_{i=1}^{n} a_i + \\int_{a}^{b} f(x) dx',
      '\\delta \\lambda \\Omega \\Rightarrow \\forall x \\in A, \\det M = 0',
      '\\frac{a}{\\frac{b}{\\frac{c}{\\frac{d}{\\frac{e}{\\frac{f}{g}}}}}}',
      '\\left( \\frac{a}{b} \\right) + |c|',
      '\\left. \\frac{a}{b} \\right)',
    ];
    for (const latex of lats) {
      at = await insertEq(parser, ast, at, latex);
    }
    const res = await validateDocxBuffer(parser.rezip());
    expect(res.ok).toBe(true);
  }, 120000);
});

describe('E4 export upgrades', () => {
  function strictParse(xml: string): Document {
    return new DOMParser({
      errorHandler: {
        warning: (w) => { throw new Error(String(w)); },
        error: (e) => { throw new Error(String(e)); },
        fatalError: (f) => { throw new Error(String(f)); },
      },
    }).parseFromString(xml, 'text/xml');
  }

  function mathFragment(html: string): string {
    const start = html.indexOf('<math');
    const end = html.lastIndexOf('</math>');
    expect(start).toBeGreaterThanOrEqual(0);
    expect(end).toBeGreaterThan(start);
    return html.slice(start, end + '</math>'.length);
  }

  it('HTML display equations render native MathML <math display> (well-formed)', async () => {
    const { parser, ast, anchor } = await fresh();
    await insertEq(parser, ast, anchor, 'x = \\frac{a}{b}');
    const html = convertAstToHtml(ast, 'E4');
    expect(html).toContain('<math');
    expect(html).toContain('display="block"');
    const frag = mathFragment(html);
    const dom = strictParse(frag);
    expect(dom.documentElement.tagName).toBe('math');
    expect(frag).toContain('<mfrac>');
  });

  it('HTML matrix equations render mtable/mtr/mtd', async () => {
    const { parser, ast, anchor } = await fresh();
    await insertEq(parser, ast, anchor, '\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}');
    const html = convertAstToHtml(ast, 'E4');
    const frag = mathFragment(html);
    strictParse(frag);
    expect(frag).toContain('<mtable>');
    expect(frag).toContain('<mtr>');
    expect(frag).toContain('<mtd>');
  });

  it('HTML inline math stays a span (no <math>)', async () => {
    const { ast } = await fresh();
    const node = ast.sections[0].nodes[0];
    const p = node.xml as Element;
    const doc = p.ownerDocument as Document;
    const runs = p.getElementsByTagName('w:r');
    expect(runs.length).toBeGreaterThan(0);
    const oMath = createMathElement(doc, 'm:oMath');
    const r = createMathElement(doc, 'm:r');
    const t = createMathElement(doc, 'm:t');
    t.appendChild(doc.createTextNode('x+1'));
    r.appendChild(t);
    oMath.appendChild(r);
    (runs[0] as Element).appendChild(oMath);
    const html = convertAstToHtml(ast, 'inline');
    expect(html).toContain('<span class="equation">');
    expect(html).not.toContain('<math');
  });

  it('Markdown display equations render $$ blocks', async () => {
    const { parser, ast, anchor } = await fresh();
    await insertEq(parser, ast, anchor, 'E = \\frac{m}{\\sqrt{c}}');
    const md = convertAstToMarkdown(ast);
    expect(md).toContain('$$');
    expect(md).toMatch(/\$\$\nE = \\frac\{m\}\{\\sqrt\{c\}\}\n\$\$/);
  });

  it('Markdown inline math is written as \\(...\\)', async () => {
    const { ast } = await fresh();
    const node = ast.sections[0].nodes[0];
    const p = node.xml as Element;
    const doc = p.ownerDocument as Document;
    const runs = p.getElementsByTagName('w:r');
    const oMath = createMathElement(doc, 'm:oMath');
    const r = createMathElement(doc, 'm:r');
    const t = createMathElement(doc, 'm:t');
    t.appendChild(doc.createTextNode('x+1'));
    r.appendChild(t);
    oMath.appendChild(r);
    (runs[0] as Element).appendChild(oMath);
    const md = convertAstToMarkdown(ast);
    expect(md).toContain('\\(x+1\\)');
    expect(md).not.toContain('$');
  });
});

describe('E4 polish (MathML arity, escaped delims, invisible delims)', () => {
  function mathFragment(html: string): string {
    const start = html.indexOf('<math');
    const end = html.lastIndexOf('</math>');
    expect(start).toBeGreaterThanOrEqual(0);
    expect(end).toBeGreaterThan(start);
    return html.slice(start, end + '</math>'.length);
  }

  function mathDoc(html: string): Document {
    return domOf(mathFragment(html));
  }

  function childEls(el: Element): Element[] {
    const out: Element[] = [];
    for (let i = 0; i < el.childNodes.length; i++) {
      const ch = el.childNodes[i] as Element;
      if (ch && (ch as unknown as Node).nodeType === 1) out.push(ch as Element);
    }
    return out;
  }

  it('multi-token numerator mfrac has exactly 2 mrow children', async () => {
    const { parser, ast, anchor } = await fresh();
    await insertEq(parser, ast, anchor, 'x = \\frac{a+b}{c}');
    const html = convertAstToHtml(ast, 'E4');
    const dom = mathDoc(html);
    const fracs = dom.getElementsByTagName('mfrac');
    expect(fracs.length).toBe(1);
    const kids = childEls(fracs[0] as unknown as Element);
    expect(kids.length).toBe(2);
    expect(kids[0].tagName).toBe('mrow');
    expect(kids[1].tagName).toBe('mrow');
  });

  it('nary operand sits outside munderover (munderover has exactly op/sub/sup)', async () => {
    const { parser, ast, anchor } = await fresh();
    await insertEq(parser, ast, anchor, '\\sum_{i=1}^{n} a_i');
    const html = convertAstToHtml(ast, 'E4');
    const dom = mathDoc(html);
    const uovers = dom.getElementsByTagName('munderover');
    expect(uovers.length).toBe(1);
    const kids = childEls(uovers[0] as unknown as Element);
    expect(kids.length).toBe(3);
    expect(kids[0].tagName).toBe('mo');
    // The integrand follows the munderover inside an outer mrow.
    const parent = (uovers[0] as unknown as Element).parentNode as Element;
    expect(parent.tagName).toBe('mrow');
    const siblings = childEls(parent);
    expect(siblings.length).toBe(2);
    expect(siblings[0].tagName).toBe('munderover');
    expect(siblings[1].tagName).toBe('mrow');
  });

  it('\\{ \\} delimiters convert to brace begChr/endChr', async () => {
    const { parser, ast, anchor } = await fresh();
    const id = await insertEq(parser, ast, anchor, '\\left\\{ x \\right\\}');
    const xml = xmlOf(parser, id);
    expect(count(xml, 'm:d')).toBe(1);
    expect(xml).toContain('m:val="{"');
    expect(xml).toContain('m:val="}"');
    expect(allRunText(xml)).not.toContain('\\');
  });

  it('invisible delimiter emits explicit empty m:begChr/m:endChr attrs', async () => {
    const { parser, ast, anchor } = await fresh();
    const id = await insertEq(parser, ast, anchor, '\\left. x \\right)');
    const xml = xmlOf(parser, id);
    expect(count(xml, 'm:d')).toBe(1);
    expect(xml).toMatch(/<m:begChr m:val=""\/>/);
    expect(xml).toContain('m:val=")"');
    const id2 = await insertEq(parser, ast, anchor, '\\left( x \\right.');
    const xml2 = xmlOf(parser, id2);
    expect(xml2).toMatch(/<m:endChr m:val=""\/>/);
    expect(xml2).toContain('m:val="("');
  });
});

describe('R-41 same-anchor after inserts stay FIFO (equation + callout)', () => {
  function astIds(ast: any): string[] {
    return ast.sections.flatMap((s: any) => s.nodes.map((n: any) => n.id));
  }

  function astTexts(ast: any): string[] {
    return ast.sections.flatMap((s: any) => s.nodes.map((n: any) => String(n.metadata?.text ?? '')));
  }

  function domIdsFrom(anchorXml: Element): string[] {
    const ids: string[] = [];
    let curr: Element | null = anchorXml;
    while (curr) {
      if (curr.nodeType === 1 && curr.tagName !== 'w:sectPr') {
        const id = curr.getAttribute('data-astra-id');
        if (id) ids.push(id);
      }
      curr = curr.nextSibling as Element | null;
    }
    return ids;
  }

  it('3 same-anchor equation inserts land in call order (AST + DOM)', async () => {
    const { parser, ast, anchor } = await fresh();
    const trailing = ast.sections[0].nodes[1].id as string;
    const a = await insertEq(parser, ast, anchor, 'EQ-A');
    const b = await insertEq(parser, ast, anchor, 'EQ-B');
    const c = await insertEq(parser, ast, anchor, 'EQ-C');
    expect(astIds(ast)).toEqual([anchor, a, b, c, trailing]);
    expect(astTexts(ast).slice(1, 4)).toEqual(['EQ-A', 'EQ-B', 'EQ-C']);
    expect(domIdsFrom(ast.sections[0].nodes[0].xml)).toEqual([anchor, a, b, c, trailing]);
  });

  it('equation delete-middle then insert lands after the live tail', async () => {
    const { parser, ast, anchor } = await fresh();
    const trailing = ast.sections[0].nodes[1].id as string;
    const a = await insertEq(parser, ast, anchor, 'EQ-A');
    const b = await insertEq(parser, ast, anchor, 'EQ-B');
    const c = await insertEq(parser, ast, anchor, 'EQ-C');
    executeDelete({ action: 'delete', target: b } as never, ast, parser);
    expect(astIds(ast)).toEqual([anchor, a, c, trailing]);
    const d = await insertEq(parser, ast, anchor, 'EQ-D');
    expect(astIds(ast)).toEqual([anchor, a, c, d, trailing]);
    expect(domIdsFrom(ast.sections[0].nodes[0].xml)).toEqual([anchor, a, c, d, trailing]);
  });

  it('equation single after / before inserts are unchanged', async () => {
    const { parser, ast, anchor } = await fresh();
    const trailing = ast.sections[0].nodes[1].id as string;
    const solo = await insertEq(parser, ast, anchor, 'solo');
    expect(astIds(ast)).toEqual([anchor, solo, trailing]);
    const { parser: p2, ast: a2, anchor: an2 } = await fresh();
    const trail2 = a2.sections[0].nodes[1].id as string;
    const bef = (await executeEquation(
      { action: 'equation', operation: 'insert', anchor: an2, latex: 'bef', position: 'before' } as never,
      a2,
      p2
    )) as { id: string };
    expect(astIds(a2)).toEqual([bef.id, an2, trail2]);
    void parser;
  });

  it('3 same-anchor callout inserts land in call order (AST + DOM)', async () => {
    const { parser, ast, anchor } = await fresh();
    const trailing = ast.sections[0].nodes[1].id as string;
    await executeCallout({ action: 'callout', anchor, position: 'after', text: 'C-A' } as never, ast, parser);
    await executeCallout({ action: 'callout', anchor, position: 'after', text: 'C-B' } as never, ast, parser);
    await executeCallout({ action: 'callout', anchor, position: 'after', text: 'C-C' } as never, ast, parser);
    expect(astTexts(ast).slice(1, 4)).toEqual(['C-A', 'C-B', 'C-C']);
    const ids = astIds(ast);
    expect(ids.length).toBe(5);
    expect(ids[0]).toBe(anchor);
    expect(ids[4]).toBe(trailing);
    expect(domIdsFrom(ast.sections[0].nodes[0].xml)).toEqual(ids);
  });

  it('callout delete-middle then insert lands after the live tail', async () => {
    const { parser, ast, anchor } = await fresh();
    const trailing = ast.sections[0].nodes[1].id as string;
    await executeCallout({ action: 'callout', anchor, position: 'after', text: 'C-A' } as never, ast, parser);
    await executeCallout({ action: 'callout', anchor, position: 'after', text: 'C-B' } as never, ast, parser);
    await executeCallout({ action: 'callout', anchor, position: 'after', text: 'C-C' } as never, ast, parser);
    const mid = ast.sections[0].nodes.find((n: any) => String(n.metadata?.text) === 'C-B')!.id as string;
    executeDelete({ action: 'delete', target: mid } as never, ast, parser);
    expect(astTexts(ast)).not.toContain('C-B');
    await executeCallout({ action: 'callout', anchor, position: 'after', text: 'C-D' } as never, ast, parser);
    expect(astTexts(ast).slice(1, 4)).toEqual(['C-A', 'C-C', 'C-D']);
    const ids = astIds(ast);
    expect(ids[0]).toBe(anchor);
    expect(ids[ids.length - 1]).toBe(trailing);
    expect(domIdsFrom(ast.sections[0].nodes[0].xml)).toEqual(ids);
  });

  it('callout single after insert is unchanged', async () => {
    const { parser, ast, anchor } = await fresh();
    const trailing = ast.sections[0].nodes[1].id as string;
    await executeCallout({ action: 'callout', anchor, position: 'after', text: 'solo' } as never, ast, parser);
    expect(astTexts(ast).slice(1, 2)).toEqual(['solo']);
    const ids = astIds(ast);
    expect(ids).toHaveLength(3);
    expect(ids[0]).toBe(anchor);
    expect(ids[2]).toBe(trailing);
  });
});

describe('callout box rows stay together across pages', () => {
  it('callout row carries cantSplit from creation and validates', async () => {
    const { parser, ast, anchor } = await fresh();
    await executeCallout({ action: 'callout', anchor, position: 'after', text: 'Stay together' } as never, ast, parser);
    const node = ast.sections[0].nodes.find((n: any) => n.type === 'table') as any;
    expect(node).toBeDefined();
    const tblDoc = domOf(xmlOf(parser, node.id));
    const trs = tblDoc.getElementsByTagName('w:tr');
    expect(trs.length).toBe(1);
    const trPr = (trs[0] as Element).getElementsByTagName('w:trPr')[0] as Element | undefined;
    expect(trPr).toBeDefined();
    expect(trPr!.getElementsByTagName('w:cantSplit').length).toBe(1);
    expect((trs[0] as Element).firstChild?.nodeName).toBe('w:trPr');
    const res = await validateDocxBuffer(parser.rezip());
    expect(res.ok).toBe(true);
    expect(res.errors).toEqual([]);
  });
});

describe('dots symbol expansion', () => {
  it('converts dots to ellipsis in equation', async () => {
    const { parser, ast, anchor } = await fresh();
    const id = await insertEq(parser, ast, anchor, 'a_1, a_2, \\dots, a_n');
    const xml = xmlOf(parser, id);
    expect(allRunText(xml)).toContain('…');
  });
});

describe('extended math function names as upright literal runs', () => {
  const funcs = [
    '\\max(0, x)',
    '\\min_i x_i',
    '\\sup x',
    '\\inf x',
    '\\exp(x)',
    '\\arg\\max x',
    '\\gcd(a, b)',
    '\\arctan x',
  ];
  it.each(funcs)('inserts %s without throwing and emits m:lit', async (latex) => {
    const { parser, ast, anchor } = await fresh();
    const id = await insertEq(parser, ast, anchor, latex);
    const xml = xmlOf(parser, id);
    expect(count(xml, 'm:lit')).toBeGreaterThan(0);
  });

  it('preserves deg as degree symbol', async () => {
    const { parser, ast, anchor } = await fresh();
    const id = await insertEq(parser, ast, anchor, '90\\deg');
    const xml = xmlOf(parser, id);
    expect(allRunText(xml)).toContain('°');
  });
});

describe('subscript and superscript operands followed by operators or punctuation', () => {
  const cases = [
    '\\alpha a^2+b^2',
    '\\alpha a_1+b_2',
    '\\alpha x^2=y^2',
    '\\alpha x_1,y_2',
    '\\alpha x_i;y_j',
    '\\alpha x_1-y_2',
    '\\alpha x^10+y^2',
    '\\alpha a_1*b_2',
    '\\alpha x_1)W_2',
    '\\alpha (x_1)W_2',
    '\\mathrm{f}(x_1)W_2',
  ];

  it.each(cases)('inserts %s and validates docx buffer', async (latex) => {
    const { parser, ast, anchor } = await fresh();
    await insertEq(parser, ast, anchor, latex);
    const res = await validateDocxBuffer(parser.rezip());
    expect(res.ok).toBe(true);
    expect(res.errors).toEqual([]);
  });

  it('structure check for a^2+b^2', async () => {
    const { parser, ast, anchor } = await fresh();
    const id = await insertEq(parser, ast, anchor, '\\alpha a^2+b^2');
    const xml = xmlOf(parser, id);
    expect(count(xml, 'm:sSup')).toBe(2);
    const text = allRunText(xml);
    expect(text).toContain('a');
    expect(text).toContain('2');
    expect(text).toContain('+b');
  });

  it('structure check for x_1,y_2', async () => {
    const { parser, ast, anchor } = await fresh();
    const id = await insertEq(parser, ast, anchor, '\\alpha x_1,y_2');
    const xml = xmlOf(parser, id);
    expect(count(xml, 'm:sSub')).toBe(2);
  });

  it('unchanged behavior for spaced exponent x^10 + y', async () => {
    const { parser, ast, anchor } = await fresh();
    const id = await insertEq(parser, ast, anchor, '\\alpha x^10 + y');
    const xml = xmlOf(parser, id);
    const doc = domOf(xml);
    const supEl = doc.getElementsByTagName('m:sup')[0];
    expect(supEl).toBeDefined();
    const tNodes = supEl.getElementsByTagName('m:t');
    let supText = '';
    for (let i = 0; i < tNodes.length; i++) supText += tNodes[i].textContent || '';
    expect(supText).toBe('10');
  });

  it('unchanged behavior for spaced exponent a^2 + b', async () => {
    const { parser, ast, anchor } = await fresh();
    const id = await insertEq(parser, ast, anchor, '\\alpha a^2 + b');
    const xml = xmlOf(parser, id);
    const doc = domOf(xml);
    const supEl = doc.getElementsByTagName('m:sup')[0];
    expect(supEl).toBeDefined();
    const tNodes = supEl.getElementsByTagName('m:t');
    let supText = '';
    for (let i = 0; i < tNodes.length; i++) supText += tNodes[i].textContent || '';
    expect(supText).toBe('2');
  });

  it('unchanged behavior for n-ary integral integrand f(x)dx', async () => {
    const { parser, ast, anchor } = await fresh();
    const id = await insertEq(parser, ast, anchor, '\\int_0^1 f(x)dx');
    const xml = xmlOf(parser, id);
    expect(count(xml, 'm:nary')).toBe(1);
    expect(allRunText(xml)).toContain('f(x)dx');
  });

  it('unchanged behavior for group exponent x^(n)+1', async () => {
    const { parser, ast, anchor } = await fresh();
    const id = await insertEq(parser, ast, anchor, '\\alpha x^(n)+1');
    const xml = xmlOf(parser, id);
    const doc = domOf(xml);
    const supEl = doc.getElementsByTagName('m:sup')[0];
    expect(supEl).toBeDefined();
    const tNodes = supEl.getElementsByTagName('m:t');
    let supText = '';
    for (let i = 0; i < tNodes.length; i++) supText += tNodes[i].textContent || '';
    expect(supText).toBe('(n)');
  });
});

describe('operatorname upright named operators', () => {
  it('converts operatorname with subscript and argument', async () => {
    const { parser, ast, anchor } = await fresh();
    const id = await insertEq(parser, ast, anchor, '\\operatorname{argmax}_x f(x)');
    const xml = xmlOf(parser, id);
    expect(count(xml, 'm:lit')).toBeGreaterThan(0);
    expect(allRunText(xml)).toContain('argmax');
  });

  it('converts operatorname with mixed case letters', async () => {
    const { parser, ast, anchor } = await fresh();
    const id = await insertEq(parser, ast, anchor, '\\operatorname{ReLU}(x)');
    const xml = xmlOf(parser, id);
    expect(count(xml, 'm:lit')).toBeGreaterThan(0);
    expect(allRunText(xml)).toContain('ReLU');
  });

  it('fails loudly when operatorname contains non-letter characters', async () => {
    const { parser, ast, anchor } = await fresh();
    await expect(insertEq(parser, ast, anchor, '\\operatorname{a1}')).rejects.toThrow(/operatorname/);
  });
});

describe('single-token superscript and subscript parsing', () => {
  function scriptPairs(xml: string, tag: 'm:sSup' | 'm:sSub'): Array<{ base: string; script: string }> {
    const doc = domOf(xml);
    const elements = doc.getElementsByTagName(tag);
    const pairs: Array<{ base: string; script: string }> = [];
    const scriptTag = tag === 'm:sSup' ? 'm:sup' : 'm:sub';
    for (let i = 0; i < elements.length; i++) {
      const el = elements[i];
      const e = el.getElementsByTagName('m:e')[0];
      const s = el.getElementsByTagName(scriptTag)[0];
      const baseText = e ? Array.from(e.getElementsByTagName('m:t')).map((t) => (t as Element).textContent || '').join('') : '';
      const scriptText = s ? Array.from(s.getElementsByTagName('m:t')).map((t) => (t as Element).textContent || '').join('') : '';
      pairs.push({ base: baseText, script: scriptText });
    }
    return pairs;
  }

  const tokenCases: Array<{
    latex: string;
    sups?: Array<{ base: string; script: string }>;
    subs?: Array<{ base: string; script: string }>;
    expectedText: string;
  }> = [
    {
      latex: '\\alpha x^2y',
      sups: [{ base: 'x', script: '2' }],
      subs: [],
      expectedText: 'αx2y',
    },
    {
      latex: '\\alpha x_ij',
      sups: [],
      subs: [{ base: 'x', script: 'i' }],
      expectedText: 'αxij',
    },
    {
      latex: '\\alpha e^ax+1',
      sups: [{ base: 'e', script: 'a' }],
      subs: [],
      expectedText: 'αeax+1',
    },
    {
      latex: '\\alpha x^10',
      sups: [{ base: 'x', script: '10' }],
      subs: [],
      expectedText: 'αx10',
    },
    {
      latex: '\\alpha x^2.5y',
      sups: [{ base: 'x', script: '2.5' }],
      subs: [],
      expectedText: 'αx2.5y',
    },
    {
      latex: '\\alpha x^-1+y',
      sups: [{ base: 'x', script: '-1' }],
      subs: [],
      expectedText: 'αx-1+y',
    },
    {
      latex: '\\alpha x^{ab}',
      sups: [{ base: 'x', script: 'ab' }],
      subs: [],
      expectedText: 'αxab',
    },
    {
      latex: '\\alpha a^2+b^2',
      sups: [
        { base: 'a', script: '2' },
        { base: '+b', script: '2' },
      ],
      subs: [],
      expectedText: 'αa2+b2',
    },
    {
      latex: '\\alpha (x_1)W_2',
      sups: [],
      subs: [
        { base: '(x', script: '1' },
        { base: ')W', script: '2' },
      ],
      expectedText: 'α(x1)W2',
    },
  ];

  it.each(tokenCases)('parses single TeX token for $latex and validates docx', async (tc) => {
    const { parser, ast, anchor } = await fresh();
    const id = await insertEq(parser, ast, anchor, tc.latex);
    const xml = xmlOf(parser, id);

    if (tc.sups !== undefined) {
      expect(scriptPairs(xml, 'm:sSup')).toEqual(tc.sups);
    }
    if (tc.subs !== undefined) {
      expect(scriptPairs(xml, 'm:sSub')).toEqual(tc.subs);
    }

    const text = allRunText(xml);
    const expectedFromInput = tc.latex.replace(/\\alpha\s*/g, 'α').replace(/[\^_{}]/g, '');
    expect(text).toBe(expectedFromInput);
    expect(text).toBe(tc.expectedText);

    await assertValidDocx(parser.rezip());
  });

  it('structure check: a^2+b^2 has two m:sSup', async () => {
    const { parser, ast, anchor } = await fresh();
    const id = await insertEq(parser, ast, anchor, '\\alpha a^2+b^2');
    const xml = xmlOf(parser, id);
    expect(count(xml, 'm:sSup')).toBe(2);
  });

  it('structure check: (x_1)W_2 has two m:sSub', async () => {
    const { parser, ast, anchor } = await fresh();
    const id = await insertEq(parser, ast, anchor, '\\alpha (x_1)W_2');
    const xml = xmlOf(parser, id);
    expect(count(xml, 'm:sSub')).toBe(2);
  });
});


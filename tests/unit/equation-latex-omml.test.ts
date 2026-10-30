/**
 * Phase E1 — equation convert-then-error.
 *
 * Graduated LaTeX→OMML conversion in the equation op path:
 *  - converted showcase set (frac/sqrt/sup/sub/nary/delimiters/symbols incl
 *    the item-8 Maxwell case) renders structurally + validates;
 *  - loud-failure cases (align, exotic macro, over-depth nesting) throw
 *    E_VALIDATION naming the construct + remedy, emitting nothing;
 *    (E4 update: matrix environments now CONVERT (see equation-omml-e4);
 *    the depth bound is now 6, so the over-depth breaker is 7-deep.)
 *  - plain-text math (no backslashes) keeps verbatim behavior unchanged;
 *  - every conversion passes the validator gate (invalid output never emitted).
 */
import { describe, it, expect } from 'vitest';
import { DOMParser } from '@xmldom/xmldom';
import {
  DocxParser,
  createBlankDocxBuffer,
  executeEquation,
} from '@attentionform/astra-engine';
import { validateDocxBuffer } from '../../packages/engine/src/ooxml/validator-gate';

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

describe('E1 converted showcase set (structural OMML, validates)', () => {
  const cases: Array<{ name: string; latex: string; check: (xml: string) => void }> = [
    {
      name: 'frac + nested sqrt + sup + pm (quadratic)',
      latex: 'x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}',
      check: (xml) => {
        expect(count(xml, 'm:f')).toBe(1);
        expect(count(xml, 'm:rad')).toBe(1);
        expect(count(xml, 'm:sSup')).toBeGreaterThanOrEqual(1);
        expect(xml).toContain('±');
        expect(allRunText(xml)).not.toContain('\\');
      },
    },
    {
      name: 'sqrt + nth root',
      latex: '\\sqrt{a^2 + b^2} = c,\\ \\sqrt[3]{x}',
      check: (xml) => {
        expect(count(xml, 'm:rad')).toBe(2);
        expect(allRunText(xml)).toContain('3');
        expect(allRunText(xml)).not.toContain('\\');
      },
    },
    {
      name: 'sub / sup / subSup',
      latex: 'Q_{\\alpha}^{2} + x_{i}',
      check: (xml) => {
        expect(count(xml, 'm:sSubSup')).toBe(1);
        expect(count(xml, 'm:sSub')).toBe(1);
        expect(allRunText(xml)).toContain('α');
        expect(allRunText(xml)).not.toContain('\\');
      },
    },
    {
      name: 'n-ary sum (undOvr)',
      latex: '\\sum_{i=1}^{n} a_i',
      check: (xml) => {
        expect(count(xml, 'm:nary')).toBe(1);
        expect(xml).toContain('m:val="∑"');
        expect(xml).toContain('m:val="undOvr"');
        expect(count(xml, 'm:sub')).toBeGreaterThanOrEqual(1);
        expect(count(xml, 'm:sup')).toBeGreaterThanOrEqual(1);
        expect(allRunText(xml)).not.toContain('\\');
      },
    },
    {
      name: 'n-ary integral (subSup)',
      latex: '\\int_{a}^{b} f(x) dx',
      check: (xml) => {
        expect(count(xml, 'm:nary')).toBe(1);
        expect(xml).toContain('m:val="∫"');
        expect(xml).toContain('m:val="subSup"');
        expect(allRunText(xml)).not.toContain('\\');
      },
    },
    {
      name: 'delimiters (\\left/\\right + bars)',
      latex: '\\left( \\frac{a}{b} \\right) + |c|',
      check: (xml) => {
        expect(count(xml, 'm:d')).toBe(2);
        expect(count(xml, 'm:f')).toBe(1);
        expect(xml).toContain('m:val="("');
        expect(xml).toContain('m:val=")"');
        expect(xml).toContain('m:val="|"');
        expect(allRunText(xml)).not.toContain('\\');
      },
    },
    {
      name: 'symbol table',
      latex: '\\nabla \\cdot \\alpha \\le \\beta \\ge \\gamma \\neq \\pi \\approx \\infty',
      check: (xml) => {
        for (const sym of ['∇', '⋅', 'α', '≤', 'β', '≥', 'γ', '≠', 'π', '≈', '∞']) {
          expect(allRunText(xml)).toContain(sym);
        }
        expect(allRunText(xml)).not.toContain('\\');
      },
    },
    {
      name: 'named functions as roman lit runs + limit',
      latex: '\\lim_{x \\to 0} \\frac{\\sin x}{x} = 1',
      check: (xml) => {
        expect(count(xml, 'm:f')).toBe(1);
        expect(xml).toContain('<m:lit');
        expect(allRunText(xml)).toContain('lim');
        expect(allRunText(xml)).toContain('sin');
        expect(allRunText(xml)).toContain('→');
        expect(allRunText(xml)).not.toContain('\\');
      },
    },
    {
      name: 'item-8 Maxwell case',
      latex: '\\nabla \\times \\mathbf{B} = \\mu_0 \\mathbf{J} + \\mu_0 \\epsilon_0 \\frac{\\partial \\mathbf{E}}{\\partial t}',
      check: (xml) => {
        expect(count(xml, 'm:f')).toBe(1);
        expect(count(xml, 'm:sSub')).toBeGreaterThanOrEqual(2);
        for (const sym of ['∇', '×', 'B', 'μ', 'J', 'ϵ', '∂', 'E']) {
          expect(allRunText(xml)).toContain(sym);
        }
        expect(allRunText(xml)).not.toContain('\\');
        expect(allRunText(xml)).not.toContain('mathbf');
      },
    },
    {
      name: 'x/y nesting',
      latex: '\\frac{x + \\frac{1}{y}}{2}',
      check: (xml) => {
        expect(count(xml, 'm:f')).toBe(2);
        expect(allRunText(xml)).not.toContain('\\');
      },
    },
  ];

  it.each(cases.map((c) => [c.name, c] as const))(
    'converts %s',
    async (_name, c) => {
      const { parser, ast, anchor } = await fresh();
      const out = (await executeEquation(
        { action: 'equation', operation: 'insert', anchor, latex: c.latex } as never,
        ast,
        parser
      )) as { id: string };
      const created = ast.sections[0].nodes.find((n) => n.id === out.id);
      expect(created).toBeDefined();
      expect(created!.type).toBe('equation');
      expect(created!.metadata.hasEquation).toBe(true);
      expect(created!.metadata.text).toBe(c.latex);
      const xml = xmlOf(parser, out.id);
      expect(xml).toContain('m:oMath');
      c.check(xml);
    }
  );

  it('combined showcase document passes the real OOXML validator', async () => {
    const { parser, ast, anchor } = await fresh();
    let at = anchor;
    const lats = [
      'x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}',
      '\\int_{-\\infty}^{\\infty} e^{-x^2} dx = \\sqrt{\\pi}',
      '\\sum_{i=1}^{n} a_i',
      '\\nabla \\times \\mathbf{B} = \\mu_0 \\mathbf{J} + \\mu_0 \\epsilon_0 \\frac{\\partial \\mathbf{E}}{\\partial t}',
      '\\left( \\frac{a}{b} \\right) + |c|',
    ];
    for (const latex of lats) {
      const out = (await executeEquation(
        { action: 'equation', operation: 'insert', anchor: at, latex } as never,
        ast,
        parser
      )) as { id: string };
      at = out.id;
    }
    const res = await validateDocxBuffer(parser.rezip());
    expect(res.ok).toBe(true);
  }, 120000);

  it('replace path converts structurally and keeps pPr', async () => {
    const { parser, ast } = await fresh();
    const target = ast.sections[0].nodes[0].id;
    await executeEquation(
      { action: 'equation', operation: 'replace', target, latex: 'E = \\frac{m}{\\sqrt{c}}' } as never,
      ast,
      parser
    );
    const xml = xmlOf(parser, target);
    expect(count(xml, 'm:f')).toBe(1);
    expect(count(xml, 'm:rad')).toBe(1);
    expect(allRunText(xml)).not.toContain('\\');
  });
  it('text/mbox groups preserve inner spaces (xml:space)', async () => {
    const { parser, ast, anchor } = await fresh();
    const out = (await executeEquation(
      { action: 'equation', operation: 'insert', anchor, latex: '\\text{a b} + \\mbox{c d}' } as never,
      ast,
      parser
    )) as { id: string };
    const xml = xmlOf(parser, out.id);
    expect(allRunText(xml)).toContain('a b');
    expect(allRunText(xml)).toContain('c d');
    expect(xml).toContain('xml:space="preserve"');
    expect(allRunText(xml)).not.toContain('\\');
    expect(allRunText(xml)).not.toContain('text');
    expect(allRunText(xml)).not.toContain('mbox');
  });
});

describe('E1 loud failures (E_VALIDATION + remedy, nothing emitted)', () => {
  it.each([
    ['alignment environment', '\\begin{align} x \\\\ y \\end{align}', /align/],
    ['exotic macro', '\\oint_{C} x dx', /\\oint/],
    ['over-depth nesting', '\\frac{a}{\\frac{b}{\\frac{c}{\\frac{d}{\\frac{e}{\\frac{f}{\\frac{g}{h}}}}}}}', /nest|depth/i],
    ['unbalanced input', '\\frac{a}{b', /brace/i],
    ['accent macro without an m:acc mapping', '\\overrightarrow{x}', /\\overrightarrow/],
    ['stray bracket', '\\alpha]', /bracket/],
  ])('%s fails loudly', async (_name, latex, namesConstruct) => {
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
    expect(err).not.toBeNull();
    const msg = (err as Error).message;
    expect(msg).toMatch(/E_VALIDATION/);
    expect(msg).toMatch(namesConstruct);
    expect(msg).toMatch(/position \d+/);
    expect(msg).toMatch(/Remedy:/);
    expect(msg).toMatch(/simplify|split|pre-built OMML/);
    expect((err as unknown as { code?: string }).code).toBe('E_VALIDATION');
    // Nothing emitted: AST + document XML untouched.
    expect(ast.sections[0].nodes.length).toBe(before);
    expect(parser.readDocumentXml()).toBe(docBefore);
  });

  it('failed replace leaves the target paragraph intact', async () => {
    const { parser, ast } = await fresh();
    const target = ast.sections[0].nodes[0].id;
    const before = xmlOf(parser, target);
    await expect(
      executeEquation(
        { action: 'equation', operation: 'replace', target, latex: '\\begin{align} x \\end{align}' } as never,
        ast,
        parser
      )
    ).rejects.toThrow(/E_VALIDATION/);
    expect(xmlOf(parser, target)).toBe(before);
  });
});

describe('E1 plain-text math unchanged (verbatim)', () => {
  it.each([['x+y=z']])('stores %s verbatim', async (latex) => {
    const { parser, ast, anchor } = await fresh();
    const out = (await executeEquation(
      { action: 'equation', operation: 'insert', anchor, latex } as never,
      ast,
      parser
    )) as { id: string };
    const xml = xmlOf(parser, out.id);
    expect(xml).toContain(latex);
    expect(count(xml, 'm:r')).toBe(1);
    expect(count(xml, 'm:t')).toBe(1);
    const node = ast.sections[0].nodes.find((n) => n.id === out.id);
    expect(node!.metadata.text).toBe(latex);
  });
});

describe('E1 safety gates intact', () => {
  it('XXE / script input still rejected before conversion', async () => {
    const { ast, anchor } = await fresh();
    await expect(
      executeEquation(
        { action: 'equation', operation: 'insert', anchor, latex: '<!ENTITY x>' } as never,
        ast
      )
    ).rejects.toThrow(/XXE/);
    await expect(
      executeEquation(
        { action: 'equation', operation: 'insert', anchor, latex: '\\frac{<script>}{x}' } as never,
        ast
      )
    ).rejects.toThrow(/script/i);
  });
});

describe('equation mode', () => {
  async function insert(latex: string, mode?: 'auto' | 'latex' | 'plain') {
    const { parser, ast, anchor } = await fresh();
    const out = (await executeEquation(
      { action: 'equation', operation: 'insert', anchor, latex, ...(mode ? { mode } : {}) } as never,
      ast,
      parser
    )) as { id: string };
    return { parser, xml: xmlOf(parser, out.id) };
  }

  it('auto converts E=mc^2, a^2+b^2=c^2 and x_1 into real sub/superscripts', async () => {
    const a = await insert('E=mc^2');
    expect(count(a.xml, 'm:sSup')).toBe(1);
    expect(allRunText(a.xml)).not.toContain('^');
    const b = await insert('a^2+b^2=c^2');
    expect(count(b.xml, 'm:sSup')).toBe(3);
    const c = await insert('x_1');
    expect(count(c.xml, 'm:sSub')).toBe(1);
    for (const r of [a, b, c]) expect((await validateDocxBuffer(r.parser.rezip())).ok).toBe(true);
  });

  it('auto keeps text without a backslash, ^ or _ as literal text', async () => {
    const { xml } = await insert('x+y=z');
    expect(count(xml, 'm:t')).toBe(1);
    expect(allRunText(xml)).toBe('x+y=z');
  });

  it("mode 'plain' keeps any text verbatim in one run", async () => {
    for (const latex of ['E=mc^2', 'a^2+b^2=c^2', 'snake_case']) {
      const { xml } = await insert(latex, 'plain');
      expect(count(xml, 'm:t')).toBe(1);
      expect(allRunText(xml)).toBe(latex);
    }
  });

  it("auto keeps 'x^' verbatim, mode 'latex' refuses it, and a backslash formula that fails is loud", async () => {
    const { xml } = await insert('x^');
    expect(allRunText(xml)).toBe('x^');
    await expect(insert('x^', 'latex')).rejects.toThrow(/E_VALIDATION/);
    await expect(insert('\\alpha x^')).rejects.toThrow(/E_VALIDATION/);
  });

  it("mode 'latex' converts text without a backslash", async () => {
    const { xml } = await insert('a_1 + b_2', 'latex');
    expect(count(xml, 'm:sSub')).toBe(2);
  });

  it('replace honors the mode', async () => {
    const { parser, ast, anchor } = await fresh();
    await executeEquation({ action: 'equation', operation: 'replace', target: anchor, latex: 'snake_case', mode: 'plain' } as never, ast, parser);
    expect(allRunText(xmlOf(parser, anchor))).toBe('snake_case');
    await executeEquation({ action: 'equation', operation: 'replace', target: anchor, latex: 'x_1' } as never, ast, parser);
    expect(count(xmlOf(parser, anchor), 'm:sSub')).toBe(1);
  });
});

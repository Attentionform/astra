import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, validateDocxBuffer, convertAstToMarkdown } from '@attentionform/astra-engine';
import { compileMarkdown, composeDocument } from '../../packages/engine/src/compose/markdown';
import { lexMarkdown } from '../../packages/engine/src/compose/math-extensions';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

type Op = Record<string, unknown>;
const ops = (md: string, dollarMath = false): Op[] => compileMarkdown(md, dollarMath ? { dollarMath: true } : undefined) as unknown as Op[];
const spansOf = (list: Op[]): Array<Record<string, unknown>> =>
  list.flatMap((o) => ((o['params'] as { mdSpans?: Array<Record<string, unknown>> } | undefined)?.mdSpans ?? []));

function compileError(md: string, dollarMath = false): string {
  try {
    compileMarkdown(md, dollarMath ? { dollarMath: true } : undefined);
  } catch (err: unknown) {
    return err instanceof Error ? err.message : String(err);
  }
  return '';
}

async function compose(md: string, dollarMath = false): Promise<{ parser: DocxParser; xml: string; bytes: Buffer }> {
  const parser = DocxParser.createNewDocument('blank');
  const res = await composeDocument(parser, md, dollarMath ? { dollarMath: true } : undefined);
  expect(res.success).toBe(true);
  const bytes = parser.rezip();
  return { parser, xml: new AdmZip(bytes).readAsText('word/document.xml'), bytes };
}

const count = (xml: string, tag: string): number => (xml.match(new RegExp(`<${tag}[ >/]`, 'g')) ?? []).length;

describe('inline math compiles to spans', () => {
  it('splits a sentence into text, math, text and keeps the plain text readable', () => {
    const list = ops('Let \\(d_k\\) be the value.\n');
    const p = list.find((o) => o['action'] === 'insert' && o['type'] === 'paragraph')!;
    expect(p['text']).toBe('Let \\(d_k\\) be the value.');
    expect(spansOf(list).map((s) => [s['text'], s['math'] ?? false])).toEqual([
      ['Let ', false],
      ['d_k', true],
      [' be the value.', false],
    ]);
  });

  it('reads two formulas in one line and a formula inside bold text', () => {
    expect(spansOf(ops('a \\(x\\) b \\(y\\) c\n')).filter((s) => s['math'])).toHaveLength(2);
    expect(spansOf(ops('**bold \\(x\\) here**\n')).filter((s) => s['math'])).toHaveLength(1);
  });

  it('keeps asterisks and underscores inside math out of emphasis', () => {
    expect(spansOf(ops('\\(a*b*c\\) and \\(a_b_c\\)\n')).filter((s) => s['math']).map((s) => s['text'])).toEqual([
      'a*b*c',
      'a_b_c',
    ]);
  });

  it('leaves code spans and escaped parentheses alone', () => {
    expect(spansOf(ops('code `\\(x\\)` span\n')).some((s) => s['math'])).toBe(false);
    expect(spansOf(ops('esc \\\\(x\\\\) and foo\\(bar\n')).some((s) => s['math'])).toBe(false);
  });

  it('leaves unclosed, empty and over-long parenthesis math as text', () => {
    expect(spansOf(ops('unclosed \\(x and more\n')).some((s) => s['math'])).toBe(false);
    expect(spansOf(ops('empty \\(\\) here\n')).some((s) => s['math'])).toBe(false);
    expect(spansOf(ops(`long \\(${'x'.repeat(600)}\\) here\n`)).some((s) => s['math'])).toBe(false);
  });
});

describe('dollar math is off unless asked for', () => {
  const PROSE = ['It costs $5 and $10 today.', 'US$5-$10', 'path $HOME/$USER', 'price $5.00 and $x$ here'];
  it('treats prices and shell text as plain text by default', () => {
    for (const line of PROSE) expect(spansOf(ops(`${line}\n`)).some((s) => s['math'])).toBe(false);
  });

  it('turns $...$ into math only with the option, and still skips prices', () => {
    expect(spansOf(ops('$x^2$ and $y_1$\n', true)).filter((s) => s['math']).map((s) => s['text'])).toEqual(['x^2', 'y_1']);
    for (const line of ['It costs $5 and $10 today.', 'US$5-$10', '$ x $', '$x $', 'cost $a$5']) {
      expect(spansOf(ops(`${line}\n`, true)).some((s) => s['math']), line).toBe(false);
    }
  });

  it('does not let $...$ run across lines or past 500 characters', () => {
    expect(spansOf(ops('$a\nb$\n', true)).some((s) => s['math'])).toBe(false);
    expect(spansOf(ops(`$${'x'.repeat(600)}$\n`, true)).some((s) => s['math'])).toBe(false);
  });

  it('lexes with both parsers without throwing on odd input', () => {
    for (const src of ['$', '$$', '\\(', '\\', '$\\$$', 'a\\$b$c$d', '$a\\', '\\(\\(x\\)\\)']) {
      expect(() => lexMarkdown(src)).not.toThrow();
      expect(() => lexMarkdown(src, { dollarMath: true })).not.toThrow();
    }
  });
});

describe('inline math in a document', () => {
  it('builds an m:oMath between the text runs and the document validates', async () => {
    const { xml, bytes } = await compose('Let \\(\\frac{a}{b}\\) be it.\n');
    expect(count(xml, 'm:oMath')).toBe(1);
    expect(count(xml, 'm:oMathPara')).toBe(0);
    expect(count(xml, 'm:f')).toBe(1);
    expect(xml).toMatch(/Let <\/w:t><\/w:r><m:oMath>/);
    expect(xml).toContain('be it.');
    expect((await validateDocxBuffer(bytes)).ok).toBe(true);
  });

  it('keeps default-off prose free of math', async () => {
    const { xml } = await compose('It costs $5 and $10, US$5-$10, path $HOME/$USER.\n');
    expect(count(xml, 'm:oMath')).toBe(0);
    expect(xml).toContain('$HOME/$USER');
  });

  it('builds math from $...$ with the option', async () => {
    const { xml, bytes } = await compose('We have $x^2$ here.\n', true);
    expect(count(xml, 'm:sSup')).toBe(1);
    expect((await validateDocxBuffer(bytes)).ok).toBe(true);
  });

  it('reloads as a paragraph whose text reads back with the formula', async () => {
    const { bytes } = await compose('Let \\(\\frac{a}{b}\\) be it.\n');
    const ast = await new DocxParser(bytes).getAST();
    const nodes = ast.sections.flatMap((s) => s.nodes);
    const node = nodes.find((n) => String(n.metadata?.text ?? '').includes('be it.'))!;
    expect(node.type).toBe('paragraph');
    expect(node.metadata?.hasEquation).toBe(true);
    expect(node.metadata?.text).toBe('Let \\(\\frac{a}{b}\\) be it.');
  });

  it('round-trips: export then import gives the same OMML', async () => {
    const first = await compose('Let \\(\\frac{a}{b} + x_1\\) be it.\n');
    const ast = await new DocxParser(first.bytes).getAST();
    const md = convertAstToMarkdown(ast as never);
    expect(md).toContain('Let \\(\\frac{a}{b}+x_1\\) be it.');
    const second = await compose(md);
    const oMath = (xml: string): string => /<m:oMath>[\s\S]*?<\/m:oMath>/.exec(xml)![0];
    expect(oMath(second.xml)).toBe(oMath(first.xml));
  });

  it('works inside a heading and a list item', async () => {
    const { xml, bytes } = await compose('# Title \\(x_1\\)\n\n- item \\(y^2\\)\n');
    expect(count(xml, 'm:oMath')).toBe(2);
    expect(xml).toContain('Heading1');
    expect((await validateDocxBuffer(bytes)).ok).toBe(true);
  });
});

describe('inline math errors name the line and change nothing', () => {
  it('refuses math inside a link, a table cell and an alert', () => {
    const link = compileError('x\n\n[see \\(x\\)](https://example.com)\n');
    expect(link).toContain('E_VALIDATION');
    expect(link).toContain('line 3');
    expect(link).toContain('inline math inside a link');
    const cell = compileError('| a | b |\n|---|---|\n| \\(x\\) | 2 |\n');
    expect(cell).toContain('inline math inside tables');
    const cellDollar = compileError('| a | b |\n|---|---|\n| $x$ | 2 |\n', true);
    expect(cellDollar).toContain('inline math inside tables');
    const alert = compileError('> [!NOTE]\n> See \\(x_1\\) here\n');
    expect(alert).toContain('inline math inside alerts');
    expect(alert).toContain('line 2');
  });

  it('reports converter failures with the Markdown line and leaves the document unchanged', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const before = parser.rezip();
    let msg = '';
    try {
      await composeDocument(parser, 'ok line\n\nbad \\(\\frac{a\\) here\n');
    } catch (err: unknown) {
      msg = err instanceof Error ? err.message : String(err);
    }
    expect(msg).toContain('E_VALIDATION');
    expect(msg).toContain('line 3');
    expect(msg).toContain('Remedy:');
    expect(Buffer.compare(parser.rezip(), before)).toBe(0);
  });

  it('refuses script and entity payloads in math', async () => {
    for (const md of ['x \\(<script>alert(1)</script>\\) y\n', 'x \\(<!DOCTYPE a [<!ENTITY e SYSTEM "file:///etc/passwd">]>\\) y\n']) {
      const parser = DocxParser.createNewDocument('blank');
      let msg = '';
      try {
        await composeDocument(parser, md);
      } catch (err: unknown) {
        msg = err instanceof Error ? err.message : String(err);
      }
      expect(msg).toContain('E_VALIDATION');
      expect(msg).toContain('line 1');
    }
  });
});

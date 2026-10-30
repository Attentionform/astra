import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, validateDocxBuffer } from '@attentionform/astra-engine';
import { compileMarkdown, composeDocument } from '../../packages/engine/src/compose/markdown';

const NO_MATH_SAMPLE = "# Title\n\nA paragraph with **bold**, *it*, `code` and a [link](https://example.com).\n\n- one\n- two\n\n| a | b |\n| - | - |\n| 1 | 2 |\n\n```ts\nconst x = \"$$ not math $$\";\n```\n\nIt costs $5 and $10 today, US$5-$10, path $HOME/$USER and 2^3 * 4_5 here.\n\nLast \\[1\\] ref\n";

const NO_MATH_OPS = [
  {
    "action": "insert",
    "type": "paragraph",
    "anchor": "end",
    "position": "after",
    "text": "Title",
    "id": "md_0001"
  },
  {
    "action": "namedStyle",
    "operation": "apply",
    "styleName": "Heading1",
    "target": "md_0001"
  },
  {
    "action": "insert",
    "type": "paragraph",
    "anchor": "md_0001",
    "position": "after",
    "text": "A paragraph with bold, it, code and a link.",
    "id": "md_0002",
    "params": {
      "mdSpans": [
        {
          "text": "A paragraph with ",
          "bold": false,
          "italic": false,
          "code": false
        },
        {
          "text": "bold",
          "bold": true,
          "italic": false,
          "code": false
        },
        {
          "text": ", ",
          "bold": false,
          "italic": false,
          "code": false
        },
        {
          "text": "it",
          "bold": false,
          "italic": true,
          "code": false
        },
        {
          "text": ", ",
          "bold": false,
          "italic": false,
          "code": false
        },
        {
          "text": "code",
          "bold": false,
          "italic": false,
          "code": true
        },
        {
          "text": " and a ",
          "bold": false,
          "italic": false,
          "code": false
        },
        {
          "text": "link",
          "bold": false,
          "italic": false,
          "code": false,
          "linkUrl": "https://example.com"
        },
        {
          "text": ".",
          "bold": false,
          "italic": false,
          "code": false
        }
      ]
    }
  },
  {
    "action": "insert",
    "type": "paragraph",
    "anchor": "md_0002",
    "position": "after",
    "text": "one",
    "id": "md_0003"
  },
  {
    "action": "insert",
    "type": "paragraph",
    "anchor": "md_0003",
    "position": "after",
    "text": "two",
    "id": "md_0004"
  },
  {
    "action": "list",
    "operation": "create",
    "target": [
      "md_0003",
      "md_0004"
    ],
    "listType": "bullet"
  },
  {
    "action": "insert",
    "type": "table",
    "anchor": "md_0004",
    "position": "after",
    "headers": [
      "a",
      "b"
    ],
    "rows": [
      [
        "1",
        "2"
      ]
    ],
    "id": "md_0005",
    "cellPadding": {
      "top": 120,
      "bottom": 120,
      "left": 160,
      "right": 160
    }
  },
  {
    "action": "table",
    "target": "md_0005",
    "operation": "setCellShading",
    "rowIndex": 1,
    "colIndex": 1,
    "fill": "D9D9D9"
  },
  {
    "action": "table",
    "target": "md_0005",
    "operation": "setCellShading",
    "rowIndex": 1,
    "colIndex": 2,
    "fill": "D9D9D9"
  },
  {
    "action": "insert",
    "type": "paragraph",
    "anchor": "md_0005",
    "position": "after",
    "text": "const x = \"$$ not math $$\";",
    "id": "md_0006",
    "params": {
      "codeBlock": true,
      "codeLanguage": "ts"
    }
  },
  {
    "action": "style",
    "target": "md_0006",
    "spacingBefore": 240
  },
  {
    "action": "insert",
    "type": "paragraph",
    "anchor": "md_0006",
    "position": "after",
    "text": "It costs $5 and $10 today, US$5-$10, path $HOME/$USER and 2^3 * 4_5 here.",
    "id": "md_0007"
  },
  {
    "action": "insert",
    "type": "paragraph",
    "anchor": "md_0007",
    "position": "after",
    "text": "Last [1] ref",
    "id": "md_0008"
  }
] as unknown as ReturnType<typeof compileMarkdown>;

type Op = Record<string, unknown>;
const ops = (md: string): Op[] => compileMarkdown(md) as unknown as Op[];
const equations = (list: Op[]): Op[] => list.filter((o) => o['action'] === 'equation');
const paragraphs = (list: Op[]): Op[] => list.filter((o) => o['action'] === 'insert' && o['type'] === 'paragraph');

function compileError(md: string): string {
  try {
    compileMarkdown(md);
  } catch (err: unknown) {
    return err instanceof Error ? err.message : String(err);
  }
  return '';
}

describe('display math blocks', () => {
  it('compiles a $$ fence to one equation insert between the paragraphs', () => {
    const list = ops('Intro\n\n$$\nE = mc^2\n$$\n\nOutro\n');
    const kinds = list.filter((o) => o['action'] === 'equation' || (o['action'] === 'insert' && o['type'] === 'paragraph'));
    expect(kinds.map((o) => o['action'])).toEqual(['insert', 'equation', 'insert']);
    expect(kinds[1]).toMatchObject({ action: 'equation', operation: 'insert', mode: 'latex', latex: 'E = mc^2' });
  });

  it('reads the one-line $$ form and the bracket form', () => {
    expect(equations(ops('$$ a_1 + b_2 $$\n'))[0]).toMatchObject({ latex: 'a_1 + b_2', mode: 'latex' });
    expect(equations(ops('\\[\nx^2\n\\]\n'))[0]).toMatchObject({ latex: 'x^2', mode: 'latex' });
  });

  it('keeps two blocks in document order', () => {
    const eq = equations(ops('$$\na\n$$\n\ntext\n\n$$\nb\n$$\n'));
    expect(eq.map((o) => o['latex'])).toEqual(['a', 'b']);
  });

  it('leaves a $$ block inside a code fence as code', () => {
    const list = ops('```\n$$\nx\n$$\n```\n');
    expect(equations(list)).toHaveLength(0);
    expect(paragraphs(list).some((o) => String(o['text']).includes('$$'))).toBe(true);
  });

  it('keeps a mid-line $$ as a paragraph', () => {
    const list = ops('not $$ at line start $$ x\n');
    expect(equations(list)).toHaveLength(0);
    expect(paragraphs(list)).toHaveLength(1);
  });

  it('reports an unclosed $$ fence with a line number', () => {
    const msg = compileError('Intro\n\n$$\nE = mc^2\n');
    expect(msg).toContain('E_VALIDATION');
    expect(msg).toContain('never closed');
    expect(msg).toMatch(/line 3/);
  });

  it('does not swallow later paragraphs after an unclosed or blank-broken fence', () => {
    for (const md of ['$$\nE = mc^2\n\nlater prose\n', '\\[\nx^2\n\nlater prose\n']) {
      const msg = compileError(md);
      expect(msg).toContain('never closed');
    }
  });

  it('keeps prose with escaped brackets as prose', () => {
    const a = ops('\\[TODO\\] fix\n\nmid\n\nend \\[note\\]\n');
    expect(equations(a)).toHaveLength(0);
    expect(paragraphs(a)).toHaveLength(3);
    const b = ops('See refs:\n\\[1\\] Smith\n\\[2\\] Jones\n');
    expect(equations(b)).toHaveLength(0);
    expect(paragraphs(b)).toHaveLength(1);
  });

  it('refuses a math block inside a list item and inside an alert body', () => {
    expect(compileError('- item\n\n  $$\n  x\n  $$\n')).toContain('E_VALIDATION');
    expect(compileError('> [!NOTE]\n> $$\n> x\n> $$\n')).toContain('E_VALIDATION');
  });

  it('accepts a long block and refuses one past the length limit', () => {
    const ok = `$$\n${'x + '.repeat(900)}y\n$$\n`; // about 3600 characters
    expect(equations(ops(ok))).toHaveLength(1);
    const msg = compileError(`$$\n${'x + '.repeat(1100)}y\n$$\n`); // past 4000 characters
    expect(msg).toContain('E_VALIDATION');
    expect(msg).toContain('never closed, or is longer than 4000 characters');
  });

  it('stays fast on a long run of lines that start with $$ but never close', () => {
    // The first lexer draft took about 7 seconds for 32000 such lines; the
    // fixed one takes milliseconds. The limit is generous so the test cannot flake.
    const t0 = Date.now();
    const msg = compileError('$$ a\n'.repeat(32000));
    expect(msg === '' || msg.includes('E_VALIDATION')).toBe(true);
    expect(Date.now() - t0).toBeLessThan(3000);
  });

  it('puts space after a table before the equation', () => {
    const list = ops('| a | b |\n|---|---|\n| 1 | 2 |\n\n$$\nx\n$$\n');
    const i = list.findIndex((o) => o['action'] === 'equation');
    expect(list[i + 1]).toMatchObject({ action: 'style', spacingBefore: 240 });
  });

  it('compiles markdown without math to the same operations as before', () => {
    // Expected value captured from compileMarkdown before math support existed.
    expect(compileMarkdown(NO_MATH_SAMPLE, { styleGuide: 'legacy' })).toEqual(NO_MATH_OPS);
  });
});

describe('display math in a document', () => {
  it('builds real OMML (m:sSub without a backslash) that passes the validator', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const res = await composeDocument(parser, '# T\n\n$$ a_1 + b_2 $$\n\nEnd.\n');
    expect(res.success).toBe(true);
    const out = parser.rezip();
    const xml = new AdmZip(out).readAsText('word/document.xml');
    expect(xml.match(/<m:sSub>/g)?.length).toBe(2);
    expect(xml).toContain('<m:oMathPara>');
    expect((await validateDocxBuffer(out)).ok).toBe(true);
  });

  it('converts an equation with a fraction', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const res = await composeDocument(parser, 'Text.\n\n$$\n\\frac{a}{b} + \\dots + x_1\n$$\n\nEnd.\n');
    expect(res.success).toBe(true);
    const out = parser.rezip();
    expect(new AdmZip(out).readAsText('word/document.xml')).toContain('<m:f>');
    expect((await validateDocxBuffer(out)).ok).toBe(true);
  });
});

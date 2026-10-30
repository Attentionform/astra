import { describe, it, expect } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import AdmZip from 'adm-zip';
import { DOMParser } from '@xmldom/xmldom';
import {
  DocxParser,
  OperationEngine,
  convertAstToMarkdown,
  convertDocxToPdf,
  hasExportPdfBackend,
  validateDocxBuffer,
  validateOperationsStrict,
} from '@attentionform/astra-engine';
import {
  COMPOSE_MARKDOWN_MAX_BYTES,
  compileMarkdown,
  composeDocument,
} from '../../packages/engine/src/compose/markdown';
import { withLoSlot } from '../helpers/lo-slot';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

// Slot + budget note: the two live LO renders below hold the suite-wide LO
// slot (see tests/helpers/lo-slot.ts) and carry 300s for queue + own run.

function documentXml(parser: DocxParser): string {
  const zip = new AdmZip(parser.rezip());
  return zip.getEntry('word/document.xml')?.getData().toString('utf8') ?? '';
}

function expectLineError(fn: () => unknown, line: number, fragment: string): void {
  let message = '';
  try {
    fn();
  } catch (err: unknown) {
    message = err instanceof Error ? err.message : String(err);
  }
  expect(message).toContain('E_VALIDATION');
  expect(message).toContain(`line ${line}`);
  expect(message).toContain(fragment);
}

describe('compose headings', () => {
  it('compiles levels one to three into inserts with real style applies', async () => {
    const ops = compileMarkdown('# Alpha\n\n## Beta\n\n### Gamma\n');
    expect(ops.filter((op) => op.action === 'namedStyle').map((op) => (op as { styleName: string }).styleName)).toEqual([
      'Heading1',
      'Heading2',
      'Heading3',
    ]);

    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, '# Alpha\n\n## Beta\n\n### Gamma\n');
    expect(result.success).toBe(true);
    const xml = documentXml(parser);
    const headingDoc = new DOMParser().parseFromString(xml, 'text/xml');
    expect(
      Array.from(headingDoc.getElementsByTagName('w:pStyle')).map((e) => (e as unknown as Element).getAttribute('w:val'))
    ).toEqual(['Heading1', 'Heading2', 'Heading3']);
    expect(
      Array.from(headingDoc.getElementsByTagName('w:t')).map((e) => (e as unknown as Element).textContent)
    ).toEqual(['Alpha', 'Beta', 'Gamma']);
  });

  it('rejects level four with a line number', () => {
    expectLineError(() => compileMarkdown('# Ok\n\n#### Too deep\n'), 3, 'heading level 4');
  });
});

describe('compose paragraphs', () => {
  it('splits on blank lines and joins single newlines with a space', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, 'First line\nsecond line\n\nSecond block\n');
    expect(result.success).toBe(true);
    const xml = documentXml(parser);
    expect(xml).toContain('<w:t>First line second line</w:t>');
    expect(xml).toContain('<w:t>Second block</w:t>');
  });
});

describe('compose lists', () => {
  it('gives bullets and numbered items real numbering', async () => {
    const ops = compileMarkdown('- a\n- b\n\n1. x\n2. y\n');
    const lists = ops.filter((op) => op.action === 'list');
    expect(lists.map((op) => (op as { listType: string }).listType)).toEqual(['bullet', 'numbered']);

    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, '- a\n- b\n\n1. x\n2. y\n');
    expect(result.success).toBe(true);
    const xml = documentXml(parser);
    const listDoc = new DOMParser().parseFromString(xml, 'text/xml');
    expect(listDoc.getElementsByTagName('w:numPr').length).toBe(4);
    expect(Array.from(listDoc.getElementsByTagName('w:t')).map((e) => (e as unknown as Element).textContent)).toEqual([
      'a',
      'b',
      'x',
      'y',
    ]);
  });

  it('rejects nested and task items with line numbers', () => {
    expectLineError(() => compileMarkdown('- top\n  - nested\n'), 2, 'nested lists');
    expectLineError(() => compileMarkdown('- [ ] task\n'), 1, 'task-list');
  });
});

describe('compose tables', () => {
  it('shades the header row and repeats it across pages', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, '| Name | Score |\n|---|---|\n| Ada | 10 |\n');
    expect(result.success).toBe(true);
    const xml = documentXml(parser);
    const tableDoc = new DOMParser().parseFromString(xml, 'text/xml');
    expect(tableDoc.getElementsByTagName('w:tblHeader').length).toBe(1);
    expect(Array.from(tableDoc.getElementsByTagName('w:t')).map((e) => (e as unknown as Element).textContent)).toEqual([
      'Name',
      'Score',
      'Ada',
      '10',
    ]);
    expect(xml).toContain('w:shd');
    // Default Corporate guide: navy header, bold white header text.
    expect(xml).toContain('w:fill="1E3A8A"');
    expect(xml).toContain('<w:b/><w:color w:val="FFFFFF"/>');
    expect(xml).not.toContain('D9D9D9');
  });

  it("keeps the old grey header with styleGuide: 'legacy'", async () => {
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, '| Name | Score |\n|---|---|\n| Ada | 10 |\n', { styleGuide: 'legacy' });
    expect(result.success).toBe(true);
    const xml = documentXml(parser);
    expect(xml).toContain('w:fill="D9D9D9"');
    expect(xml).not.toContain('1E3A8A');
  });

  it('stripes the data rows from the guide alternating fill and leaves the header alone', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, '| N |\n|---|\n| a |\n| b |\n| c |\n');
    expect(result.success).toBe(true);
    const fills = [...documentXml(parser).matchAll(/<w:tr>[\s\S]*?<\/w:tr>/g)].map(
      (m) => /w:fill="([0-9A-F]{6})"/.exec(m[0])?.[1],
    );
    expect(fills).toEqual(['1E3A8A', 'F8FAFC', 'FFFFFF', 'F8FAFC']);
  });

  it('rejects ragged rows naming the line and the cell counts', () => {
    expectLineError(
      () => compileMarkdown('# T\n\n| a | b |\n|---|---|\n| x | y | z |\n'),
      3,
      'table row 1 has 3 cells, header has 2'
    );
  });
});

describe('compose dividers', () => {
  it('emits a real divider and never a break op', async () => {
    const ops = compileMarkdown('Before\n\n---\n\nAfter\n');
    expect(ops.some((op) => (op as { type?: string }).type === 'divider')).toBe(true);
    for (const op of ops) {
      expect((op as { type?: string }).type).not.toBe('pageBreak');
      expect((op as { type?: string }).type).not.toBe('sectionBreak');
    }

    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, 'Before\n\n---\n\nAfter\n');
    expect(result.success).toBe(true);
    const xml = documentXml(parser);
    expect(xml).toContain('w:pBdr');
    expect(xml).not.toContain('w:type="page"');
  });
});

describe('compose code fences', () => {
  it('keeps line breaks, monospace runs, and the fence language', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, '```python\nprint(1)\nprint(2)\n```\n');
    expect(result.success).toBe(true);
    const xml = documentXml(parser);
    const fenceDoc = new DOMParser().parseFromString(xml, 'text/xml');
    expect(Array.from(fenceDoc.getElementsByTagName('w:t')).map((e) => (e as unknown as Element).textContent)).toEqual([
      'print(1)',
      'print(2)',
    ]);
    expect(fenceDoc.getElementsByTagName('w:br').length).toBe(1);
    // Code blocks use the template Code character style (Consolas/indigo/
    // 10pt) with block shading, a box border, an indent, and a language
    // bookmark so export rebuilds the fence with its tag.
    expect(xml).toContain('w:val="Code"');
    expect(xml).toContain('Consolas');
    expect(xml).toContain('w:val="4338CA"');
    expect(xml).toContain('w:fill="F1F5F9"');
    expect(xml).toContain('_AstraCode_python');

    const ops = compileMarkdown('```python\nprint(1)\n```\n');
    const insert = ops.find((op) => op.action === 'insert');
    expect((insert as { params?: { codeLanguage?: string } }).params?.codeLanguage).toBe('python');
    expect((insert as { params?: { codeBlock?: boolean } }).params?.codeBlock).toBe(true);
  });

  it('rejects chart fences with a line number', () => {
    expectLineError(() => compileMarkdown('```mermaid\ngraph TD\n```\n'), 1, 'chart');
  });

  it('keeps autolinks and footnote markers literal inside fences', async () => {
    const ops = compileMarkdown('```text\n<https://example.com>\n[^a] literal\n```\n');
    const insert = ops.find((op) => op.action === 'insert');
    expect((insert as { text?: string }).text).toContain('<https://example.com>');
    expect((insert as { text?: string }).text).toContain('[^a] literal');

    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, '```text\n<https://example.com>\n[^a] literal\n```\n');
    expect(result.success).toBe(true);
    const xml = documentXml(parser);
    expect(xml).toContain('https://example.com');
  });
});

describe('compose footnotes', () => {
  it('resolves definitions used before they are defined', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, 'Claim[^late] here.\n\n[^late]: The late body\n');
    expect(result.success).toBe(true);
    const xml = documentXml(parser);
    const footnoteDoc = new DOMParser().parseFromString(xml, 'text/xml');
    expect(footnoteDoc.getElementsByTagName('w:footnoteReference').length).toBe(1);
    expect(xml).not.toContain('[^late]');
    const zip = new AdmZip(parser.rezip());
    const notes = zip.getEntry('word/footnotes.xml')?.getData().toString('utf8') ?? '';
    expect(notes).toContain('The late body');
  });

  it('fails undefined references, unused definitions, and empty bodies with line numbers', () => {
    expectLineError(() => compileMarkdown('Claim[^missing] here.\n'), 1, 'undefined footnote reference');
    expectLineError(() => compileMarkdown('Plain text.\n\n[^orphan]: Never used\n'), 3, 'unused footnote definition');
    expectLineError(() => compileMarkdown('Text.\n\n[^empty]:\n'), 3, 'empty body');
  });
});

describe('compose rejects out-of-scope input', () => {
  const cases: Array<[string, string, number, string]> = [
    ['image', '![alt](pic.png)\n', 1, 'needs the folder'],
    ['html block', '<div>hi</div>\n', 1, 'raw HTML'],
    ['inline html', 'Text <b>bold</b> here\n', 1, 'raw HTML'],
    ['autolink url', 'Contact <https://example.com> now\n', 1, 'raw HTML'],
    ['autolink bare', '<https://example.com>\n', 1, 'raw HTML'],
    ['autolink mailto', 'Mail <mailto:foo@bar.com> here\n', 1, 'raw HTML'],
    ['autolink email', 'Email <user@example.com> test\n', 1, 'raw HTML'],
    ['autolink heading', '# Title <https://example.com>\n', 1, 'raw HTML'],
    ['autolink list item', '- item <https://example.com>\n', 1, 'raw HTML'],
    ['autolink table cell', '| a |\n|---|\n| <https://example.com> |\n', 1, 'raw HTML'],
    ['autolink footnote body', 'Note[^a].\n\n[^a]: See <https://example.com>\n', 3, 'image or HTML'],
    ['blockquote', '> quoted\n', 1, 'blockquotes'],
    ['container fence', '::: note\ntext\n:::\n', 1, 'container fences'],
    ['heading attribute', '# Title {#custom}\n', 1, 'custom heading attributes'],
    ['table cell note', '| a |\n|---|\n| x[^1] |\n\n[^1]: body\n', 3, 'inside tables'],
  ];
  for (const [name, source, line, fragment] of cases) {
    it(`fails ${name} with a line number`, () => {
      expectLineError(() => compileMarkdown(source), line, fragment);
    });
  }

  it('writes nothing when compilation fails', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const before = parser.rezip();
    await expect(composeDocument(parser, '# Fine\n\n![alt](pic.png)\n')).rejects.toThrow(/E_VALIDATION/);
    expect(parser.rezip().equals(before)).toBe(true);
  });

  it('compiled output passes the strict operation gate', () => {
    const ops = compileMarkdown(
      '# Title\n\nBody text\n\n- a\n- b\n\n| h1 | h2 |\n|---|---|\n| a | b |\n\n---\n\n```js\ncode()\n```\n\nNote[^1].\n\n[^1]: Body\n'
    );
    expect(validateOperationsStrict(ops).ok).toBe(true);
  });
});

describe('compose golden op sequence', () => {
  it('stays byte-stable for a small document', () => {
    expect(compileMarkdown('# Hi\n\nBody text\n')).toEqual([
      { action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'Hi', id: 'md_0001' },
      { action: 'namedStyle', operation: 'apply', styleName: 'Heading1', target: 'md_0001' },
      { action: 'insert', type: 'paragraph', anchor: 'md_0001', position: 'after', text: 'Body text', id: 'md_0002' },
    ]);
  });
});

describe('compose placement', () => {
  it('appends after existing content: seed text stays before composed text', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const seedText = 'Seed paragraph stays first';
    const before = await parser.getAST();
    const seeder = new OperationEngine(before, parser);
    const seeded = await seeder.executeBatch({
      operations: [
        { action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: seedText, id: 'seed-first' },
      ],
      sequential: true,
    });
    expect(seeded.success).toBe(true);

    const result = await composeDocument(parser, '# Composed title\n\nComposed body\n');
    expect(result.success).toBe(true);

    const summary = parser.getStructuralSummary();
    const texts = summary.nodes.map((node) => node.text ?? '');
    const seedIdx = texts.findIndex((text) => text.includes(seedText));
    const titleIdx = texts.findIndex((text) => text.includes('Composed title'));
    const bodyIdx = texts.findIndex((text) => text.includes('Composed body'));
    expect(seedIdx).toBeGreaterThanOrEqual(0);
    expect(titleIdx).toBeGreaterThan(seedIdx);
    expect(bodyIdx).toBeGreaterThan(seedIdx);

    const xml = documentXml(parser);
    expect(xml.indexOf(seedText)).toBeGreaterThanOrEqual(0);
    expect(xml.indexOf(seedText)).toBeLessThan(xml.indexOf('<w:t>Composed title</w:t>'));
    expect(xml.indexOf(seedText)).toBeLessThan(xml.indexOf('<w:t>Composed body</w:t>'));
  });

  it('replaces the pristine blank template seeds instead of appending', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, '# Composed title\n\nComposed body\n');
    expect(result.success).toBe(true);

    const texts = parser.getStructuralSummary().nodes.map((node) => node.text ?? '');
    expect(texts.some((text) => text.includes('Blank Document'))).toBe(false);
    expect(texts.some((text) => text.includes('Start typing'))).toBe(false);
    expect(texts).toContain('Composed title');
    expect(texts).toContain('Composed body');

    const xml = documentXml(parser);
    expect(xml).not.toContain('Blank Document');
    expect(xml).not.toContain('Start typing');
    expect(xml).toContain('<w:t>Composed title</w:t>');

    const strict = await validateDocxBuffer(parser.rezip());
    expect(strict.ok).toBe(true);
    expect(strict.errors).toEqual([]);
  });

  it('keeps append-only behavior once the blank template is no longer pristine', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const before = await parser.getAST();
    const seeder = new OperationEngine(before, parser);
    const seeded = await seeder.executeBatch({
      operations: [
        { action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'User note stays', id: 'user-note' },
      ],
      sequential: true,
    });
    expect(seeded.success).toBe(true);

    const result = await composeDocument(parser, '# Composed title\n\nComposed body\n');
    expect(result.success).toBe(true);

    const texts = parser.getStructuralSummary().nodes.map((node) => node.text ?? '');
    expect(texts.some((text) => text.includes('Blank Document'))).toBe(true);
    expect(texts).toContain('User note stays');
    expect(texts).toContain('Composed title');
    expect(texts).toContain('Composed body');

    const xml = documentXml(parser);
    expect(xml.indexOf('User note stays')).toBeLessThan(xml.indexOf('<w:t>Composed title</w:t>'));
  });

  it('keeps append-only behavior when a seed carries a footnote', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const seeder = new OperationEngine(ast, parser);
    const seedId = parser.getStructuralSummary().nodes[0].id;
    const noted = await seeder.executeBatch({
      operations: [{ action: 'footnote', anchor: seedId, text: 'seed note' }],
      sequential: true,
    });
    expect(noted.success).toBe(true);

    const result = await composeDocument(parser, '# Composed title\n\nComposed body\n');
    expect(result.success).toBe(true);
    const texts = parser.getStructuralSummary().nodes.map((node) => node.text ?? '');
    expect(texts.some((text) => text.includes('Blank Document'))).toBe(true);
    expect(texts).toContain('Composed title');
    const zip = new AdmZip(parser.rezip());
    expect(zip.getEntry('word/footnotes.xml')).not.toBeNull();
  });

  it('keeps append-only behavior when a seed carries an endnote', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const seeder = new OperationEngine(ast, parser);
    const seedId = parser.getStructuralSummary().nodes[0].id;
    const noted = await seeder.executeBatch({
      operations: [{ action: 'endnote', anchor: seedId, text: 'seed end' }],
      sequential: true,
    });
    expect(noted.success).toBe(true);

    const result = await composeDocument(parser, '# Composed title\n\nComposed body\n');
    expect(result.success).toBe(true);
    const texts = parser.getStructuralSummary().nodes.map((node) => node.text ?? '');
    expect(texts.some((text) => text.includes('Blank Document'))).toBe(true);
    const zip = new AdmZip(parser.rezip());
    expect(zip.getEntry('word/endnotes.xml')).not.toBeNull();
  });

  it('keeps append-only behavior when a seed carries a comment', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const seeder = new OperationEngine(ast, parser);
    const seedId = parser.getStructuralSummary().nodes[0].id;
    const noted = await seeder.executeBatch({
      operations: [{ action: 'comment', operation: 'insert', target: seedId, text: 'seed comment' }],
      sequential: true,
    });
    expect(noted.success).toBe(true);

    const result = await composeDocument(parser, '# Composed title\n\nComposed body\n');
    expect(result.success).toBe(true);
    const texts = parser.getStructuralSummary().nodes.map((node) => node.text ?? '');
    expect(texts.some((text) => text.includes('Blank Document'))).toBe(true);
    const zip = new AdmZip(parser.rezip());
    expect(zip.getEntry('word/comments.xml')).not.toBeNull();
  });

  it('keeps append-only behavior when a seed paragraph style changes', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const seeder = new OperationEngine(ast, parser);
    const seedId = parser.getStructuralSummary().nodes[0].id;
    const restyled = await seeder.executeBatch({
      operations: [{ action: 'namedStyle', operation: 'apply', styleName: 'Heading1', target: seedId }],
      sequential: true,
    });
    expect(restyled.success).toBe(true);

    const result = await composeDocument(parser, '# Composed title\n\nComposed body\n');
    expect(result.success).toBe(true);
    const texts = parser.getStructuralSummary().nodes.map((node) => node.text ?? '');
    expect(texts.some((text) => text.includes('Blank Document'))).toBe(true);
    expect(texts).toContain('Composed title');
  });

  it('restores the pre-batch package and zeroes the count when seed removal fails', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const before = parser.rezip();
    const original = OperationEngine.prototype.executeBatch;
    let restoreCalls = 0;
    const innerRestore = parser.restorePackageSnapshot.bind(parser);
    parser.restorePackageSnapshot = (snap: Map<string, Buffer>) => {
      restoreCalls += 1;
      return innerRestore(snap);
    };
    OperationEngine.prototype.executeBatch = async function (
      this: OperationEngine,
      req: { operations: Array<{ action?: string }>; sequential?: boolean }
    ): Promise<Awaited<ReturnType<OperationEngine['executeBatch']>>> {
      const ops = req.operations as Array<{ action?: string }>;
      if (ops.some((op) => op.action === 'delete')) {
        return {
          success: false,
          appliedCount: 0,
          errors: [{ operationIndex: 0, error: 'injected delete failure', hint: 'retry the compose' }],
          rollbackTriggered: true,
          skippedIds: ['skipped-seed'],
          warnings: ['delete-warn'],
          revisionsPending: 3,
          revisionsRemedy: 'accept revisions',
          possiblyStaleBodyRefs: [{ nodeId: 'seed-node', preview: 'seed preview' }],
        } as unknown as Awaited<ReturnType<OperationEngine['executeBatch']>>;
      }
      return original.call(this, req as never) as Promise<Awaited<ReturnType<OperationEngine['executeBatch']>>>;
    };
    try {
      const result = await composeDocument(parser, '# Composed title\n\nComposed body\n');
      expect(result.success).toBe(false);
      expect(result.rollbackTriggered).toBe(true);
      expect(result.appliedCount).toBe(0);
      expect(result.revisionsPending).toBe(3);
      expect(result.revisionsRemedy).toBe('accept revisions');
      expect(result.skippedIds).toContain('skipped-seed');
      expect(result.warnings).toContain('delete-warn');
      expect(result.possiblyStaleBodyRefs.some((ref) => ref.nodeId === 'seed-node')).toBe(true);
      expect(restoreCalls).toBeGreaterThan(0);
      expect(parser.rezip().equals(before)).toBe(true);
    } finally {
      OperationEngine.prototype.executeBatch = original;
    }
  });

  it('merges delete-side revision metadata on a successful seed removal', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const original = OperationEngine.prototype.executeBatch;
    OperationEngine.prototype.executeBatch = async function (
      this: OperationEngine,
      req: { operations: Array<{ action?: string }>; sequential?: boolean }
    ): Promise<Awaited<ReturnType<OperationEngine['executeBatch']>>> {
      const out = (await original.call(this, req as never)) as Awaited<
        ReturnType<OperationEngine['executeBatch']>
      >;
      const ops = req.operations as Array<{ action?: string }>;
      if (ops.some((op) => op.action === 'delete')) {
        return {
          ...out,
          warnings: [...out.warnings, 'seed-delete-warn'],
          skippedIds: [...out.skippedIds, 'seed-skipped'],
          revisionsPending: 2,
          revisionsRemedy: 'accept revisions',
        };
      }
      return out;
    };
    try {
      const result = await composeDocument(parser, '# Composed title\n\nComposed body\n');
      expect(result.success).toBe(true);
      expect(result.warnings).toContain('seed-delete-warn');
      expect(result.skippedIds).toContain('seed-skipped');
      expect(result.revisionsPending).toBe(2);
      expect(result.revisionsRemedy).toBe('accept revisions');
      const texts = parser.getStructuralSummary().nodes.map((node) => node.text ?? '');
      expect(texts.some((text) => text.includes('Blank Document'))).toBe(false);
    } finally {
      OperationEngine.prototype.executeBatch = original;
    }
  });
});

describe('compose long manuscript acceptance', () => {
  it('builds a valid multi-page document that round-trips its text', async () => {
    const sections: string[] = [];
    for (let s = 1; s <= 30; s++) {
      const sentence = `Section ${s} background material fills the page with domain prose about testing aged garlic and endurance protocols`;
      sections.push(`# Report section ${s}`);
      sections.push('');
      for (let p = 1; p <= 6; p++) {
        sections.push(
          `Section ${s} paragraph ${p} states that ${sentence} in trial ${p} of block ${s}, and the follow-up review confirms the same endurance reading held across the full observation window without protocol drift.`
        );
        sections.push('');
      }
      sections.push(`Key outcome ${s} of the endurance block[^note-${s}] closes the narrative.`);
      sections.push('');
      sections.push(`- finding alpha ${s}`);
      sections.push(`- finding beta ${s}`);
      sections.push(`- finding gamma ${s}`);
      sections.push('');
      sections.push('| Metric | Value | Unit |');
      sections.push('|---|---|---|');
      sections.push(`| Endurance ${s} | ${(s * 7) % 100} | minutes |`);
      sections.push(`| Recovery ${s} | ${(s * 13) % 60} | hours |`);
      sections.push('');
      sections.push('```text');
      sections.push(`log line one for section ${s}`);
      sections.push(`log line two for section ${s}`);
      sections.push('```');
      sections.push('');
      sections.push('---');
      sections.push('');
      sections.push(`[^note-${s}]: Supporting detail for section ${s} recorded during the trial review.`);
      sections.push('');
    }
    const manuscript = sections.join('\n');

    const words = manuscript.split(/\s+/).length;
    expect(words).toBeGreaterThan(9000);

    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, manuscript);
    expect(result.success).toBe(true);
    expect(result.errors).toEqual([]);

    const buffer = parser.rezip();
    const strict = await validateDocxBuffer(buffer);
    expect(strict.ok).toBe(true);
    expect(strict.errors).toEqual([]);
    const ooxml = await parser.validateOOXML();
    expect(ooxml.ok).toBe(true);
    expect(ooxml.errors).toEqual([]);

    const ast = await parser.getAST();
    const roundTripped = convertAstToMarkdown(ast, parser);
    expect(roundTripped).toContain('Report section 7');
    expect(roundTripped).toContain('Section 12 paragraph 3');
    expect(roundTripped).toContain('finding beta 21');
    expect(roundTripped).toContain('Endurance 30');
    expect(roundTripped).toContain('Supporting detail for section 5');

    const summary = parser.getStructuralSummary();
    expect(summary.nodes.length).toBeGreaterThan(300);

    const pdfHealthy = hasExportPdfBackend();
    if (!pdfHealthy) {
      console.log('PDF backend absent (hasExportPdfBackend false): skipping PDF build, env flag honest');
      return;
    }
    const outPdf = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'astra-compose-')), 'manuscript.pdf');
    const built = await withLoSlot(() => convertDocxToPdf(buffer, outPdf));
    expect(fs.existsSync(built)).toBe(true);
    expect(fs.statSync(built).size).toBeGreaterThan(0);
  }, 300000);
});

describe('compose endnotes', () => {
  it('routes en- labels through the endnote path and keeps footnotes separate', async () => {
    const ops = compileMarkdown('Foot[^a] and end[^en-1].\n\n[^a]: Foot body\n\n[^en-1]: End body\n');
    expect(ops.some((op) => op.action === 'footnote')).toBe(true);
    expect(ops.some((op) => op.action === 'endnote')).toBe(true);

    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, 'Foot[^a] and end[^en-1].\n\n[^a]: Foot body\n\n[^en-1]: End body\n');
    expect(result.success).toBe(true);
    const xml = documentXml(parser);
    expect(xml).toContain('w:footnoteReference');
    expect(xml).toContain('w:endnoteReference');
    expect(xml).not.toContain('[^a]');
    expect(xml).not.toContain('[^en-1]');
    const zip = new AdmZip(parser.rezip());
    expect(zip.getEntry('word/footnotes.xml')?.getData().toString('utf8') ?? '').toContain('Foot body');
    expect(zip.getEntry('word/endnotes.xml')?.getData().toString('utf8') ?? '').toContain('End body');
  });

  it('fails undefined, unused, and empty endnote definitions with line numbers', () => {
    expectLineError(() => compileMarkdown('Claim[^en-x] here.\n'), 1, 'undefined footnote reference');
    expectLineError(() => compileMarkdown('Plain.\n\n[^en-orphan]: Never used\n'), 3, 'unused endnote definition');
    expectLineError(() => compileMarkdown('Text.\n\n[^en-empty]:\n'), 3, 'empty body');
  });

  it('writes nothing when endnote compilation fails', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const before = parser.rezip();
    await expect(composeDocument(parser, 'Claim[^en-missing] here.\n')).rejects.toThrow(/E_VALIDATION/);
    expect(parser.rezip().equals(before)).toBe(true);
  });
});

describe('compose citations', () => {
  it('numbers keys in first-appearance order and builds a references section', async () => {
    const ops = compileMarkdown('First [@doe24] then [@smith23] then [@doe24] again.\n\n[@doe24]: Doe entry\n\n[@smith23]: Smith entry\n');
    const texts = ops.filter((op) => op.action === 'insert').map((op) => (op as { text?: string }).text ?? '');
    expect(texts[0]).toBe('First [1] then [2] then [1] again.');
    expect(texts).toContain('References');
    expect(texts).toContain('[1] Doe entry');
    expect(texts).toContain('[2] Smith entry');

    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, 'First [@doe24] then [@smith23] then [@doe24] again.\n\n[@doe24]: Doe entry\n\n[@smith23]: Smith entry\n');
    expect(result.success).toBe(true);
    const xml = documentXml(parser);
    expect(xml).toContain('References');
    expect(xml).toContain('[1] Doe entry');
    expect(xml).toContain('[2] Smith entry');
  });

  it('fails undefined, unused, empty, duplicate, and multi-key citations with line numbers', () => {
    expectLineError(() => compileMarkdown('See [@missing] now.\n'), 1, 'undefined citation');
    expectLineError(() => compileMarkdown('Plain.\n\n[@orphan]: Never used\n'), 3, 'unused citation definition');
    expectLineError(() => compileMarkdown('Text.\n\n[@empty]:\n'), 3, 'empty body');
    expectLineError(() => compileMarkdown('A [@k].\n\n[@k]: One\n\n[@k]: Two\n'), 5, 'duplicate citation definition');
    expectLineError(() => compileMarkdown('See [@a; @b].\n\n[@a]: A\n\n[@b]: B\n'), 1, 'malformed');
  });

  it('writes nothing when citation compilation fails', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const before = parser.rezip();
    await expect(composeDocument(parser, 'See [@missing] now.\n')).rejects.toThrow(/E_VALIDATION/);
    expect(parser.rezip().equals(before)).toBe(true);
  });
});

describe('compose image and chart references', () => {
  const png1x1 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
  function assetDir(): string {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-assets-'));
    fs.writeFileSync(path.join(dir, 'tiny.png'), Buffer.from(png1x1, 'base64'));
    fs.writeFileSync(
      path.join(dir, 'chart.json'),
      JSON.stringify({ chartType: 'column', title: 'Sales', categories: ['Q1', 'Q2'], series: [{ name: 'Rev', values: [10, 20] }] })
    );
    fs.writeFileSync(path.join(dir, 'bad.json'), '{ not json');
    return dir;
  }

  function wideBmp(width: number, height: number): Buffer {
    const rowBytes = Math.ceil((width * 3) / 4) * 4;
    const pixels = Buffer.alloc(rowBytes * height, 0x80);
    const header = Buffer.alloc(54, 0);
    header.write('BM', 0);
    header.writeUInt32LE(54 + pixels.length, 2);
    header.writeUInt32LE(54, 10);
    header.writeUInt32LE(40, 14);
    header.writeInt32LE(width, 18);
    header.writeInt32LE(height, 22);
    header.writeUInt16LE(1, 26);
    header.writeUInt16LE(24, 28);
    header.writeUInt32LE(pixels.length, 34);
    return Buffer.concat([header, pixels]);
  }

  it('embeds images from the allowed directory in source order', async () => {
    const dir = assetDir();
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, 'Before\n\n![Tiny](tiny.png)\n\nAfter\n', { allowedDir: dir });
    expect(result.success).toBe(true);
    const xml = documentXml(parser);
    expect(xml).toContain('w:drawing');
    expect(xml.indexOf('<w:t>Before</w:t>')).toBeLessThan(xml.indexOf('<w:t>After</w:t>'));
  });

  it('fits images to the page width keeping the aspect ratio', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-wide-'));
    fs.writeFileSync(path.join(dir, 'wide.bmp'), wideBmp(1200, 600));
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, 'Before\n\n![Wide](wide.bmp)\n\nAfter\n', { allowedDir: dir });
    expect(result.success).toBe(true);

    const ops = compileMarkdown('![Wide](wide.bmp)\n', { allowedDir: dir });
    const imageOp = ops.find((op) => (op as { type?: string }).type === 'image') as unknown as {
      width?: number;
      height?: number;
    };
    expect(imageOp.width).toBe(5486400);
    expect(imageOp.height).toBe(2743200);

    const xml = documentXml(parser);
    expect(xml).toContain('w:drawing');
    const extent = xml.match(/<wp:extent cx="(\d+)" cy="(\d+)"\/>/);
    expect(extent?.slice(1, 3)).toEqual(['5486400', '2743200']);
    const graphicExtent = xml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
    expect(graphicExtent?.slice(1, 3)).toEqual(['5486400', '2743200']);
    const ratio = Number(extent?.[1]) / Number(extent?.[2]);
    expect(ratio).toBeCloseTo(2, 2);
    expect(xml.indexOf('<w:t>Before</w:t>')).toBeLessThan(xml.indexOf('<w:t>After</w:t>'));

    const strict = await validateDocxBuffer(parser.rezip());
    expect(strict.ok).toBe(true);
    expect(strict.errors).toEqual([]);

    if (!hasExportPdfBackend()) {
      console.log('PDF backend absent (hasExportPdfBackend false): skipping fitted-image PDF build, env flag honest');
      return;
    }
    const outPdf = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'astra-compose-img-')), 'wide.pdf');
    const built = await withLoSlot(() => convertDocxToPdf(parser.rezip(), outPdf));
    expect(fs.existsSync(built)).toBe(true);
    expect(fs.statSync(built).size).toBeGreaterThan(0);
  }, 300000);

  it('keeps small images at intrinsic size instead of stretching', async () => {
    const dir = assetDir();
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, '![Tiny](tiny.png)\n', { allowedDir: dir });
    expect(result.success).toBe(true);
    const xml = documentXml(parser);
    const extent = xml.match(/<wp:extent cx="(\d+)" cy="(\d+)"\/>/);
    expect(extent?.slice(1, 3)).toEqual(['9525', '9525']);
  });

  it('builds charts from JSON files in the allowed directory', async () => {
    const dir = assetDir();
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, 'Before\n\n![Sales](chart.json)\n\nAfter\n', { allowedDir: dir });
    expect(result.success).toBe(true);
    const xml = documentXml(parser);
    expect(xml.includes('w:drawing') || xml.includes('c:chart')).toBe(true);
    expect(xml).toContain('Sales');
  });

  it('fails closed without an allowed directory, outside the jail, and inline', () => {
    expectLineError(() => compileMarkdown('![Tiny](tiny.png)\n'), 1, 'needs the folder');
    const dir = assetDir();
    expectLineError(() => compileMarkdown('![Evil](../../etc/passwd)\n', { allowedDir: dir }), 1, 'outside the allowed directory');
    expectLineError(() => compileMarkdown('![Remote](https://example.com/pic.png)\n', { allowedDir: dir }), 1, 'outside the allowed directory');
    expectLineError(() => compileMarkdown('Text ![Tiny](tiny.png) inline\n', { allowedDir: dir }), 1, 'on their own block line');
    expectLineError(() => compileMarkdown('![Bad](bad.json)\n', { allowedDir: dir }), 1, 'not valid JSON');
  });

  it('writes nothing when asset compilation fails', async () => {
    const dir = assetDir();
    const parser = DocxParser.createNewDocument('blank');
    const before = parser.rezip();
    await expect(composeDocument(parser, '![Evil](../../etc/passwd)\n', { allowedDir: dir })).rejects.toThrow(/E_VALIDATION/);
    expect(parser.rezip().equals(before)).toBe(true);
  });
});

describe('compose template seeds', () => {
  it('appends onto matching report and memo seeds', async () => {
    for (const template of ['report', 'memo'] as const) {
      const parser = DocxParser.createNewDocument(template);
      const result = await composeDocument(parser, 'Composed body\n', { template });
      expect(result.success).toBe(true);
      const texts = parser.getStructuralSummary().nodes.map((node) => node.text ?? '');
      const marker = template === 'report' ? 'Q3 Executive' : 'MEMORANDUM';
      const seedIdx = texts.findIndex((text) => text.includes(marker));
      const bodyIdx = texts.findIndex((text) => text.includes('Composed body'));
      expect(seedIdx).toBeGreaterThanOrEqual(0);
      expect(bodyIdx).toBeGreaterThan(seedIdx);
    }
  });

  it('replaces matching blank seeds instead of keeping them', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, 'Composed body\n', { template: 'blank' });
    expect(result.success).toBe(true);
    const texts = parser.getStructuralSummary().nodes.map((node) => node.text ?? '');
    expect(texts.some((text) => text.includes('Blank Document'))).toBe(false);
    expect(texts.some((text) => text.includes('Start typing'))).toBe(false);
    expect(texts.some((text) => text.includes('Composed body'))).toBe(true);
  });

  it('rejects mismatched seeds and bad names before any write', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const before = parser.rezip();
    await expect(composeDocument(parser, 'Body\n', { template: 'report' })).rejects.toThrow(/does not hold the report seed/);
    expect(parser.rezip().equals(before)).toBe(true);
    await expect(composeDocument(DocxParser.createNewDocument('blank'), 'Body\n', { template: 'other' as never })).rejects.toThrow(/template must be one of/);
  });
});

describe('compose inline spans', () => {
  it('splits bold, italic, code, and links into runs with real hyperlink targets', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, 'Hello **bold** and *italic* and `code` and [docs](https://example.com) end\n');
    expect(result.success).toBe(true);
    const xml = documentXml(parser);
    const spanDoc = new DOMParser().parseFromString(xml, 'text/xml');
    expect(spanDoc.getElementsByTagName('w:b').length).toBe(1);
    expect(spanDoc.getElementsByTagName('w:i').length).toBe(1);
    expect(spanDoc.getElementsByTagName('w:hyperlink').length).toBe(1);
    expect(
      Array.from(spanDoc.getElementsByTagName('w:rFonts')).filter(
        (e) => (e as unknown as Element).getAttribute('w:ascii') === 'Courier New'
      ).length
    ).toBe(1);
    expect(Array.from(spanDoc.getElementsByTagName('w:t')).map((e) => (e as unknown as Element).textContent)).toEqual([
      'Hello ',
      'bold',
      ' and ',
      'italic',
      ' and ',
      'code',
      ' and ',
      'docs',
      ' end',
    ]);
    expect(xml).not.toContain('**bold**');
    expect(xml).not.toContain('[docs](https://example.com)');
    const zip = new AdmZip(parser.rezip());
    const rels = zip.getEntry('word/_rels/document.xml.rels')?.getData().toString('utf8') ?? '';
    expect(rels).toContain('https://example.com');
  });

  it('keeps span markers literal inside fenced code', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, '```text\n**not bold** [not link](https://example.com)\n```\n');
    expect(result.success).toBe(true);
    const xml = documentXml(parser);
    expect(xml).toContain('**not bold**');
  });

  it('fails bad link schemes and rich text in tables with line numbers', () => {
    expectLineError(() => compileMarkdown('See [x](javascript:alert(1)) here\n'), 1, 'not allowed');
    expectLineError(() => compileMarkdown('| H |\n|---|\n| **bold** cell |\n'), 3, 'inside tables');
    expectLineError(() => compileMarkdown('| H |\n|---|\n| [t](https://example.com) |\n'), 3, 'inside tables');
  });

  it('writes nothing when span compilation fails', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const before = parser.rezip();
    await expect(composeDocument(parser, 'See [x](javascript:alert(1)) here\n')).rejects.toThrow(/E_VALIDATION/);
    expect(parser.rezip().equals(before)).toBe(true);
  });
});

describe('compose inline data-url images', () => {
  const png1x1 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

  function bigBmp(width: number, height: number): Buffer {
    const rowBytes = Math.ceil((width * 3) / 4) * 4;
    const pixels = Buffer.alloc(rowBytes * height, 0x80);
    const header = Buffer.alloc(54, 0);
    header.write('BM', 0);
    header.writeUInt32LE(54 + pixels.length, 2);
    header.writeUInt32LE(54, 10);
    header.writeUInt32LE(40, 14);
    header.writeInt32LE(width, 18);
    header.writeInt32LE(height, 22);
    header.writeUInt16LE(1, 26);
    header.writeUInt16LE(24, 28);
    header.writeUInt32LE(pixels.length, 34);
    return Buffer.concat([header, pixels]);
  }

  it('embeds data-url bytes with no allowed directory and keeps intrinsic size', async () => {
    const ops = compileMarkdown(`![Tiny](data:image/png;base64,${png1x1})\n`);
    const imageOp = ops.find((op) => (op as { type?: string }).type === 'image') as unknown as {
      fileData?: string;
      width?: number;
      height?: number;
      altText?: string;
    };
    expect(imageOp.fileData).toContain('data:image/png;base64,');
    expect([imageOp.width, imageOp.height]).toEqual([9525, 9525]);
    expect(imageOp.altText).toBe('Tiny');

    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, `Before\n\n![Tiny](data:image/png;base64,${png1x1})\n\nAfter\n`);
    expect(result.success).toBe(true);
    const xml = documentXml(parser);
    expect(xml).toContain('w:drawing');
    expect(xml.indexOf('<w:t>Before</w:t>')).toBeLessThan(xml.indexOf('<w:t>After</w:t>'));
    const strict = await validateDocxBuffer(parser.rezip());
    expect(strict.ok).toBe(true);
    expect(strict.errors).toEqual([]);
  });

  it('fails closed on garbage data urls with line numbers', () => {
    expectLineError(() => compileMarkdown('![X](data:image/png;base64,!!!)\n'), 1, 'invalid base64');
    expectLineError(() => compileMarkdown('![X](data:image/png;base64,)\n'), 1, 'no bytes');
    expectLineError(() => compileMarkdown('![X](data:text/plain;base64,aGk=)\n'), 1, 'not a supported image');
    expectLineError(() => compileMarkdown('![X](data:image/png;base64,aGk=)\n'), 1, 'not a known image');
    // Oversize payloads die fail-closed at the raw-input ceiling before the
    // lexer runs, so pin the closed limit failure here (not a line error).
    {
      const big = `![Big](data:image/bmp;base64,${bigBmp(3000, 2000).toString('base64')})\n`;
      let message = '';
      let code: unknown;
      try {
        compileMarkdown(big);
      } catch (err: unknown) {
        message = err instanceof Error ? err.message : String(err);
        code = (err as { code?: unknown }).code;
      }
      expect(message).toContain('E_LIMIT');
      expect(message).toContain('3000000');
      expect(code).toBe('E_LIMIT');
    }
  });

  it('round-trips an exported data-url image back to valid and stable markdown', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const first = await composeDocument(parser, `Before\n\n![Tiny](data:image/png;base64,${png1x1})\n\nAfter\n`);
    expect(first.success).toBe(true);
    const exported = convertAstToMarkdown(await parser.getAST(), parser);
    expect(exported).toContain('data:image/png;base64,');

    const again = DocxParser.createNewDocument('blank');
    const second = await composeDocument(again, exported);
    expect(second.success).toBe(true);
    const strict = await validateDocxBuffer(again.rezip());
    expect(strict.ok).toBe(true);
    expect(strict.errors).toEqual([]);
    const reExported = convertAstToMarkdown(await again.getAST(), again);
    expect(reExported).toBe(exported);
  });
});

describe('compose alert blockquotes', () => {
  it('compiles every alert tag to a callout op with the matching theme', () => {
    const cases: Array<[string, string]> = [
      ['NOTE', 'indigo'],
      ['TIP', 'emerald'],
      ['IMPORTANT', 'slate'],
      ['WARNING', 'amber'],
      ['CAUTION', 'rose'],
    ];
    for (const [tag, theme] of cases) {
      expect(compileMarkdown(`> [!${tag}]\n> Body text\n`)).toEqual([
        { action: 'callout', anchor: 'end', position: 'after', text: 'Body text', theme },
      ]);
    }
  });

  it('renders titled alerts as callout containers that validate', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, '> [!WARNING]\n> **Watch out**\n> body here\n');
    expect(result.success).toBe(true);
    const xml = documentXml(parser);
    expect(xml).toContain('FFFBEB');
    expect(xml).toContain('Watch out');
    expect(xml).toContain('body here');
    const strict = await validateDocxBuffer(parser.rezip());
    expect(strict.ok).toBe(true);
    expect(strict.errors).toEqual([]);
  });

  it('still fails general blockquotes and malformed alerts with line numbers', () => {
    expectLineError(() => compileMarkdown('> quoted\n'), 1, 'blockquotes');
    expectLineError(() => compileMarkdown('> [!BOGUS]\n> x\n'), 1, 'blockquotes');
    expectLineError(() => compileMarkdown('> [!NOTE]\n'), 1, 'needs body text');
    expectLineError(() => compileMarkdown('> [!NOTE]\n> - item\n'), 1, 'list blocks');
    expectLineError(() => compileMarkdown('> [!NOTE]\n> use **bold** here\n'), 2, 'rich text inside alerts');
    expectLineError(() => compileMarkdown('> [!NOTE]\n> see [t](https://example.com)\n'), 2, 'links inside alerts');
    expectLineError(() => compileMarkdown('> [!NOTE]\n> note[^a]\n\n[^a]: body\n'), 2, 'footnote markers inside alerts');
  });

  it('round-trips composed alert content back to valid and stable markdown', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const first = await composeDocument(parser, '> [!WARNING]\n> **Watch out**\n> body here\n');
    expect(first.success).toBe(true);
    // The exporter re-emits the callout from reparsed XML without its
    // in-memory theme metadata, so it lands as a plain table here; the title
    // and body text must survive that shape change intact.
    const exported = convertAstToMarkdown(await parser.getAST(), parser);
    expect(exported).toContain('Watch out');
    expect(exported).toContain('body here');

    const again = DocxParser.createNewDocument('blank');
    const second = await composeDocument(again, exported);
    expect(second.success).toBe(true);
    const strict = await validateDocxBuffer(again.rezip());
    expect(strict.ok).toBe(true);
    expect(strict.errors).toEqual([]);
    const reExported = convertAstToMarkdown(await again.getAST(), again);
    expect(reExported).toBe(exported);
  });
});

describe('compose rollback keeps relationships clean', () => {
  it('leaves zero orphan rels when seed removal fails after links commit', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const before = parser.rezip();
    const relsBefore =
      new AdmZip(before).getEntry('word/_rels/document.xml.rels')?.getData().toString('utf8') ?? '';
    const original = OperationEngine.prototype.executeBatch;
    let restoreCalls = 0;
    const innerRestore = parser.restorePackageSnapshot.bind(parser);
    parser.restorePackageSnapshot = (snap: Map<string, Buffer>) => {
      restoreCalls += 1;
      return innerRestore(snap);
    };
    OperationEngine.prototype.executeBatch = async function (
      this: OperationEngine,
      req: { operations: Array<{ action?: string }>; sequential?: boolean }
    ): Promise<Awaited<ReturnType<OperationEngine['executeBatch']>>> {
      const ops = req.operations as Array<{ action?: string }>;
      if (ops.some((op) => op.action === 'delete')) {
        return {
          success: false,
          appliedCount: 0,
          errors: [{ operationIndex: 0, error: 'injected delete failure', hint: 'retry the compose' }],
          rollbackTriggered: true,
          skippedIds: [],
          warnings: [],
          revisionsPending: 0,
          possiblyStaleBodyRefs: [],
        } as unknown as Awaited<ReturnType<OperationEngine['executeBatch']>>;
      }
      return original.call(this, req as never) as Promise<Awaited<ReturnType<OperationEngine['executeBatch']>>>;
    };
    try {
      const result = await composeDocument(parser, '# Composed title\n\nSee [docs](https://example.com) here\n');
      expect(result.success).toBe(false);
      expect(result.rollbackTriggered).toBe(true);
      expect(result.appliedCount).toBe(0);
      expect(restoreCalls).toBeGreaterThan(0);
      expect(parser.rezip().equals(before)).toBe(true);
      const relsAfter =
        new AdmZip(parser.rezip()).getEntry('word/_rels/document.xml.rels')?.getData().toString('utf8') ?? '';
      expect(relsAfter).toBe(relsBefore);
      expect(relsAfter).not.toContain('https://example.com');
    } finally {
      OperationEngine.prototype.executeBatch = original;
    }
  });
});

describe('compose span newlines', () => {
  it('normalizes newlines to spaces in compiled spans and emitted runs', async () => {
    const ops = compileMarkdown('Hello **bold\nsplit** end\n');
    const insert = ops.find((op) => op.action === 'insert') as unknown as {
      params?: { mdSpans?: Array<{ text: string }> };
    };
    for (const span of insert.params?.mdSpans ?? []) {
      expect(span.text).not.toContain('\n');
    }

    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, 'Hello **bold\nsplit** and `code\nspan` end\n');
    expect(result.success).toBe(true);
    const xml = documentXml(parser);
    const runs = [...xml.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((match) => match[1]);
    expect(runs.length).toBeGreaterThan(0);
    for (const run of runs) {
      expect(run).not.toContain('\n');
    }
    expect(xml).toContain('>bold split<');
    expect(xml).toContain('>code span<');
    const strict = await validateDocxBuffer(parser.rezip());
    expect(strict.ok).toBe(true);
    expect(strict.errors).toEqual([]);
  });
});

describe('compose raw input ceiling', () => {
  it('exposes a byte ceiling inside the agreed range', () => {
    expect(COMPOSE_MARKDOWN_MAX_BYTES).toBe(3_000_000);
    expect(COMPOSE_MARKDOWN_MAX_BYTES).toBeGreaterThanOrEqual(2_000_000);
    expect(COMPOSE_MARKDOWN_MAX_BYTES).toBeLessThanOrEqual(5_000_000);
  });

  it('refuses oversize input with a named limit and remedy', () => {
    const big = 'a\n\n'.repeat(Math.ceil((COMPOSE_MARKDOWN_MAX_BYTES + 100) / 3));
    expect(Buffer.byteLength(big, 'utf8')).toBeGreaterThan(COMPOSE_MARKDOWN_MAX_BYTES);
    let message = '';
    let code: unknown;
    let remedy: unknown;
    try {
      compileMarkdown(big);
    } catch (err: unknown) {
      message = err instanceof Error ? err.message : String(err);
      code = (err as { code?: unknown }).code;
      remedy = (err as { remedy?: unknown }).remedy;
    }
    expect(code).toBe('E_LIMIT');
    expect(message).toContain('E_LIMIT');
    expect(message).toContain(String(COMPOSE_MARKDOWN_MAX_BYTES));
    expect(message).toContain('Remedy:');
    expect(String(remedy)).toContain('smaller parts');
  });

  it('writes nothing when the ceiling trips', async () => {
    const big = 'b\n\n'.repeat(Math.ceil((COMPOSE_MARKDOWN_MAX_BYTES + 100) / 3));
    const parser = DocxParser.createNewDocument('blank');
    const before = parser.rezip();
    await expect(composeDocument(parser, big)).rejects.toMatchObject({ code: 'E_LIMIT' });
    expect(parser.rezip().equals(before)).toBe(true);
  });

  it('counts multibyte characters as bytes', () => {
    const unit = 'é';
    expect(Buffer.byteLength(unit, 'utf8')).toBe(2);
    const count = Math.ceil((COMPOSE_MARKDOWN_MAX_BYTES + 10) / 2);
    const big = `${'a\n\n'}${unit.repeat(count)}`;
    expect(Buffer.byteLength(big, 'utf8')).toBeGreaterThan(COMPOSE_MARKDOWN_MAX_BYTES);
    expect(() => compileMarkdown(big)).toThrowError(/E_LIMIT/);
  });

  it('normal manuscripts compile under the ceiling', () => {
    const ops = compileMarkdown('# Title\n\nBody text\n');
    expect(ops.length).toBeGreaterThan(0);
    expect(Buffer.byteLength('# Title\n\nBody text\n', 'utf8')).toBeLessThan(
      COMPOSE_MARKDOWN_MAX_BYTES
    );
  });
});

describe('compose code fence round-trip', () => {
  it('exports fences with their language and reaches a stable fixpoint', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, '```js\nconst x = 1;\nconsole.log(x);\n```\n');
    expect(result.success).toBe(true);
    const ast = await parser.getAST();
    const first = convertAstToMarkdown(ast);
    expect(first.trim()).toBe('```js\nconst x = 1;\nconsole.log(x);\n```');

    const parser2 = DocxParser.createNewDocument('blank');
    const second = await composeDocument(parser2, first);
    expect(second.success).toBe(true);
    expect(convertAstToMarkdown(await parser2.getAST())).toBe(first);
  });

  it('exports single-line fences verbatim and reaches a stable fixpoint', async () => {
    for (const source of ['```js\nconsole.log("hello");\n```\n', '```\njust code\n```\n']) {
      const parser = DocxParser.createNewDocument('blank');
      const result = await composeDocument(parser, source);
      expect(result.success).toBe(true);
      const first = convertAstToMarkdown(await parser.getAST());
      expect(first).toBe(source);

      const parser2 = DocxParser.createNewDocument('blank');
      const second = await composeDocument(parser2, first);
      expect(second.success).toBe(true);
      expect(convertAstToMarkdown(await parser2.getAST())).toBe(first);
    }
  });
});

describe('compose image alt text', () => {
  const png1x1 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

  it('preserves alt into the drawing description and warns when missing', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, `![Sales chart](data:image/png;base64,${png1x1})\n`);
    expect(result.success).toBe(true);
    expect(result.warnings ?? []).toEqual([]);
    expect(documentXml(parser)).toContain('descr="Sales chart"');
    const md = convertAstToMarkdown(await parser.getAST());
    expect(md).toContain('![Sales chart](data:image/png;base64,');

    const bare = DocxParser.createNewDocument('blank');
    const noAlt = await composeDocument(bare, `![](data:image/png;base64,${png1x1})\n`);
    expect(noAlt.success).toBe(true);
    expect((noAlt.warnings ?? []).join(' ')).toContain('no alt text');
    const bareMd = convertAstToMarkdown(await bare.getAST());
    expect(bareMd).not.toMatch(/!\[[0-9A-Fa-f_]+\]\(/);
  });
});

describe('compose title block and stray alerts', () => {
  it('makes an italic-only line directly under the title the Subtitle', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, '# Report\n\n*Quarterly review*\n\nBody text.\n');
    expect(result.success).toBe(true);
    const xml = documentXml(parser);
    expect(xml).toContain('<w:pStyle w:val="Subtitle"/>');
    // The style carries the italics; the run holds no direct emphasis.
    expect(xml).toMatch(/<w:r><w:t>Quarterly review<\/w:t><\/w:r>/);
    expect(xml.match(/w:val="Subtitle"/g)).toHaveLength(1);
  });

  it('leaves italic lines elsewhere, and mixed lines under the title, as body text', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(
      parser,
      '# Title\n\nSome *mixed* line.\n\n*late italic*\n\n## Section\n\n*under a section*\n',
    );
    expect(result.success).toBe(true);
    expect(documentXml(parser)).not.toContain('w:val="Subtitle"');
  });

  it('warns when [!NOTE] is not inside a blockquote, and still writes the text', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, '# T\n\n[!NOTE] Remember this.\n\n> [!TIP]\n> Fine.\n');
    expect(result.success).toBe(true);
    const stray = result.warnings.filter((w) => w.includes('[!NOTE]'));
    expect(stray).toHaveLength(1);
    expect(stray[0]).toMatch(/line 3/);
    expect(stray[0]).toMatch(/blockquote/);
    expect(result.warnings.some((w) => w.includes('[!TIP]'))).toBe(false);
    expect(documentXml(parser)).toContain('[!NOTE] Remember this.');
  });
});

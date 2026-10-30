/**
 * Who and when a revision or comment is attributed to.
 *
 * A tracked cell edit used to ignore the caller's author and always write "ASTra", and every
 * revision carried a fixed 1980 date. Now each revision writer takes an author and a date from the
 * call, falls back to the document default author, and then to the real current time.
 */
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import {
  DocxParser,
  OperationEngine,
  createBlankDocxBuffer,
  executeDelete,
  executeReplace,
  executeTable,
  executeTrackedChange,
  validateOperationStrict,
} from '@attentionform/astra-engine';
import { resolveDefaultAuthor } from '../../packages/engine/src/ooxml/track-changes-emitter';

const saved = {
  on: process.env.ASTRA_REPRODUCIBLE,
  off: process.env.ASTRA_NO_REPRODUCIBLE,
  epoch: process.env.SOURCE_DATE_EPOCH,
  author: process.env.ASTRA_AUTHOR,
};
function restore(key: keyof typeof saved, name: string): void {
  if (saved[key] === undefined) delete process.env[name];
  else process.env[name] = saved[key];
}
beforeEach(() => {
  delete process.env.ASTRA_REPRODUCIBLE;
  delete process.env.ASTRA_NO_REPRODUCIBLE;
  delete process.env.SOURCE_DATE_EPOCH;
  delete process.env.ASTRA_AUTHOR;
});
afterEach(() => {
  restore('on', 'ASTRA_REPRODUCIBLE');
  restore('off', 'ASTRA_NO_REPRODUCIBLE');
  restore('epoch', 'SOURCE_DATE_EPOCH');
  restore('author', 'ASTRA_AUTHOR');
});

async function withTable() {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  await engine.executeBatch({
    operations: [
      { action: 'insert', type: 'table', anchor: ast.sections[0].nodes[0].id, position: 'after', headers: ['H1', 'H2'], rows: [['a', 'b']] },
    ],
  } as any);
  const tableId = ast.sections[0].nodes.find((n: any) => n.type === 'table').id as string;
  return { parser, ast, tableId, xml: () => new AdmZip(parser.rezip(ast)).readAsText('word/document.xml') };
}

const attr = (xml: string, name: string) => [...xml.matchAll(new RegExp(`<w:(?:ins|del) [^>]*w:${name}="([^"]*)"`, 'g'))].map((m) => m[1]);

describe('tracked table cell edit', () => {
  it('uses the author and date the caller gave (it used to write "ASTra" and a fixed date)', async () => {
    const { ast, tableId, xml } = await withTable();
    await executeTable(
      { action: 'table', target: tableId, trackChanges: true, params: { op: 'updateCell', rowIndex: 2, colIndex: 1, text: 'changed', author: 'Dana Reviewer', date: '2026-10-24T09:30:00Z' } } as any,
      ast,
    );
    const out = xml();
    expect(attr(out, 'author').length).toBeGreaterThan(0);
    expect(new Set(attr(out, 'author'))).toEqual(new Set(['Dana Reviewer']));
    expect(new Set(attr(out, 'date'))).toEqual(new Set(['2026-10-24T09:30:00Z']));
  });

  it('accepts author and date flat as well as nested', async () => {
    const { ast, tableId, xml } = await withTable();
    await executeTable(
      { action: 'table', target: tableId, op: 'updateCell', trackChanges: true, rowIndex: 2, colIndex: 2, text: 'x', author: 'Flat Author' } as any,
      ast,
    );
    expect(new Set(attr(xml(), 'author'))).toEqual(new Set(['Flat Author']));
  });

  it('falls back to the document default author, then to the current time', async () => {
    const { ast, tableId, xml } = await withTable();
    ast.settings.defaultAuthor = 'Document Owner';
    const before = Date.now();
    await executeTable({ action: 'table', target: tableId, trackChanges: true, params: { op: 'updateCell', rowIndex: 2, colIndex: 1, text: 'y' } } as any, ast);
    const after = Date.now();
    const out = xml();
    expect(new Set(attr(out, 'author'))).toEqual(new Set(['Document Owner']));
    for (const d of attr(out, 'date')) {
      expect(Date.parse(d)).toBeGreaterThanOrEqual(before);
      expect(Date.parse(d)).toBeLessThanOrEqual(after);
    }
  });

  it('refuses a malformed date before anything is applied', () => {
    expect(() =>
      validateOperationStrict({ action: 'table', target: 't_001', op: 'updateCell', rowIndex: 1, colIndex: 1, text: 'x', date: 'yesterday' }),
    ).toThrow(/date must be an ISO 8601 timestamp/);
  });

  it('does not accept author on a sub-operation that writes no revision', () => {
    expect(() => validateOperationStrict({ action: 'table', target: 't_001', op: 'setBorder', params: { author: 'X' } })).toThrow(
      /unknown parameter "author" for table setBorder/,
    );
  });
});

describe('every revision writer takes author and date', () => {
  it('replace, delete and trackedChange insert honour both and agree', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const first = ast.sections[0].nodes[0];
    const date = '2025-03-04T05:06:07Z';
    executeReplace({ action: 'replace', target: first.id, text: 'new words', trackChanges: true, author: 'R', date } as any, ast);
    let out = new AdmZip(parser.rezip(ast)).readAsText('word/document.xml');
    expect(new Set(attr(out, 'author'))).toEqual(new Set(['R']));
    expect(new Set(attr(out, 'date'))).toEqual(new Set([date]));

    const p2 = new DocxParser(createBlankDocxBuffer('blank'));
    const ast2: any = await p2.getAST();
    executeDelete({ action: 'delete', target: ast2.sections[0].nodes[0].id, trackChanges: true, author: 'D', date } as any, ast2, p2);
    out = new AdmZip(p2.rezip(ast2)).readAsText('word/document.xml');
    expect(new Set(attr(out, 'author'))).toEqual(new Set(['D']));
    expect(new Set(attr(out, 'date'))).toEqual(new Set([date]));

    const p3 = new DocxParser(createBlankDocxBuffer('blank'));
    const ast3: any = await p3.getAST();
    executeTrackedChange({ action: 'trackedChange', operation: 'insert', target: ast3.sections[0].nodes[0].id, text: 'z', author: 'T', date } as any, ast3, p3);
    out = new AdmZip(p3.rezip(ast3)).readAsText('word/document.xml');
    expect(new Set(attr(out, 'author'))).toEqual(new Set(['T']));
    expect(new Set(attr(out, 'date'))).toEqual(new Set([date]));
  });

  it('a tracked paragraph insert takes author and date', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    await engine.executeBatch({
      operations: [
        { action: 'insert', type: 'paragraph', anchor: ast.sections[0].nodes[0].id, position: 'after', text: 'tracked', trackChanges: true, author: 'Ins Author', date: '2024-01-02T03:04:05Z' },
      ],
    } as any);
    const out = new AdmZip(parser.rezip(ast)).readAsText('word/document.xml');
    expect(attr(out, 'author')).toContain('Ins Author');
    expect(attr(out, 'date')).toContain('2024-01-02T03:04:05Z');
  });

  it('delete accepts author and date nested in params too', () => {
    expect(() => validateOperationStrict({ action: 'delete', target: 'p_1', trackChanges: true, params: { author: 'A', date: '2025-03-04T05:06:07Z' } })).not.toThrow();
  });
});

describe('the document default author', () => {
  it('is "ASTra" unless ASTRA_AUTHOR names one', async () => {
    expect(resolveDefaultAuthor({})).toBe('ASTra');
    expect(resolveDefaultAuthor({ ASTRA_AUTHOR: '  Lee  ' } as NodeJS.ProcessEnv)).toBe('Lee');
    expect(resolveDefaultAuthor({ ASTRA_AUTHOR: '   ' } as NodeJS.ProcessEnv)).toBe('ASTra');
    process.env.ASTRA_AUTHOR = 'Env Author';
    const ast: any = await new DocxParser(createBlankDocxBuffer('blank')).getAST();
    expect(ast.settings.defaultAuthor).toBe('Env Author');
  });

  it('is what an edit with no author is attributed to', async () => {
    process.env.ASTRA_AUTHOR = 'Env Author';
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    executeReplace({ action: 'replace', target: ast.sections[0].nodes[0].id, text: 'q', trackChanges: true } as any, ast);
    expect(new Set(attr(new AdmZip(parser.rezip(ast)).readAsText('word/document.xml'), 'author'))).toEqual(new Set(['Env Author']));
  });
});

describe('an author with characters XML cannot hold', () => {
  // eslint-disable-next-line no-control-regex -- intentional: asserting that no XML-forbidden control character survives
  const CONTROL = /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/;

  it('loses them from the environment default, and falls back when nothing readable is left', () => {
    expect(resolveDefaultAuthor({ ASTRA_AUTHOR: 'Ann\u0001 Lee\u001f' } as NodeJS.ProcessEnv)).toBe('Ann Lee');
    expect(resolveDefaultAuthor({ ASTRA_AUTHOR: '\u0001\u0002' } as NodeJS.ProcessEnv)).toBe('ASTra');
  });

  it('loses them from an operation author, and from the document default when the operation names none', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const target = ast.sections[0].nodes[0].id;
    executeReplace({ action: 'replace', target, text: 'one', trackChanges: true, author: 'Bob\u0001 Ray' } as any, ast);
    ast.settings.defaultAuthor = 'Cy\u0002 Lo';
    executeReplace({ action: 'replace', target, text: 'two', trackChanges: true } as any, ast);
    const xml = new AdmZip(parser.rezip(ast)).readAsText('word/document.xml');
    expect(xml).not.toMatch(CONTROL);
    expect(new Set(attr(xml, 'author'))).toEqual(new Set(['Bob Ray', 'Cy Lo']));
  });

  it('never reaches the core properties of a new document', () => {
    process.env.ASTRA_AUTHOR = 'Ann\u0001 Lee';
    const core = new AdmZip(DocxParser.createNewDocument('blank').rezip()).readAsText('docProps/core.xml');
    expect(core).not.toMatch(CONTROL);
    expect(core).toContain('<dc:creator>Ann Lee</dc:creator>');
    expect(core).toContain('<cp:lastModifiedBy>Ann Lee</cp:lastModifiedBy>');
  });
});

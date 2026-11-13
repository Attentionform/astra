import { describe, it, expect, vi, afterEach, beforeEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { AdmZip, DocxParser, ToonGenerator, createBlankDocxBuffer } from '@attentionform/astra-engine';
import { handleExecuteBatch, handleInspectDocument } from '../../packages/mcp-server/src/lib/mcp-tool-adapter';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';
import { run } from '../../packages/cli/src/cli';

/**
 * The `outline` tier lists headings only. A heading is a top-level body paragraph whose own style id contains
 * `heading`, `title` or `subtitle` (the rule of the structural summary). `ast.sections` are the file's page
 * sections, so the Section column is built from the headings themselves, nested by level.
 */
const para = (text: string, style?: string): string =>
  `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`;
const TABLE =
  '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>cell text</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';

/** The blank template's own Title and Subtitle are dropped, so the body is exactly `body`. */
function docWith(body: string): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  const xml = zip.readAsText('word/document.xml');
  const start = xml.indexOf('<w:body>') + '<w:body>'.length;
  const at = xml.lastIndexOf('<w:sectPr');
  zip.updateFile('word/document.xml', Buffer.from(`${xml.slice(0, start)}${body}${xml.slice(at)}`, 'utf8'));
  return zip.toBuffer();
}

const SAMPLE =
  para('Alpha', 'Heading1') + para('body a1') + para('body a2') +
  para('Beta', 'Heading2') + para('body b1') + TABLE +
  para('Deep', 'Heading4') + para('body d1') +
  para('Gamma', 'Heading2') + para('body g1') + para('body g2') + para('body g3') +
  para('Omega', 'Heading1') + para('body o1') + para('Sub', 'Subtitle') + para('body o2');

async function load(body = SAMPLE) {
  const parser = new DocxParser(docWith(body));
  const ast = await parser.getAST();
  return { parser, ast, toon: new ToonGenerator() };
}

describe('outline tier (engine)', () => {
  it('lists exactly the headings with id, level, style and the headings above', async () => {
    const { ast, toon } = await load();
    const table = toon.toon(ast, { fidelity: 'outline' });
    expect(table.tier).toBe('outline');
    expect(table.headers).toEqual(['ID', 'L', 'Style', 'Text', 'Section']);
    expect(table.rows.map((r) => [r.text, r.level, r.style, r.path])).toEqual([
      ['Alpha', 1, 'Heading1', ''],
      ['Beta', 2, 'Heading2', 'Alpha'],
      ['Deep', 4, 'Heading4', 'Alpha > Beta'],
      ['Gamma', 2, 'Heading2', 'Alpha'],
      ['Omega', 1, 'Heading1', ''],
      ['Sub', 2, 'Subtitle', 'Omega'],
    ]);
    const text = toon.renderToonText(table);
    expect(text).toContain('ID | L | Style | Text | Section');
    expect(text).not.toContain('body a1');
    expect(text).not.toContain('cell text');
  });

  it('agrees with the structural summary and with compact about which nodes are headings, and their text', async () => {
    const { parser, ast, toon } = await load();
    const outline = toon.toon(ast, { fidelity: 'outline' }).rows;
    const fromSummary = parser.getStructuralSummary().nodes.filter((n) => n.isHeading);
    expect(outline.map((r) => [r.id, r.level])).toEqual(fromSummary.map((n) => [n.id, n.headingLevel]));
    const compact = new Map(toon.toon(ast, { fidelity: 'compact' }).rows.map((r) => [r.id, r]));
    for (const row of outline) {
      expect(compact.get(row.id)?.text).toBe(row.text);
      expect(compact.get(row.id)?.style).toBe(row.style);
    }
  });

  it('recognition: Title 1, Subtitle 2, Heading4 4, TOCHeading 1; custom style, h1, cell and textbox host are not listed', async () => {
    const box = '<w:p><w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml"><v:textbox><w:txbxContent><w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>in a box</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>';
    const { ast, toon } = await load(
      para('T', 'Title') + para('S', 'Subtitle') + para('H4', 'Heading4') + para('TOC', 'TOCHeading') +
        para('K', 'Kapitel') + para('h', 'h1') + TABLE + box
    );
    const rows = toon.toon(ast, { fidelity: 'outline' }).rows;
    expect(rows.map((r) => [r.text, r.level])).toEqual([['T', 1], ['S', 2], ['H4', 4], ['TOC', 1]]);
  });

  it('cuts text at 80 characters by default and at maxTextLength when given', async () => {
    const long = 'x'.repeat(120);
    const { ast, toon } = await load(para(long, 'Heading1'));
    expect(toon.toon(ast, { fidelity: 'outline' }).rows[0].text).toBe(`${'x'.repeat(80)}...`);
    expect(toon.toon(ast, { fidelity: 'outline', maxTextLength: 10 }).rows[0].text).toBe(`${'x'.repeat(10)}...`);
  });

  it('a sectionId that is a node id lists that node only; the parents above it are still shown', async () => {
    const { ast, toon } = await load();
    const deep = toon.toon(ast, { fidelity: 'outline' }).rows.find((r) => r.text === 'Deep')!;
    const scoped = toon.toon(ast, { fidelity: 'outline', sectionId: deep.id }).rows;
    expect(scoped.map((r) => [r.text, r.path])).toEqual([['Deep', 'Alpha > Beta']]);
    const body = ast.sections[0].nodes.find((n) => n.metadata.text === 'body a1')!;
    expect(toon.toon(ast, { fidelity: 'outline', sectionId: body.id }).rows).toEqual([]);
  });

  it('a document without headings gives no rows', async () => {
    const { ast, toon } = await load(para('just text') + para('more text'));
    expect(toon.toon(ast, { fidelity: 'outline' }).rows).toEqual([]);
  });

  it('compact and standard output do not change', async () => {
    const { ast, toon } = await load();
    const before = ['compact', 'standard'].map((f) => toon.renderToonText(toon.toon(ast, { fidelity: f as 'compact' | 'standard' })));
    toon.toon(ast, { fidelity: 'outline' });
    const after = ['compact', 'standard'].map((f) => toon.renderToonText(toon.toon(ast, { fidelity: f as 'compact' | 'standard' })));
    expect(after).toEqual(before);
    expect(before[0]).toContain('body a1');
  });

  it('rejects an unknown tier and names outline', async () => {
    const { ast, toon } = await load();
    expect(() => toon.toon(ast, { fidelity: 'headings' as never })).toThrow(/outline\|compact\|standard\|full/);
  });
});

describe('outline cells: separators, backslashes and line breaks in a heading', () => {
  const decodeCell = (cell: string): string => cell.replace(/\\(\\|\||n|r|t)/g, (_m, c: string) => (c === 'n' ? '\n' : c === 'r' ? '\r' : c === 't' ? '\t' : c));
  const NAMES = ['Revenue | 2025', 'back\\slash', 'ends with\\', 'a\\| both', 'line\nbreak', '| leading pipe', 'tab\there', 'a > b'];
  const body = NAMES.map((n, i) => para(n.replace(/&/g, '&amp;').replace(/</g, '&lt;'), i === 0 ? 'Heading1' : 'Heading2')).join('') + para('Child of last', 'Heading3');

  it('each row has exactly five cells and the cells decode to the heading text', async () => {
    const { ast, toon } = await load(body);
    const text = toon.renderToonText(toon.toon(ast, { fidelity: 'outline', maxTextLength: 200 }));
    const lines = text.split(/\n(?=[0-9A-F]{8} \| )/).slice(1).map((l) => l.replace(/\n$/, ''));
    expect(lines).toHaveLength(NAMES.length + 1);
    lines.slice(0, NAMES.length).forEach((line, i) => {
      const cells = line.split(' | ');
      expect(cells, line).toHaveLength(5);
      expect(decodeCell(cells[3])).toBe(NAMES[i]);
    });
    expect(text.split('\n').every((l) => !l.includes('\r'))).toBe(true);
  });

  it('a heading with a pipe does not break the Section of the heading under it', async () => {
    const { ast, toon } = await load(para('Revenue | 2025', 'Heading1') + para('Q1', 'Heading2'));
    const line = toon.renderToonText(toon.toon(ast, { fidelity: 'outline' })).split('\n').find((l) => l.includes('| Q1 |'))!;
    const cells = line.split(' | ');
    expect(cells).toHaveLength(5);
    expect(cells[4]).toBe('Revenue \\| 2025');
  });

  it('a > inside an ancestor heading is escaped so the parts can be told apart', async () => {
    const { ast, toon } = await load(para('A > B', 'Heading1') + para('C', 'Heading2') + para('D', 'Heading3'));
    const rows = toon.toon(ast, { fidelity: 'outline' }).rows;
    expect(rows[2].path).toBe('A \\> B > C');
  });

  it('the CLI prints the same escaped cells', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-outline-esc-'));
    const file = path.join(dir, 'doc.docx');
    fs.writeFileSync(file, docWith(para('Revenue | 2025', 'Heading1') + para('Q1', 'Heading2')));
    const seen: string[] = [];
    const log = vi.spyOn(console, 'log').mockImplementation((...a: unknown[]) => void seen.push(a.join(' ')));
    try {
      expect(await run(['inspect', file, '--outline'])).toBe(0);
    } finally {
      log.mockRestore();
      fs.rmSync(dir, { recursive: true, force: true });
    }
    const out = seen.join('\n');
    expect(out).toContain('| Heading1 | Revenue \\| 2025 | ');
    expect(out).toContain('| Heading2 | Q1 | Revenue \\| 2025');
  });
});

describe('outline tier through inspect_document', () => {
  async function open(body = SAMPLE) {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(docWith(body), 'outline.docx');
    return { store, handle };
  }
  const call = async (store: ReturnType<typeof createMcpSessionStore>, sessionId: string, extra: Record<string, unknown> = {}) => {
    const res = await handleInspectDocument(store, { sessionId, fidelity: 'outline', ...extra });
    if (res.ok !== true) throw new Error(`inspect failed: ${JSON.stringify(res)}`);
    return res;
  };
  const idOf = (handle: { ast: { sections: Array<{ nodes: Array<{ id: string; metadata: { text?: string } }> }> } }, text: string): string =>
    handle.ast.sections.flatMap((s) => s.nodes).find((n) => n.metadata.text === text)!.id;

  it('returns heading rows and the whole-document counts', async () => {
    const { store, handle } = await open();
    const res = await call(store, handle.sessionId);
    expect(res.fidelity).toBe('outline');
    expect(res.rowCount).toBe(6);
    expect(res.totalNodes).toBe(6);
    expect(res.outline).toMatchObject({ headings: 6, tables: 1 });
    expect(res.outline!.nodes).toBeGreaterThan(6);
    expect(res.outline!.words).toBeGreaterThan(0);
    expect(res.toonText).toContain('Alpha');
    expect(res.toonText).not.toContain('body a1');
  });

  it('other tiers carry no outline field', async () => {
    const { store, handle } = await open();
    const res = await handleInspectDocument(store, { sessionId: handle.sessionId, fidelity: 'compact' });
    expect(res).not.toHaveProperty('outline');
  });

  it('pages over heading rows; counts stay whole-document', async () => {
    const many = Array.from({ length: 120 }, (_, i) => para(`Head ${i}`, 'Heading1') + para(`body ${i}`)).join('');
    const { store, handle } = await open(many);
    // A small outline with no paging arguments comes back whole; 40 is not the default limit, so it pages.
    const whole = await call(store, handle.sessionId);
    expect(whole).toMatchObject({ totalNodes: 120, returnedNodes: 120, hasMore: false });
    const first = await call(store, handle.sessionId, { limit: 40 });
    expect(first).toMatchObject({ totalNodes: 120, returnedNodes: 40, hasMore: true, nextOffset: 40 });
    expect(first.outline).toMatchObject({ headings: 120, nodes: 240 });
    const third = await call(store, handle.sessionId, { offset: 80, limit: 40 });
    expect(third).toMatchObject({ returnedNodes: 40, hasMore: false, nextOffset: null, totalNodes: 120 });
    const beyond = await handleInspectDocument(store, { sessionId: handle.sessionId, fidelity: 'outline', offset: 500, limit: 40 });
    expect(beyond.ok).not.toBe(true);
  });

  it('aroundNode on a heading centres on it', async () => {
    const { store, handle } = await open();
    const res = await call(store, handle.sessionId, { aroundNode: idOf(handle, 'Gamma'), window: 1 });
    expect(res.toonText).toContain('Beta');
    expect(res.toonText).toContain('Gamma');
    expect(res.toonText).toContain('Omega');
    expect(res.aroundResolvedTo).toBeUndefined();
  });

  it('aroundNode on a body paragraph resolves to the heading above it', async () => {
    const { store, handle } = await open();
    const res = await call(store, handle.sessionId, { aroundNode: idOf(handle, 'body g2'), window: 1 });
    expect(res.aroundResolvedTo).toBe(idOf(handle, 'Gamma'));
    expect(res.note).toMatch(/not a heading/);
    expect(res.toonText).toContain('Gamma');
    expect(res.toonText).not.toMatch(/\| Alpha \|/);
  });

  it('aroundNode before the first heading shows the first rows and says so', async () => {
    const { store, handle } = await open(para('lead in') + SAMPLE);
    const res = await call(store, handle.sessionId, { aroundNode: idOf(handle, 'lead in'), window: 1 });
    expect(res.aroundResolvedTo).toBeNull();
    expect(res.note).toMatch(/no heading above/);
    expect(res.toonText).toContain('Alpha');
  });

  it('aroundNode in a document with no headings returns an empty outline with the explanation', async () => {
    const { store, handle } = await open(para('just text'));
    const res = await call(store, handle.sessionId, { aroundNode: idOf(handle, 'just text') });
    expect(res.aroundResolvedTo).toBeNull();
    expect(res.note).toMatch(/No headings/);
    expect(res.returnedNodes).toBe(0);
  });

  it('aroundNode with an unknown id is an error', async () => {
    const { store, handle } = await open();
    const res = await handleInspectDocument(store, { sessionId: handle.sessionId, fidelity: 'outline', aroundNode: 'NOPE0000' });
    expect(res.ok).not.toBe(true);
  });

  it('aroundNode outside the sectionId scope is rejected; one inside it resolves within it', async () => {
    const { store, handle } = await open();
    const sec = handle.ast.sections[0].id;
    const outside = await handleInspectDocument(store, {
      sessionId: handle.sessionId, fidelity: 'outline', sectionId: idOf(handle, 'Gamma'), aroundNode: idOf(handle, 'body a1'),
    });
    expect(outside.ok).not.toBe(true);
    expect(JSON.stringify(outside)).toContain('is not in section');
    const inside = await call(store, handle.sessionId, { sectionId: sec, aroundNode: idOf(handle, 'body b1'), window: 1 });
    expect(inside.aroundResolvedTo).toBe(idOf(handle, 'Beta'));
  });

  it('cache and counts follow edits', async () => {
    const { store, handle } = await open();
    const rowsOf = (r: { toonText: string }): string => r.toonText.split('\n').filter((l) => /\| Heading|\| Subtitle/.test(l)).join('\n');
    const first = await call(store, handle.sessionId);
    const again = await call(store, handle.sessionId);
    expect(again.toonText).toBe(first.toonText);
    // A body edit that adds words: heading rows identical, words count up by the words added, output equals a fresh read.
    const added = 'one two three four five';
    const edit = await handleExecuteBatch(store, { sessionId: handle.sessionId, operations: [{ action: 'replace', target: idOf(handle, 'body a1'), text: added }], sequential: true });
    expect(edit).toMatchObject({ ok: true });
    const afterBody = await call(store, handle.sessionId);
    expect(rowsOf(afterBody)).toBe(rowsOf(first));
    expect(afterBody.outline!.words - first.outline!.words).toBe(added.split(' ').length - 2);
    // A heading edit: the changed row, and the path of the heading under it, change.
    await handleExecuteBatch(store, { sessionId: handle.sessionId, operations: [{ action: 'replace', target: idOf(handle, 'Alpha'), text: 'Renamed' }], sequential: true });
    const afterHeading = await call(store, handle.sessionId);
    expect(afterHeading.toonText).toContain('Renamed');
    expect(afterHeading.toonText).not.toContain('| Alpha |');
    expect(afterHeading.toonText).toContain('Renamed > Beta');
  });
});

describe('scoped outline cache follows the headings outside the scope', () => {
  it('editing an ancestor heading changes the breadcrumb of a scoped outline', async () => {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(docWith(SAMPLE), 'scoped.docx');
    const idOf = (text: string): string => handle.ast.sections.flatMap((s) => s.nodes).find((n) => n.metadata.text === text)!.id;
    const deep = idOf('Deep');
    const read = async (): Promise<string> => {
      const res = await handleInspectDocument(store, { sessionId: handle.sessionId, fidelity: 'outline', sectionId: deep });
      if (res.ok !== true) throw new Error(JSON.stringify(res));
      return res.toonText;
    };
    expect(await read()).toContain('Alpha > Beta');
    for (const [target, from, to] of [[idOf('Alpha'), 'Alpha', 'Renamed'], [idOf('Beta'), 'Beta', 'Second']] as const) {
      const edit = await handleExecuteBatch(store, { sessionId: handle.sessionId, operations: [{ action: 'replace', target, text: to }], sequential: true });
      expect(edit).toMatchObject({ ok: true });
      const text = await read();
      expect(text, `after ${from} -> ${to}`).not.toContain(from === 'Alpha' ? 'Alpha >' : '> Beta');
      expect(text).toContain(to);
    }
    expect(await read()).toContain('Renamed > Second');
  });

  it('a style change that turns an ancestor into a non-heading changes the scoped breadcrumb', async () => {
    const { ast } = await load();
    const toon = new ToonGenerator();
    const deep = ast.sections[0].nodes.find((n) => n.metadata.text === 'Deep')!;
    const before = toon.toon(ast, { fidelity: 'outline', sectionId: deep.id }).rows[0].path;
    expect(before).toBe('Alpha > Beta');
    const beta = ast.sections[0].nodes.find((n) => n.metadata.text === 'Beta')!;
    beta.xml.getElementsByTagName('w:pStyle')[0].setAttribute('w:val', 'Normal');
    const after = toon.toon(ast, { fidelity: 'outline', sectionId: deep.id }).rows[0].path;
    expect(after).toBe('Alpha');
  });
});

describe('astra inspect --outline', () => {
  let dir: string;
  let file: string;
  let seen: string[];
  beforeEach(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-outline-'));
    file = path.join(dir, 'doc.docx');
    fs.writeFileSync(file, docWith(SAMPLE));
    seen = [];
    vi.spyOn(console, 'log').mockImplementation((...a: unknown[]) => void seen.push(a.join(' ')));
    vi.spyOn(console, 'error').mockImplementation((...a: unknown[]) => void seen.push(a.join(' ')));
  });
  afterEach(() => {
    vi.restoreAllMocks();
    fs.rmSync(dir, { recursive: true, force: true });
  });

  it('--outline prints heading rows and the footer says what the rows are', async () => {
    expect(await run(['inspect', file, '--outline'])).toBe(0);
    const out = seen.join('\n');
    expect(out).toContain('Alpha');
    expect(out).not.toContain('body a1');
    expect(out).toContain('outline: rows are headings (6 listed of');
  });

  it('--outline equals --fidelity outline', async () => {
    await run(['inspect', file, '--outline']);
    const a = seen.join('\n');
    seen.length = 0;
    await run(['inspect', file, '--fidelity', 'outline']);
    expect(seen.join('\n')).toBe(a);
  });

  it('--json carries the outline counts', async () => {
    expect(await run(['--json', 'inspect', file, '--outline'])).toBe(0);
    const json = JSON.parse(seen.filter((l) => l.trim().startsWith('{')).pop()!);
    expect(json.fidelity).toBe('outline');
    expect(json.outline).toMatchObject({ headings: 6, tables: 1 });
  });

  it('a bad --fidelity lists outline; --outline with another --fidelity is a usage error', async () => {
    expect(await run(['inspect', file, '--fidelity', 'headings'])).toBe(2);
    expect(seen.join('\n')).toContain('outline|compact|standard|full');
    seen.length = 0;
    expect(await run(['inspect', file, '--outline', '--fidelity', 'compact'])).toBe(2);
  });
});

describe('outline recognises a heading whatever type the parser gave the paragraph', () => {
  const hd = (inner: string, text: string, extra = ''): string => `<w:p><w:pPr><w:pStyle w:val="Heading2"/>${extra}</w:pPr>${inner}<w:r><w:t>${text}</w:t></w:r></w:p>`;
  const EQUATION = '<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara>';
  const CHART =
    '<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId99"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  const body =
    para('Top', 'Heading1') + hd(EQUATION, 'Eq head') + para('under eq', 'Heading3') +
    hd(CHART, 'Chart head') + para('under chart', 'Heading3') +
    hd('', 'Numbered head', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') + para('under numbered', 'Heading3');

  it('lists the same headings as the structural summary, with the right breadcrumbs and count', async () => {
    const { parser, ast, toon } = await load(body);
    const types = new Set(ast.sections[0].nodes.map((n) => n.type));
    expect(types.has('equation') && types.has('chart'), [...types].join()).toBe(true); // the case under test
    const rows = toon.toon(ast, { fidelity: 'outline' }).rows;
    const fromSummary = parser.getStructuralSummary().nodes.filter((n) => n.isHeading);
    expect(rows.map((r) => [r.id, r.level])).toEqual(fromSummary.map((n) => [n.id, n.headingLevel]));
    expect(rows.map((r) => r.text)).toEqual(['Top', 'Eq head', 'under eq', 'Chart head', 'under chart', 'Numbered head', 'under numbered']);
    expect(rows.find((r) => r.text === 'under eq')!.path).toBe('Top > Eq head');
    expect(rows.find((r) => r.text === 'under chart')!.path).toBe('Top > Chart head');
    expect(rows.find((r) => r.text === 'under numbered')!.path).toBe('Top > Numbered head');
  });

  it('inspect_document counts them in outline.headings', async () => {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(docWith(body), 'types.docx');
    const res = await handleInspectDocument(store, { sessionId: handle.sessionId, fidelity: 'outline' });
    if (res.ok !== true) throw new Error(JSON.stringify(res));
    expect(res.outline!.headings).toBe(7);
    expect(res.totalNodes).toBe(7);
  });
});

import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { AdmZip, DocxParser, createBlankDocxBuffer, validateDocxBuffer } from '@attentionform/astra-engine';
import { validateOperationStrict } from '../../packages/engine/src/operation-schema';
import { COLOUR_NAMES, HIGHLIGHT_COLOURS, colourPattern, normalizeColour } from '../../packages/engine/src/operations/colour';
import { run as runCli, EXIT_FAIL, EXIT_OK } from '../../packages/cli/src/cli';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';
import { handleExecuteBatch } from '../../packages/mcp-server/src/lib/mcp-tool-adapter';
import { handleSaveDocument } from '../../packages/mcp-server/src/lib/mcp/handlers/save-close';
import { handleComposeDocument } from '../../packages/mcp-server/src/tools/compose-document';

/**
 * Colour, highlight and size values are checked at the call, with the same rule the writers use. A colour
 * name from the shortlist is mapped to hex; a value no form accepts is refused naming the field.
 * Needs the native validator for the saved-file cases (CI, or OOXML_VALIDATOR_CLI).
 */
const TARGET = 'p_001';
const check = (op: Record<string, unknown>): void => void validateOperationStrict({ action: 'style', target: TARGET, ...op });
const refused = (op: Record<string, unknown>, ...parts: string[]): void => {
  let message = '';
  try {
    check(op);
  } catch (error) {
    message = String((error as Error).message);
  }
  expect(message, JSON.stringify(op)).not.toBe('');
  for (const part of parts) expect(message).toContain(part);
};

describe('rejected at the call', () => {
  it('a colour that no form accepts names the field and the forms', () => {
    refused({ color: 'banana' }, 'color', 'FF0000', 'auto', 'red');
    refused({ color: '#GGGGGG' }, 'color');
    refused({ color: 'F00' }, 'color');
    refused({ color: '' }, 'color');
  });

  it('highlight outside the closed list, fontSize above Word\'s limit', () => {
    refused({ highlight: 'pink' }, 'highlight');
    refused({ highlight: 'Yellow' }, 'highlight');
    refused({ fontSize: 9999 }, '1638 pt');
    refused({ fontSize: '1639pt' }, '1638 pt');
    refused({ fontSize: 3277 }, '1638 pt');
  });

  it('page border: colour and width', () => {
    expect(() => validateOperationStrict({ action: 'pageBorder', color: 'banana' })).toThrow(/color/);
    expect(() => validateOperationStrict({ action: 'pageBorder', sz: '12pt' })).toThrow(/border width/);
    expect(() => validateOperationStrict({ action: 'pageBorder', sz: 97 })).toThrow(/at most 96/);
    expect(() => validateOperationStrict({ action: 'pageBorder', sz: 0 })).toThrow(/at least 1/);
    expect(() => validateOperationStrict({ action: 'pageBorder', sz: 2.5 })).toThrow(/whole number/);
  });

  it('table: border width as a string is not a width, colours name the field', () => {
    expect(() => validateOperationStrict({ action: 'table', target: 't_001', operation: 'setBorder', sz: '12pt' })).toThrow(/border width/);
    expect(() => validateOperationStrict({ action: 'table', target: 't_001', operation: 'setCellShading', fill: 'banana' })).toThrow(/color/);
  });
});

describe('accepted at the call', () => {
  it('hex, with or without #, surrounding spaces, auto, and every shortlisted name', () => {
    for (const color of ['FF0000', '#ff0000', ' FF0000 ', ' #FF0000', 'auto', 'AUTO', 'Red', ...Object.keys(COLOUR_NAMES)]) {
      expect(() => check({ color }), color).not.toThrow();
    }
  });

  it('every highlight value of the closed list', () => {
    expect(HIGHLIGHT_COLOURS).toHaveLength(17);
    for (const highlight of HIGHLIGHT_COLOURS) expect(() => check({ highlight }), highlight).not.toThrow();
  });

  it('font sizes at and below the limit; border widths at the bounds', () => {
    for (const fontSize of ['12pt', '10.5pt', '1638pt', 3276, 24, 0]) expect(() => check({ fontSize }), String(fontSize)).not.toThrow();
    for (const sz of [1, 2, 8, 96]) expect(() => validateOperationStrict({ action: 'pageBorder', sz }), String(sz)).not.toThrow();
  });

  it('page border colour: names and auto, like text', () => {
    for (const color of ['red', ' #FF0000', 'auto']) expect(() => validateOperationStrict({ action: 'pageBorder', color }), color).not.toThrow();
  });

  it('chart series and watermark take any string (a value no form accepts falls back, as before)', () => {
    expect(() => validateOperationStrict({ action: 'watermark', text: 'X', color: 'FF00ZZ' })).not.toThrow();
    expect(() => validateOperationStrict({ action: 'watermark', text: 'X', color: 'red' })).not.toThrow();
  });
});

describe('the schema pattern agrees with the helper', () => {
  const samples = [
    'FF0000', '#ff0000', ' FF0000 ', ' #FF0000', '##FF0000', 'F00', 'FF000', 'FF00000', 'FF00ZZ', '#GGGGGG', '',
    'auto', 'AUTO', 'Auto', ' auto', 'red', 'Red', 'RED', ' red ', 'green', 'blue', 'black', 'white', 'yellow',
    'orange', 'gray', 'grey', 'purple', 'banana', 'pink', 'reds', 're d', '0xFF0000', 'rgb(1,2,3)',
  ];
  for (const allowAuto of [true, false]) {
    it(`allowAuto ${allowAuto}: the same verdict on ${samples.length} strings`, () => {
      const pattern = new RegExp(colourPattern({ allowAuto }));
      for (const sample of samples) {
        expect(pattern.test(sample), `${JSON.stringify(sample)} (allowAuto ${allowAuto})`).toBe(normalizeColour(sample, { allowAuto }) !== undefined);
      }
    });
  }

  it('the generated MCP schema carries the pattern and a description on a colour field', () => {
    const schema = fs.readFileSync(path.resolve(__dirname, '../../schemas/mcp/execute_batch.json'), 'utf8');
    expect(schema).toContain(colourPattern({ allowAuto: true }).replace(/\\/g, '\\\\'));
    expect(schema).toContain('A colour: 6 hex digits such as FF0000');
  });
});

describe('the value that was checked is the value that was written', () => {
  let dir: string;
  let prevRoots: string | undefined;
  beforeEach(() => {
    dir = fs.realpathSync.native(fs.mkdtempSync(path.join(os.tmpdir(), 'astra-colour-')));
    prevRoots = process.env.ASTRA_ALLOWED_DIRS;
    process.env.ASTRA_ALLOWED_DIRS = dir;
    vi.spyOn(console, 'log').mockImplementation(() => undefined);
    vi.spyOn(console, 'error').mockImplementation(() => undefined);
  });
  afterEach(() => {
    vi.restoreAllMocks();
    if (prevRoots === undefined) delete process.env.ASTRA_ALLOWED_DIRS;
    else process.env.ASTRA_ALLOWED_DIRS = prevRoots;
    fs.rmSync(dir, { recursive: true, force: true });
  });

  async function composed(): Promise<{ file: string; paragraph: string; table: string }> {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(createBlankDocxBuffer('blank'), 'colour.docx');
    const out = await handleComposeDocument(store, {
      sessionId: handle.sessionId,
      markdown: 'First paragraph.\n\n| A | B |\n| - | - |\n| 1 | 2 |\n',
    });
    if (!out.ok) throw new Error(JSON.stringify(out));
    const file = path.join(dir, 'base.docx');
    await handleSaveDocument(store, { sessionId: handle.sessionId, outPath: file });
    const ast = await new DocxParser(fs.readFileSync(file)).getAST();
    const nodes = ast.sections.flatMap((s: { nodes: Array<{ id: string; type: string; text?: string }> }) => s.nodes);
    // Ids from compose do not survive a save and reopen; read the ids the saved file gets.
    const tableAt = nodes.findIndex((n: { type: string }) => n.type === 'table');
    return { file, paragraph: nodes[tableAt - 1].id, table: nodes[tableAt].id };
  }

  const documentXml = (file: string): string => new AdmZip(fs.readFileSync(file)).readAsText('word/document.xml');
  const runCliBatch = async (file: string, ops: unknown[]): Promise<number> => {
    const opsFile = path.join(dir, 'ops.json');
    fs.writeFileSync(opsFile, JSON.stringify({ operations: ops }));
    return runCli(['batch', file, opsFile, '--no-history']);
  };
  const viaMcp = async (file: string, ops: unknown[]): Promise<{ ok: boolean; xml: string }> => {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(fs.readFileSync(file), 'in.docx');
    const res = (await handleExecuteBatch(store, { sessionId: handle.sessionId, operations: ops })) as { ok: boolean };
    const saved = path.join(dir, `mcp-${Math.random().toString(16).slice(2)}.docx`);
    await handleSaveDocument(store, { sessionId: handle.sessionId, outPath: saved });
    return { ok: res.ok, xml: documentXml(saved) };
  };

  for (const [input, written] of [[' #FF0000', 'FF0000'], ['red', 'FF0000'], ['Green', '008000'], ['auto', 'auto'], ['#abcdef', 'abcdef']] as const) {
    it(`text colour ${JSON.stringify(input)} is written as ${written}, through the CLI and through MCP`, async () => {
      const base = await composed();
      const file = path.join(dir, 'cli.docx');
      fs.copyFileSync(base.file, file);
      expect(await runCliBatch(file, [{ action: 'style', target: base.paragraph, color: input }])).toBe(EXIT_OK);
      expect(documentXml(file)).toContain(`<w:color w:val="${written}"/>`);
      expect((await validateDocxBuffer(fs.readFileSync(file))).ok).toBe(true);

      const mcp = await viaMcp(base.file, [{ action: 'style', target: base.paragraph, color: input }]);
      expect(mcp.ok).toBe(true);
      expect(mcp.xml).toContain(`<w:color w:val="${written}"/>`);
    });
  }

  it('a colour no form accepts is refused through both entry points and the document bytes do not change', async () => {
    const base = await composed();
    const before = fs.readFileSync(base.file);
    expect(await runCliBatch(base.file, [{ action: 'style', target: base.paragraph, color: 'banana' }])).toBe(EXIT_FAIL);
    expect(fs.readFileSync(base.file).equals(before)).toBe(true);
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(before, 'in.docx');
    const res = (await handleExecuteBatch(store, { sessionId: handle.sessionId, operations: [{ action: 'style', target: base.paragraph, color: 'banana' }] })) as { ok: boolean; message?: string };
    expect(res.ok).toBe(false);
    expect(JSON.stringify(res)).toContain('operations');
  });

  it('highlight: all 17 values are valid in the saved file', async () => {
    const base = await composed();
    for (const highlight of HIGHLIGHT_COLOURS) {
      const file = path.join(dir, `h-${highlight}.docx`);
      fs.copyFileSync(base.file, file);
      expect(await runCliBatch(file, [{ action: 'style', target: base.paragraph, highlight }]), highlight).toBe(EXIT_OK);
      expect((await validateDocxBuffer(fs.readFileSync(file))).ok, highlight).toBe(true);
    }
  });

  it('page border: names and auto reach all four sides; the width bounds validate', async () => {
    const base = await composed();
    for (const [input, written] of [[' #FF0000', 'FF0000'], ['red', 'FF0000'], ['auto', 'auto']] as const) {
      const file = path.join(dir, `pb-${written}.docx`);
      fs.copyFileSync(base.file, file);
      expect(await runCliBatch(file, [{ action: 'pageBorder', color: input }])).toBe(EXIT_OK);
      const borders = /<w:pgBorders\b[\s\S]*?<\/w:pgBorders>/.exec(documentXml(file))?.[0] ?? '';
      const sides = borders.match(/<w:(top|left|bottom|right)\b[^>]*w:color="([^"]+)"/g) ?? [];
      expect(sides.filter((s) => s.includes(`w:color="${written}"`)).length).toBe(4);
      expect((await validateDocxBuffer(fs.readFileSync(file))).ok).toBe(true);
    }
    for (const sz of [1, 2, 96]) {
      const file = path.join(dir, `pb-sz-${sz}.docx`);
      fs.copyFileSync(base.file, file);
      expect(await runCliBatch(file, [{ action: 'pageBorder', sz }]), String(sz)).toBe(EXIT_OK);
      expect((await validateDocxBuffer(fs.readFileSync(file))).ok, String(sz)).toBe(true);
    }
  });

  it('table colours: red is mapped for fill, text colour and border; banana still fails at run time', async () => {
    const base = await composed();
    const file = path.join(dir, 'tbl.docx');
    fs.copyFileSync(base.file, file);
    expect(await runCliBatch(file, [
      { action: 'table', target: base.table, operation: 'setCellShading', rowIndex: 1, colIndex: 1, fill: 'red' },
      { action: 'table', target: base.table, operation: 'setCellTextColor', rowIndex: 1, colIndex: 2, color: 'blue' },
      { action: 'table', target: base.table, operation: 'setBorder', color: 'green', sz: 8 },
    ])).toBe(EXIT_OK);
    const xml = documentXml(file);
    expect(xml).toContain('w:fill="FF0000"');
    expect(xml).toContain('w:val="0000FF"');
    expect(xml).toContain('w:color="008000"');
    expect((await validateDocxBuffer(fs.readFileSync(file))).ok).toBe(true);
    const bad = path.join(dir, 'tbl-bad.docx');
    fs.copyFileSync(base.file, bad);
    const before = fs.readFileSync(bad);
    expect(await runCliBatch(bad, [{ action: 'table', target: base.table, operation: 'setBorder', color: 'banana', sz: 8 }])).toBe(EXIT_FAIL);
    expect(fs.readFileSync(bad).equals(before)).toBe(true);
  });

  it('watermark and chart series: a name is mapped, a value no form accepts keeps today\'s outcome', async () => {
    const base = await composed();
    const wm = path.join(dir, 'wm.docx');
    fs.copyFileSync(base.file, wm);
    expect(await runCliBatch(wm, [{ action: 'watermark', text: 'DRAFT', color: 'red' }])).toBe(EXIT_OK);
    const headerXml = new AdmZip(fs.readFileSync(wm)).getEntries().filter((e) => /word\/header\d*\.xml$/.test(e.entryName)).map((e) => e.getData().toString('utf8')).join('');
    expect(headerXml).toContain('FF0000');
    const wm2 = path.join(dir, 'wm2.docx');
    fs.copyFileSync(base.file, wm2);
    expect(await runCliBatch(wm2, [{ action: 'watermark', text: 'DRAFT', color: 'FF00ZZ' }])).toBe(EXIT_OK);
    const header2 = new AdmZip(fs.readFileSync(wm2)).getEntries().filter((e) => /word\/header\d*\.xml$/.test(e.entryName)).map((e) => e.getData().toString('utf8')).join('');
    expect(header2).toContain('D1D5DB');

    const ch = path.join(dir, 'ch.docx');
    fs.copyFileSync(base.file, ch);
    expect(await runCliBatch(ch, [{ action: 'chart', operation: 'insert', chartType: 'bar', anchor: base.paragraph, categories: ['a', 'b'], series: [{ name: 's', values: [1, 2], color: 'red' }] }])).toBe(EXIT_OK);
    const chart = new AdmZip(fs.readFileSync(ch)).getEntries().filter((e) => /word\/charts\/chart\d*\.xml$/.test(e.entryName)).map((e) => e.getData().toString('utf8')).join('');
    expect(chart).toContain('<a:srgbClr val="FF0000"');
    const ch2 = path.join(dir, 'ch2.docx');
    fs.copyFileSync(base.file, ch2);
    expect(await runCliBatch(ch2, [{ action: 'chart', operation: 'insert', chartType: 'bar', anchor: base.paragraph, categories: ['a', 'b'], series: [{ name: 's', values: [1, 2], color: 'FF00ZZ' }] }])).toBe(EXIT_FAIL);
  });
});

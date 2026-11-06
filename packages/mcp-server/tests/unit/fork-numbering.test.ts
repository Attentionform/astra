import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import * as engine from '@attentionform/astra-engine';
import { AdmZip, DocxParser, NumberingManipulator } from '@attentionform/astra-engine';
import { run as runCli } from '../../../cli/src/cli';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { readParserNormalizations } from '../../src/lib/mcp/normalize-read';
import { adoptForkNumbering } from '../../src/lib/fork-numbering';
import { handleValidateDocument, handleNormalizeDocument } from '../../src/lib/mcp/handlers/parity';
import { handleExportDocument } from '../../src/tools/export';
import { handleVisualInspect } from '../../src/tools/visual';

// Generated from an ASTra bullet list with Word's mc:Ignorable="w15" root
// and w15:restartNumberingAfterBreak on its abstract numbering definition.
const fixture = fs.readFileSync(path.resolve(__dirname, '../../../../tests/fixtures/numbering-w15.docx'));
const numbering = (buf: Buffer) => new AdmZip(buf).readAsText('word/numbering.xml');
let dir: string;
let store: ReturnType<typeof createMcpSessionStore>;
let sessionId: string;

beforeEach(async () => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-fork-numbering-'));
  vi.stubEnv('ASTRA_ALLOWED_DIRS', dir);
  store = createMcpSessionStore();
  sessionId = (await store.openBuffer(fixture, 'numbering-w15.docx')).sessionId;
});

afterEach(async () => {
  await store.closeDocument(sessionId);
  vi.restoreAllMocks();
  vi.unstubAllEnvs();
  fs.rmSync(dir, { recursive: true, force: true });
});

function liveState() {
  const handle = store.get(sessionId);
  return {
    revision: handle.revision,
    dirty: store.isDirty(handle),
    source: JSON.stringify(handle.ast.numbering.source),
    abstract: Object.entries(handle.ast.numbering.abstractNum).map(([id, xml]) => [id, xml.toString()]),
    concrete: Object.entries(handle.ast.numbering.concreteNum).map(([id, xml]) => [id, xml.toString()]),
    normalizations: readParserNormalizations(handle.parser),
  };
}

describe('MCP numbering forks', () => {
  it.each([false, true])('matches CLI validation with w15 metadata (unrelated invalid markup: %s)', async (invalid) => {
    const zip = new AdmZip(fixture);
    if (invalid) {
      zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace(
        '</w:body>', '<w:p><w:r><w:bogus/></w:r></w:p></w:body>',
      )));
      await store.closeDocument(sessionId);
      sessionId = (await store.openBuffer(zip.toBuffer(), 'invalid-w15.docx')).sessionId;
    }
    const file = path.join(dir, 'input.docx');
    fs.writeFileSync(file, zip.toBuffer());
    const log = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    const code = await runCli(['validate', file, '--json']);
    const cli = JSON.parse(log.mock.calls.map(([text]) => String(text)).join(''));
    expect(code).toBe(invalid ? 1 : 0);
    const before = liveState();
    const result = await handleValidateDocument(store, { sessionId });
    expect(result.ok).toBe(true);
    if (!result.ok) throw new Error(JSON.stringify(result));
    expect(result.validation.valid).toBe(!invalid);
    expect(result.validation.errorCount).toBe(cli.errorCount ?? 0);
    expect(result.validation.errors.some(e => e.description.includes('restartNumberingAfterBreak'))).toBe(false);
    expect(liveState()).toEqual(before);
  });

  it('normalizes a valid Word-style list without introducing errors or changing its numbering bytes', async () => {
    const result = await handleNormalizeDocument(store, { sessionId });
    expect(result.ok).toBe(true);
    if (!result.ok) throw new Error(JSON.stringify(result));
    expect(result.normalized).toEqual([]);
    expect(result.remaining).toBe(0);
    expect(result.validation.valid).toBe(true);
    expect(numbering(store.get(sessionId).parser.rezip())).toBe(numbering(fixture));
  });

  it.each([
    ['pdf', false], ['visual', false], ['pdf', true], ['visual', true],
  ] as const)('%s preserves numbering and leaves the live session untouched (new list: %s)', async (tool, changed) => {
    const newList = changed ? new NumberingManipulator(store.get(sessionId).ast).getOrCreateList('numbered') : undefined;
    let rendered: Buffer | undefined;
    if (tool === 'pdf') {
      vi.spyOn(engine, 'hasPdfBackend').mockReturnValue(true);
      vi.spyOn(engine, 'convertDocxToPdf').mockImplementation(async (bytes, out) => {
        rendered = bytes;
        fs.writeFileSync(out, '%PDF-1.4\n');
        return out;
      });
    } else {
      vi.spyOn(engine, 'renderDocxToImages').mockImplementation(async bytes => {
        rendered = bytes;
        return { ok: true, images: [], totalPages: 1, outDir: dir, format: 'png', backend: 'libreoffice' };
      });
    }
    const before = liveState();
    const result = tool === 'pdf'
      ? await handleExportDocument(store, { sessionId, format: 'pdf', outPath: path.join(dir, 'out.pdf') })
      : await handleVisualInspect(store, { sessionId, withVisual: true, pages: [1] });
    expect(result.ok, JSON.stringify(result)).toBe(true);
    expect(rendered).toBeDefined();
    const xml = numbering(rendered!);
    if (changed) {
      expect(xml).not.toBe(numbering(fixture));
      expect(xml).toContain(`w:numId="${newList}"`);
      expect(xml).toContain('mc:Ignorable="w15"');
    } else {
      expect(xml).toBe(numbering(fixture));
    }
    expect((await new DocxParser(rendered!).validateOOXML()).ok).toBe(true);
    expect(liveState()).toEqual(before);
  });

  it('validates newly created list definitions with their root metadata instead of reusing stale bytes', async () => {
    const handle = store.get(sessionId);
    const original = numbering(fixture);
    const list = new NumberingManipulator(handle.ast).getOrCreateList('numbered');
    expect(Object.keys(handle.ast.numbering.concreteNum).length).toBeGreaterThan(1);
    const before = liveState();
    const adopted = vi.spyOn(DocxParser.prototype, 'adoptAstForNumbering');
    const result = await handleValidateDocument(store, { sessionId });
    expect(result.ok).toBe(true);
    if (!result.ok) throw new Error(JSON.stringify(result));
    expect(result.validation.valid).toBe(true);
    const ast = adopted.mock.calls[0][0];
    const fork = new DocxParser(fixture);
    fork.adoptAstForNumbering(ast);
    const xml = numbering(fork.rezip());
    expect(xml).not.toBe(original);
    expect(xml).toContain('mc:Ignorable="w15"');
    expect(xml).toContain(`w:numId="${list}"`);
    expect(liveState()).toEqual(before);
  });

  it('isolates XML definitions, relationship records and every mutable source field', () => {
    const live = store.get(sessionId).ast;
    const before = liveState();
    const fork = new DocxParser(fixture);
    const adopted = vi.spyOn(fork, 'adoptAstForNumbering');
    adoptForkNumbering(fork, live);
    const copy = adopted.mock.calls[0][0];
    expect(Object.keys(copy).sort()).toEqual(['numbering', 'relationships']);
    expect(copy.numbering.source).toEqual(live.numbering.source);
    expect(copy.numbering.source).not.toBe(live.numbering.source);
    expect(copy.numbering.abstractNum['0']).not.toBe(live.numbering.abstractNum['0']);
    expect(copy.numbering.concreteNum['1']).not.toBe(live.numbering.concreteNum['1']);
    copy.numbering.abstractNum['0'].setAttribute('test', 'changed');
    copy.numbering.concreteNum['1'].setAttribute('test', 'changed');
    copy.numbering.source!.rootAttributes[0].value = 'changed';
    copy.numbering.source!.leading.push('changed');
    copy.numbering.source!.trailing.push('changed');
    const rel = Object.keys(live.relationships)[0];
    expect(copy.relationships[rel]).not.toBe(live.relationships[rel]);
    const originalRel = { ...live.relationships[rel] };
    copy.relationships[rel].target = 'changed';
    expect(live.relationships[rel]).toEqual(originalRel);
    expect(liveState()).toEqual(before);
  });

  it('supports an AST without numbering source metadata', async () => {
    const blank = DocxParser.createNewDocument('blank');
    const ast = await blank.getAST();
    delete ast.numbering.source;
    const fork = new DocxParser(blank.rezip());
    adoptForkNumbering(fork, ast);
    expect((await new DocxParser(fork.rezip()).validateOOXML()).ok).toBe(true);
  });
});

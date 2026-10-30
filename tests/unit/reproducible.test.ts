import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import crypto from 'crypto';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import {
  DocxParser,
  createBlankDocxBuffer,
  resolveDeterministicZipDate,
  toDosTimeUtc,
  clampCorePropertiesDates,
  generateEmbeddedWorkbook,
  ZIP_FIXED_EPOCH_MS,
  isReproducibleEnv,
  resolveReproducibleMode,
  resolveCoverDateText,
  setParserClockForTests,
  setParserRandomForTests,
  resetParserClockForTests,
  resetCommentParaSeqForTests,
} from '@attentionform/astra-engine';
import { run as runCli, EXIT_OK } from '../../packages/cli/src/cli';

function sha256(buf: Buffer): string {
  return crypto.createHash('sha256').update(buf).digest('hex');
}

function entryNames(buf: Buffer): string[] {
  return new AdmZip(buf).getEntries().map((e) => e.entryName);
}

function coreXmlOf(buf: Buffer): string | null {
  return new AdmZip(buf).getEntry('docProps/core.xml')?.getData().toString('utf8') ?? null;
}

const SENTINEL_CORE =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
  `xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ` +
  `xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
  `<dc:creator>Sentinel</dc:creator>` +
  `<dcterms:created xsi:type="dcterms:W3CDTF">2001-02-03T04:05:06Z</dcterms:created>` +
  `<dcterms:modified xsi:type="dcterms:W3CDTF">2025-06-07T08:09:10Z</dcterms:modified>` +
  `</cp:coreProperties>`;

function docWithCoreXml(): Buffer {
  const parser = DocxParser.createNewDocument('blank');
  parser.getZipAdapter().writeText('docProps/core.xml', SENTINEL_CORE);
  return parser.rezip({ reproducible: false });
}

describe('hash-stable artifacts', () => {
  const savedEpoch = process.env.SOURCE_DATE_EPOCH;
  const savedFlag = process.env.ASTRA_REPRODUCIBLE;
  const savedNoFlag = process.env.ASTRA_NO_REPRODUCIBLE;

  beforeEach(() => {
    delete process.env.SOURCE_DATE_EPOCH;
    // Pinned output is opt-in. This file is about what the pinned mode guarantees, so it starts in it;
    // tests/unit/track-a-correctness.test.ts covers the real-time default.
    process.env.ASTRA_REPRODUCIBLE = '1';
    delete process.env.ASTRA_NO_REPRODUCIBLE;
    resetParserClockForTests();
    try {
      resetCommentParaSeqForTests();
    } catch { /* best-effort when helper absent */ }
    vi.spyOn(console, 'log').mockImplementation(() => undefined);
    vi.spyOn(console, 'error').mockImplementation(() => undefined);
  });

  afterEach(() => {
    vi.restoreAllMocks();
    resetParserClockForTests();
    try {
      resetCommentParaSeqForTests();
    } catch { /* best-effort */ }
    if (savedEpoch !== undefined) process.env.SOURCE_DATE_EPOCH = savedEpoch;
    else delete process.env.SOURCE_DATE_EPOCH;
    if (savedFlag !== undefined) process.env.ASTRA_REPRODUCIBLE = savedFlag;
    else delete process.env.ASTRA_REPRODUCIBLE;
    if (savedNoFlag !== undefined) process.env.ASTRA_NO_REPRODUCIBLE = savedNoFlag;
    else delete process.env.ASTRA_NO_REPRODUCIBLE;
  });

  it('same input twice yields byte-identical bytes when pinned (ASTRA_REPRODUCIBLE=1)', async () => {
    const base = createBlankDocxBuffer('blank');
    const build = (): Buffer => {
      const p = new DocxParser(base);
      const summary = p.getStructuralSummary();
      const firstId = summary.nodes[0]?.id;
      expect(firstId).toBeTruthy();
      p.updateNode(
        firstId as string,
        '<w:p><w:r><w:t>Stable proof text</w:t></w:r></w:p>'
      );
      return p.rezip();
    };
    const a = build();
    await new Promise((r) => setTimeout(r, 60));
    const b = build();
    expect(a.equals(b)).toBe(true);
    expect(sha256(a)).toBe(sha256(b));
  });

  it('same input twice yields byte-identical bytes on the explicit flagged path', async () => {
    const base = createBlankDocxBuffer('blank');
    const build = (): Buffer => {
      const p = new DocxParser(base);
      const summary = p.getStructuralSummary();
      const firstId = summary.nodes[0]?.id;
      expect(firstId).toBeTruthy();
      p.updateNode(
        firstId as string,
        '<w:p><w:r><w:t>Stable proof text</w:t></w:r></w:p>'
      );
      return p.rezip({ reproducible: true });
    };
    const a = build();
    await new Promise((r) => setTimeout(r, 60));
    const b = build();
    expect(a.equals(b)).toBe(true);
    expect(sha256(a)).toBe(sha256(b));
  });

  it('inserted paragraphs keep identical identities across repeated runs when pinned', async () => {
    const base = createBlankDocxBuffer('blank');
    const build = async (): Promise<Buffer> => {
      const p = new DocxParser(base);
      const ast = await p.getAST();
      const { OperationEngine } = await import('@attentionform/astra-engine');
      const engine = new OperationEngine(ast, p);
      const res = await engine.executeBatch({
        operations: [
          { action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'Seeded insert proof' },
        ] as never,
        sequential: true,
      });
      expect(res.success).toBe(true);
      return p.rezip();
    };
    const a = await build();
    await new Promise((r) => setTimeout(r, 60));
    const b = await build();
    expect(a.equals(b)).toBe(true);
  });

  it('entry order is alphabetical regardless of insertion order', () => {
    const names = entryNames(createBlankDocxBuffer('blank'));
    expect(names.length).toBeGreaterThan(1);
    const sorted = [...names].sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
    expect(names).toEqual(sorted);
  });

  it('all entry timestamps clamp to one stamp', () => {
    process.env.SOURCE_DATE_EPOCH = '1700000000';
    const expected = toDosTimeUtc(resolveDeterministicZipDate());
    const buf = DocxParser.createNewDocument('blank').rezip();
    const entries = new AdmZip(buf).getEntries();
    expect(entries.length).toBeGreaterThan(0);
    for (const e of entries) {
      expect((e.header as { timeval: number }).timeval).toBe(expected);
    }
  });

  it('freezes docProps modified to the epoch when pinned while preserving created; opt-out preserves both', () => {
    const input = docWithCoreXml();
    expect(coreXmlOf(input)).toContain('2001-02-03T04:05:06Z');
    expect(coreXmlOf(input)).toContain('2025-06-07T08:09:10Z');

    const epochIso = new Date(ZIP_FIXED_EPOCH_MS).toISOString();
    const plain = new DocxParser(input).rezip();
    const plainCore = coreXmlOf(plain);
    expect(plainCore).toContain('2001-02-03T04:05:06Z');
    expect(plainCore).not.toContain('2025-06-07T08:09:10Z');
    expect(plainCore).toContain(epochIso);
    expect((plainCore?.match(new RegExp(epochIso.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g')) ?? []).length).toBe(1);

    const flagged = new DocxParser(input).rezip({ reproducible: true });
    const core = coreXmlOf(flagged);
    expect(core).toContain('2001-02-03T04:05:06Z');
    expect(core).not.toContain('2025-06-07T08:09:10Z');
    expect(core).toContain(epochIso);
    expect((core?.match(new RegExp(epochIso.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g')) ?? []).length).toBe(1);

    const optedOut = new DocxParser(input).rezip({ reproducible: false });
    expect(coreXmlOf(optedOut)).toContain('2001-02-03T04:05:06Z');
    expect(coreXmlOf(optedOut)).toContain('2025-06-07T08:09:10Z');

    process.env.ASTRA_NO_REPRODUCIBLE = '1';
    const envOut = new DocxParser(input).rezip();
    expect(coreXmlOf(envOut)).toContain('2001-02-03T04:05:06Z');
    expect(coreXmlOf(envOut)).toContain('2025-06-07T08:09:10Z');
    delete process.env.ASTRA_NO_REPRODUCIBLE;

    process.env.ASTRA_REPRODUCIBLE = '0';
    const envOut2 = new DocxParser(input).rezip();
    expect(coreXmlOf(envOut2)).toContain('2001-02-03T04:05:06Z');
    expect(coreXmlOf(envOut2)).toContain('2025-06-07T08:09:10Z');
  });

  it('honors SOURCE_DATE_EPOCH for the docProps freeze when pinned', () => {
    process.env.SOURCE_DATE_EPOCH = '1700000000';
    const epochIso = resolveDeterministicZipDate().toISOString();
    const flagged = new DocxParser(docWithCoreXml()).rezip({ reproducible: true });
    expect(coreXmlOf(flagged)).toContain(epochIso);
    const plain = new DocxParser(docWithCoreXml()).rezip();
    expect(coreXmlOf(plain)).toContain(epochIso);
  });

  it('deterministic when pinned; explicit and env opt-outs restore wall-clock', () => {
    expect(isReproducibleEnv()).toBe(true);
    expect(resolveReproducibleMode(undefined)).toBe(true);
    expect(resolveReproducibleMode(true)).toBe(true);
    expect(resolveReproducibleMode(false)).toBe(false);
    process.env.ASTRA_NO_REPRODUCIBLE = '1';
    expect(isReproducibleEnv()).toBe(false);
    expect(resolveReproducibleMode(undefined)).toBe(false);
    expect(resolveReproducibleMode(true)).toBe(true);
    delete process.env.ASTRA_NO_REPRODUCIBLE;
    process.env.ASTRA_REPRODUCIBLE = '0';
    expect(isReproducibleEnv()).toBe(false);
    expect(resolveReproducibleMode(undefined)).toBe(false);
    delete process.env.ASTRA_REPRODUCIBLE;
    process.env.ASTRA_REPRODUCIBLE = '1';
    expect(isReproducibleEnv()).toBe(true);
  });

  it('env flag enables the docProps freeze without an explicit option (legacy)', () => {
    process.env.ASTRA_REPRODUCIBLE = '1';
    const epochIso = new Date(ZIP_FIXED_EPOCH_MS).toISOString();
    const flagged = new DocxParser(docWithCoreXml()).rezip();
    expect(coreXmlOf(flagged)).toContain(epochIso);
  });

  it('embedded chart workbooks are byte-stable across builds', () => {
    const series = [{ name: 'Revenue', values: [10, 20] }];
    const a = generateEmbeddedWorkbook(series, ['Q1', 'Q2']);
    const b = generateEmbeddedWorkbook(series, ['Q1', 'Q2']);
    expect(a.equals(b)).toBe(true);
  });

  it('tracked-change default dates are epoch-pinned when pinned and follow SOURCE_DATE_EPOCH', async () => {
    const epochIso = new Date(ZIP_FIXED_EPOCH_MS).toISOString();
    const p = new DocxParser(createBlankDocxBuffer('blank'));
    const summary = p.getStructuralSummary();
    p.insertTrackedChange(summary.nodes[0].id, 'revision proof');
    const xml = new AdmZip(p.rezip()).getEntry('word/document.xml')?.getData().toString('utf8') ?? '';
    expect(xml).toContain(`w:date="${epochIso}"`);

    process.env.SOURCE_DATE_EPOCH = '1700000000';
    const expected = new Date(1700000000 * 1000).toISOString();
    const p2 = new DocxParser(createBlankDocxBuffer('blank'));
    const summary2 = p2.getStructuralSummary();
    p2.insertTrackedChange(summary2.nodes[0].id, 'revision proof');
    const xml2 = new AdmZip(p2.rezip()).getEntry('word/document.xml')?.getData().toString('utf8') ?? '';
    expect(xml2).toContain(`w:date="${expected}"`);
    delete process.env.SOURCE_DATE_EPOCH;

    process.env.ASTRA_NO_REPRODUCIBLE = '1';
    const p3 = new DocxParser(createBlankDocxBuffer('blank'));
    const summary3 = p3.getStructuralSummary();
    const before = Date.now();
    p3.insertTrackedChange(summary3.nodes[0].id, 'revision proof');
    const after = Date.now();
    const xml3 = new AdmZip(p3.rezip()).getEntry('word/document.xml')?.getData().toString('utf8') ?? '';
    const m = xml3.match(/w:date="([^"]+)"/)?.[1];
    expect(m).toBeTruthy();
    expect(m).not.toBe(epochIso);
    expect(Date.parse(m!)).toBeGreaterThanOrEqual(before - 5000);
    expect(Date.parse(m!)).toBeLessThanOrEqual(after + 5000);
  });

  it('clamp helper is a safe no-op without the part', () => {
    const zip = new AdmZip(createBlankDocxBuffer('blank'));
    expect(clampCorePropertiesDates(zip as never)).toBe(false);
  });

  it('cli batch twice with ASTRA_REPRODUCIBLE=1 yields byte-identical files; opt-out differs', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-repro-'));
    try {
      const input = path.join(dir, 'in.docx');
      const outA = path.join(dir, 'a.docx');
      const outB = path.join(dir, 'b.docx');
      const outC = path.join(dir, 'c.docx');
      fs.writeFileSync(input, docWithCoreXml());
      const ops = path.join(dir, 'ops.json');
      fs.writeFileSync(
        ops,
        JSON.stringify({ operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'CLI proof' }], sequential: true })
      );
      await expect(runCli(['batch', input, ops, '--out', outA])).resolves.toBe(EXIT_OK);
      await new Promise((r) => setTimeout(r, 60));
      await expect(runCli(['batch', input, ops, '--out', outB])).resolves.toBe(EXIT_OK);
      const a = fs.readFileSync(outA);
      const b = fs.readFileSync(outB);
      expect(a.equals(b)).toBe(true);
      const epochIso = new Date(ZIP_FIXED_EPOCH_MS).toISOString();
      expect(coreXmlOf(a)).toContain(epochIso);

      await expect(runCli(['batch', input, ops, '--out', outC, '--no-reproducible'])).resolves.toBe(EXIT_OK);
      const c = fs.readFileSync(outC);
      expect(coreXmlOf(c)).toContain('2001-02-03T04:05:06Z');
      expect(c.equals(a)).toBe(false);
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });

  it('cli batch --reproducible twice yields byte-identical files (legacy flag)', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-repro-'));
    try {
      const input = path.join(dir, 'in.docx');
      const outA = path.join(dir, 'a.docx');
      const outB = path.join(dir, 'b.docx');
      fs.writeFileSync(input, docWithCoreXml());
      const ops = path.join(dir, 'ops.json');
      fs.writeFileSync(
        ops,
        JSON.stringify({ operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'CLI proof' }], sequential: true })
      );
      await expect(runCli(['batch', input, ops, '--out', outA, '--reproducible'])).resolves.toBe(EXIT_OK);
      await new Promise((r) => setTimeout(r, 60));
      await expect(runCli(['batch', input, ops, '--out', outB, '--reproducible'])).resolves.toBe(EXIT_OK);
      const a = fs.readFileSync(outA);
      const b = fs.readFileSync(outB);
      expect(a.equals(b)).toBe(true);
      const epochIso = new Date(ZIP_FIXED_EPOCH_MS).toISOString();
      expect(coreXmlOf(a)).toContain(epochIso);
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });

  it('cli batch --no-reproducible preserves docProps dates (opt-out)', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-repro-'));
    try {
      const input = path.join(dir, 'in.docx');
      const out = path.join(dir, 'out.docx');
      fs.writeFileSync(input, docWithCoreXml());
      await expect(runCli(['batch', input, '--append', 'Plain proof', '--out', out, '--no-reproducible'])).resolves.toBe(EXIT_OK);
      expect(coreXmlOf(fs.readFileSync(out))).toContain('2001-02-03T04:05:06Z');
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });

  it('cli batch pins docProps dates when pinned', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-repro-'));
    try {
      const input = path.join(dir, 'in.docx');
      const out = path.join(dir, 'out.docx');
      fs.writeFileSync(input, docWithCoreXml());
      await expect(runCli(['batch', input, '--append', 'Plain proof', '--out', out])).resolves.toBe(EXIT_OK);
      const epochIso = new Date(ZIP_FIXED_EPOCH_MS).toISOString();
      expect(coreXmlOf(fs.readFileSync(out))).toContain(epochIso);
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });

  it('coverPage date is epoch-pinned when pinned and wall-clock on opt-out', async () => {
    const { OperationEngine } = await import('@attentionform/astra-engine');
    const runCover = async (): Promise<string> => {
      const parser = new DocxParser(createBlankDocxBuffer('blank'));
      const ast = await parser.getAST();
      const engine = new OperationEngine(ast, parser);
      const res = await engine.executeBatch({
        operations: [{ action: 'coverPage', title: 'Pinned Cover', subtitle: 'Sub' } as never],
      });
      expect(res.success).toBe(true);
      const buf = parser.rezip();
      return new AdmZip(buf).getEntry('word/document.xml')?.getData().toString('utf8') ?? '';
    };
    const pinnedDate = resolveCoverDateText(undefined);
    expect(pinnedDate).toBe(new Date(ZIP_FIXED_EPOCH_MS).toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' }));
    const a = await runCover();
    await new Promise((r) => setTimeout(r, 10));
    const b = await runCover();
    expect(a).toBe(b);
    expect(a).toContain(pinnedDate);

    process.env.ASTRA_NO_REPRODUCIBLE = '1';
    expect(resolveCoverDateText(undefined)).not.toBe(pinnedDate);
    delete process.env.ASTRA_NO_REPRODUCIBLE;
  });

  it('parser snapshot clock/seed injection is deterministic', () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    setParserClockForTests(() => 1234567890000);
    setParserRandomForTests(() => 0.1234);
    const s1 = parser.createAtomicSnapshot('injected');
    const s2 = parser.createAtomicSnapshot('injected');
    expect(s1.timestamp).toBe(1234567890000);
    expect(s1.id).toBe(s2.id);
    const s3 = parser.createAtomicSnapshot('seeded', { seed: 'abc' });
    const s4 = parser.createAtomicSnapshot('seeded', { seed: 'abc' });
    expect(s3.id).toBe(s4.id);
    resetParserClockForTests();
    const d1 = parser.createAtomicSnapshot('default');
    const d2 = parser.createAtomicSnapshot('default');
    expect(d1.timestamp).toBe(new Date(ZIP_FIXED_EPOCH_MS).getTime());
    expect(d2.timestamp).toBe(new Date(ZIP_FIXED_EPOCH_MS).getTime());
    expect(d1.id.startsWith(`snap_${new Date(ZIP_FIXED_EPOCH_MS).getTime()}_`)).toBe(true);
    expect(d2.id.startsWith(`snap_${new Date(ZIP_FIXED_EPOCH_MS).getTime()}_`)).toBe(true);
    expect(d1.id).not.toBe(d2.id);
    resetParserClockForTests();
    const e1 = parser.createAtomicSnapshot('default');
    expect(e1.id).toBe(d1.id);
  });

  it('insertRootNode routes through injected clock in default and opt-out modes', () => {
    const base = createBlankDocxBuffer('blank');
    setParserClockForTests(() => 1111111111111);
    setParserRandomForTests(() => 0.25);
    const p1 = new DocxParser(base);
    const id1 = p1.insertRootNode('<w:p><w:r><w:t>hi</w:t></w:r></w:p>');
    const p2 = new DocxParser(base);
    const id2 = p2.insertRootNode('<w:p><w:r><w:t>hi</w:t></w:r></w:p>');
    expect(id1).toBe('node_1111111111111_250');
    expect(id2).toBe(id1);
    process.env.ASTRA_NO_REPRODUCIBLE = '1';
    try {
      const p3 = new DocxParser(base);
      const id3 = p3.insertRootNode('<w:p><w:r><w:t>hi</w:t></w:r></w:p>');
      expect(id3).toBe('node_1111111111111_250');
    } finally {
      delete process.env.ASTRA_NO_REPRODUCIBLE;
    }
    resetParserClockForTests();
  });
});

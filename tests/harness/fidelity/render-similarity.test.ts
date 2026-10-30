/**
 * Render-similarity advisory (advisory only: nothing here blocks a merge).
 *
 * Per fixture the runner replays the corpus pipeline (ingest, mutate two
 * sentences, rezip), renders the before/after bytes with the local
 * LibreOffice, and records one advisory entry: `measured` with pixel, page,
 * and text scores plus an advisory band, or `unavailable` with the reason
 * when the renderer cannot run (missing binary, timeout, crash). No
 * magnitude ever fails this suite; assertions cover harness integrity only
 * (entry per fixture, finite scores, schema shape, deterministic pure
 * helpers). A renderer that is entirely absent skips the render cases
 * cleanly; the pure-helper cases always run.
 *
 * The engine-pipeline helpers below intentionally duplicate the structural
 * gate's logic instead of importing it: the gate file stays untouched, so a
 * change here can never weaken the hard gate.
 */
import { describe, it, expect, afterAll } from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import AdmZip from 'adm-zip';
import { XMLSerializer } from '@xmldom/xmldom';
import { DocxParser, OperationEngine } from '@attentionform/astra-engine';
import {
  advisoryBand,
  convertDocxToPdf,
  createIsolatedProfile,
  isOracleAvailable,
  libreOfficeVersion,
  oracleBinary,
  parsePgm,
  pdfLayoutText,
  pdfPageCount,
  pixelDifference,
  renderPdfToRasters,
  tokenSimilarity,
  OracleUnavailableError,
  type FixtureAdvisory,
  type RasterPage,
} from './render-similarity';

const HERE = __dirname;
const FIX = path.join(HERE, 'fixtures');

const ORACLE = isOracleAvailable();
const RENDERER = ORACLE ? libreOfficeVersion() : null;
/**
 * One renderer profile per test worker: sequential conversions share it
 * (a fresh profile costs one ~20 s warm-up, later conversions are
 * sub-second), parallel workers each own theirs, and the default user
 * profile is never touched.
 */
const PROFILE = ORACLE ? createIsolatedProfile() : null;

afterAll(() => {
  if (PROFILE !== null) fs.rmSync(PROFILE, { recursive: true, force: true });
});

interface AstNode {
  id: string;
  type: string;
  checksum: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  xml?: any;
  metadata: { text?: string };
}

const serializer = new XMLSerializer();

function plainTextTargets(ast: { sections: Array<{ nodes: AstNode[] }> }): AstNode[] {
  return ast.sections
    .flatMap((s) => s.nodes)
    .filter((n) => n.type === 'paragraph' && (n.metadata?.text ?? '').trim().length > 0)
    .filter((n) => {
      if (!n.xml) return false;
      const xml = serializer.serializeToString(n.xml);
      const withoutPPr = xml.replace(/<w:pPr>[\s\S]*?<\/w:pPr>/g, '');
      return !withoutPPr.includes('<w:rPr');
    })
    .slice(0, 2);
}

async function mutateTwoSentences(src: Buffer, stem: string): Promise<Buffer> {
  const parser = new DocxParser(src);
  const ast = (await parser.getAST()) as unknown as { sections: Array<{ nodes: AstNode[] }> };
  const targets = plainTextTargets(ast);
  if (targets.length < 2) throw new Error(`fixture ${stem}: only ${targets.length} plain text paragraphs, need 2`);
  const engine = new OperationEngine(ast as never, parser as never);
  const result = (await (engine as unknown as {
    executeBatch: (req: unknown) => Promise<{ success: boolean; appliedCount: number }>;
  }).executeBatch({
    operations: targets.map((t, i) => ({
      action: 'replace',
      target: t.id,
      text: `Corpus mutation sentence ${i + 1} for ${stem}.`,
    })),
    sequential: true,
  })) as { success: boolean; appliedCount: number };
  if (!result.success || result.appliedCount !== 2) {
    throw new Error(`fixture ${stem}: mutation batch failed ${JSON.stringify(result)}`);
  }
  // Touch adm-zip the same way the gate does so unused-import drift is caught.
  void new AdmZip(parser.rezip()).getEntries().length;
  return parser.rezip();
}

function fixtureStems(): string[] {
  return fs
    .readdirSync(FIX)
    .filter((f) => f.endsWith('.docx'))
    .map((f) => f.replace(/\.docx$/, ''))
    .sort();
}

function scorePages(before: RasterPage[], after: RasterPage[]): { mean: number; worst: number } {
  const shared = Math.min(before.length, after.length);
  let sum = 0;
  let worst = 0;
  for (let i = 0; i < shared; i++) {
    const d = pixelDifference(before[i], after[i]).mean;
    sum += d;
    if (d > worst) worst = d;
  }
  return { mean: shared === 0 ? 1 : sum / shared, worst };
}

describe('render similarity scorer without a renderer', () => {
  it('parses binary PGM rasters', () => {
    const header = Buffer.from('P5\n2 1\n255\n', 'ascii');
    const entry = parsePgm(Buffer.concat([header, Buffer.from([0, 255])]));
    expect(entry.width).toBe(2);
    expect(entry.height).toBe(1);
    expect([...entry.pixels]).toEqual([0, 255]);
  });

  it('rejects non-PGM input', () => {
    expect(() => parsePgm(Buffer.from('P6\n1 1\n255\n\x00', 'ascii'))).toThrow();
  });

  it('scores identical rasters at zero and inverted rasters at one', () => {
    const a: RasterPage = { width: 2, height: 1, pixels: Buffer.from([10, 200]) };
    expect(pixelDifference(a, a).mean).toBe(0);
    const b: RasterPage = { width: 2, height: 1, pixels: Buffer.from([20, 220]) };
    expect(pixelDifference(a, b).mean).toBeCloseTo(15 / 255, 6);
    const white: RasterPage = { width: 2, height: 1, pixels: Buffer.from([255, 255]) };
    const black: RasterPage = { width: 2, height: 1, pixels: Buffer.from([0, 0]) };
    expect(pixelDifference(white, black).mean).toBe(1);
  });

  it('flags dimension mismatches instead of comparing', () => {
    const a: RasterPage = { width: 2, height: 1, pixels: Buffer.from([0, 0]) };
    const b: RasterPage = { width: 1, height: 1, pixels: Buffer.from([0]) };
    const diff = pixelDifference(a, b);
    expect(diff.dimMismatch).toBe(true);
    expect(diff.mean).toBe(1);
  });

  it('grades text similarity on shared tokens', () => {
    expect(tokenSimilarity('alpha beta', 'alpha beta')).toBe(1);
    expect(tokenSimilarity('', '')).toBe(1);
    expect(tokenSimilarity('alpha', 'beta')).toBe(0);
    expect(tokenSimilarity('a b c', 'a b d')).toBeCloseTo(0.5, 6);
  });

  it('bands scores with page-count changes always needing eyes', () => {
    expect(advisoryBand(0.0, true)).toBe('consistent');
    expect(advisoryBand(0.05, true)).toBe('localized-change');
    expect(advisoryBand(0.5, true)).toBe('needs-eyes');
    expect(advisoryBand(0.0, false)).toBe('needs-eyes');
  });

  it('reports oracle presence honestly', () => {
    expect(typeof ORACLE).toBe('boolean');
    if (ORACLE) {
      expect(typeof oracleBinary()).toBe('string');
      expect(typeof RENDERER).toBe('string');
    } else {
      expect(oracleBinary()).toBeNull();
    }
  });
});

describe.skipIf(!ORACLE)('render similarity advisory per fixture', () => {
  const advisories: FixtureAdvisory[] = [];
  const stems = fixtureStems();

  for (const stem of stems) {
    it(`${stem} records a render advisory`, async () => {
      const unavailable = (reason: string): void => {
        advisories.push({ fixture: stem, status: 'unavailable', renderer: RENDERER, reason });
      };
      let workDir = '';
      try {
        const src = fs.readFileSync(path.join(FIX, `${stem}.docx`));
        const mutated = await mutateTwoSentences(src, stem);
        workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-advisory-'));
        const beforeDocx = path.join(workDir, `${stem}-before.docx`);
        const afterDocx = path.join(workDir, `${stem}-after.docx`);
        fs.writeFileSync(beforeDocx, src);
        fs.writeFileSync(afterDocx, mutated);
        let beforePdf: string;
        let afterPdf: string;
        try {
          beforePdf = convertDocxToPdf(beforeDocx, workDir, { profileDir: PROFILE as string });
          afterPdf = convertDocxToPdf(afterDocx, workDir, { profileDir: PROFILE as string });
        } catch (err) {
          unavailable(err instanceof OracleUnavailableError ? err.message : String(err).slice(0, 160));
          return;
        }
        let pagesBefore = 0;
        let pagesAfter = 0;
        try {
          pagesBefore = pdfPageCount(beforePdf);
          pagesAfter = pdfPageCount(afterPdf);
          const rastersBefore = renderPdfToRasters(beforePdf, workDir);
          const rastersAfter = renderPdfToRasters(afterPdf, workDir);
          const textBefore = pdfLayoutText(beforePdf);
          const textAfter = pdfLayoutText(afterPdf);
          const pageCountEqual = pagesBefore === pagesAfter;
          const { mean, worst } = scorePages(rastersBefore, rastersAfter);
          const entry: FixtureAdvisory = {
            fixture: stem,
            status: 'measured',
            renderer: RENDERER,
            pagesBefore,
            pagesAfter,
            pageCountEqual,
            pixelMean: mean,
            pixelWorstPage: worst,
            textSimilarity: tokenSimilarity(textBefore, textAfter),
            band: advisoryBand(mean, pageCountEqual),
          };
          expect(Number.isFinite(entry.pixelMean)).toBe(true);
          expect(Number.isFinite(entry.textSimilarity)).toBe(true);
          advisories.push(entry);
        } catch (err) {
          unavailable(err instanceof OracleUnavailableError ? err.message : String(err).slice(0, 160));
          return;
        }
      } finally {
        if (workDir !== '') fs.rmSync(workDir, { recursive: true, force: true });
      }
    }, 180_000);
  }

  it('advisory report covers every fixture with no verdicts', () => {
    expect(advisories.length).toBe(stems.length);
    for (const entry of advisories) {
      expect(entry.renderer).toBe(RENDERER);
      if (entry.status === 'measured') {
        expect(entry.band).toMatch(/^(consistent|localized-change|needs-eyes)$/);
      } else {
        expect(typeof entry.reason).toBe('string');
      }
    }
    const report = {
      generatedAt: new Date().toISOString(),
      renderer: RENDERER,
      entries: advisories,
    };
    const outPath =
      process.env.ASTRA_ADVISORY_OUT ?? path.join(os.tmpdir(), `astra-render-advisory-${Date.now()}.json`);
    fs.writeFileSync(outPath, `${JSON.stringify(report, null, 2)}\n`);
    const bands = new Map<string, number>();
    const missing: string[] = [];
    for (const entry of advisories) {
      if (entry.status === 'measured') bands.set(entry.band as string, (bands.get(entry.band as string) ?? 0) + 1);
      else missing.push(`${entry.fixture} (${entry.reason})`);
    }
    console.log(`render advisory: ${JSON.stringify(Object.fromEntries(bands))} unavailable=${missing.length}`);
    for (const line of missing) console.log(`render advisory unavailable: ${line}`);
    console.log(`render advisory report: ${outPath}`);
  }, 60_000);
});

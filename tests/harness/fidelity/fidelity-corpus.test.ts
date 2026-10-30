/**
 * Adversarial fidelity corpus runner.
 *
 * Per fixture: manifest gate -> ingest -> normalize (engine-owned, logged)
 * -> mutate two sentences -> rezip -> validateOOXML, asserting the HARD gate:
 * schema passes, the semantic diff shows ONLY the intended mutation (node
 * checksums), no tag is stripped or added (document.xml tag histogram), and
 * every non-document package part is preserved (XML parts semantically
 * identical, binary parts compared exactly; numbering parts re-serialize
 * larger with no meaning change: duo 862 to 1140 bytes, restart 840 to 1052
 * bytes). A second full pass from the same source bytes must match exactly
 * (determinism). Per hostile twin: the gate must refuse at the documented stage.
 */
import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import AdmZip from 'adm-zip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { DocxParser, OperationEngine } from '@attentionform/astra-engine';
import schema from './manifest.schema.json';
import { useReproducibleOutput } from '../../helpers/reproducible';

useReproducibleOutput();

const HERE = __dirname;
const FIX = path.join(HERE, 'fixtures');
const HOSTILE = path.join(HERE, 'hostile');
const SUITE_START = Date.now();
const CI_BUDGET_MS = 300_000;

function countDocx(dir: string): number {
  try {
    return fs.readdirSync(dir).filter((f) => f.endsWith('.docx')).length;
  } catch {
    return 0;
  }
}

function missingFixtureSiblings(): string[] {
  const missing: string[] = [];
  try {
    for (const f of fs.readdirSync(FIX)) {
      if (f.endsWith('.manifest.json')) {
        const stem = f.slice(0, -'.manifest.json'.length);
        if (!fs.existsSync(path.join(FIX, `${stem}.docx`))) missing.push(`fixtures/${stem}.docx`);
      }
    }
  } catch {
    // readdir failure means the generator step must recreate dirs.
  }
  try {
    for (const f of fs.readdirSync(HOSTILE)) {
      if (f.endsWith('.twin.json')) {
        const twin = f.slice(0, -'.json'.length);
        if (!fs.existsSync(path.join(HOSTILE, `${twin}.docx`))) missing.push(`hostile/${twin}.docx`);
      }
    }
  } catch {
    // same as above.
  }
  return missing;
}

/**
 * Self-heal for CI-only "No test found in suite": fixture/hostile .docx
 * binaries are gitignored (*.docx, except test-gallery) and absent on fresh
 * checkouts. The generator (generate-fixtures.ts) is runner-of-record. When
 * binaries are present this is a two-readdir no-op. It runs synchronously at
 * import time so test discovery (fixtureStems()/twinDescriptors() loops
 * below) sees rebuilt files. Any generation failure throws naming the
 * generator step — never a silent skip or empty suite. Hostile twins are
 * rebuilt by the same generator invocation.
 */
function ensureCorpusFixtures(): void {
  const beforeFix = countDocx(FIX);
  const beforeHostile = countDocx(HOSTILE);
  if (beforeFix > 0 && beforeHostile > 0 && missingFixtureSiblings().length === 0) return;
  const generator = path.join(HERE, 'generate-fixtures.ts');
  try {
    execFileSync('npx', ['--yes', 'tsx', generator], {
      cwd: path.resolve(HERE, '../../..'),
      stdio: 'inherit',
      timeout: 120_000,
    });
  } catch (err) {
    throw new Error(
      `fidelity fixture self-heal failed: generator step failed (npx tsx ${generator}): ${err instanceof Error ? err.message : String(err)}`,
    );
  }
  const afterFix = countDocx(FIX);
  const afterHostile = countDocx(HOSTILE);
  const stillMissing = missingFixtureSiblings();
  if (afterFix === 0 || afterHostile === 0 || stillMissing.length > 0) {
    throw new Error(
      `fidelity fixture self-heal failed: generator step completed but binaries still missing (fixtures=${afterFix}, hostile=${afterHostile}, missing=${stillMissing.join(', ') || 'none'}); generator: ${generator}`,
    );
  }
}

ensureCorpusFixtures();

const wordCount = (s: string): number => s.trim().split(/\s+/).filter(Boolean).length;

interface ManifestDoc {
  fixture?: unknown;
  family?: unknown;
  motivatingPathology?: unknown;
  knownLimitations?: unknown;
  scopeExclusions?: unknown;
  ecma376Difficulty?: unknown;
  licenseProvenance?: unknown;
  disposition?: unknown;
  sourceDocument?: unknown;
  [key: string]: unknown;
}

function validateManifest(manifest: unknown, stem: string): string[] {
  const errors: string[] = [];
  if (manifest === null || typeof manifest !== 'object' || Array.isArray(manifest)) {
    return ['manifest root must be a JSON object'];
  }
  const doc = manifest as ManifestDoc;
  for (const field of schema.requiredFields) {
    if (!(field in doc)) errors.push(`missing required field '${field}'`);
  }
  const allowed = new Set<string>([...schema.requiredFields, ...schema.optionalFields]);
  for (const key of Object.keys(doc)) {
    if (!allowed.has(key)) {
      if ((schema.bannedAliases as string[]).includes(key)) {
        errors.push(`banned alias '${key}' is not a schema field`);
      } else {
        errors.push(`unknown field '${key}' (no additional properties allowed)`);
      }
    }
  }
  if (typeof doc.fixture !== 'string' || doc.fixture !== stem) {
    errors.push(`fixture must equal file stem '${stem}'`);
  }
  if (typeof doc.fixture === 'string' && !new RegExp(schema.namePattern).test(doc.fixture)) {
    errors.push('fixture must match ' + schema.namePattern);
  }
  if (typeof doc.family !== 'string' || !new RegExp(schema.namePattern).test(doc.family)) {
    errors.push('family must match ' + schema.namePattern);
  }
  const ranges = schema.wordRanges as unknown as Record<string, [number, number]>;
  for (const [field, [lo, hi]] of Object.entries(ranges)) {
    if (field.startsWith('disposition.')) continue;
    const value = doc[field];
    if (typeof value !== 'string') {
      if (field in doc) errors.push(`'${field}' must be a string`);
      continue;
    }
    const n = wordCount(value);
    if (n < lo || n > hi) errors.push(`'${field}' has ${n} words, required ${lo}-${hi}`);
  }
  const disp = doc.disposition as { verdict?: unknown; reason?: unknown } | undefined;
  if (disp === null || typeof disp !== 'object' || Array.isArray(disp)) {
    errors.push('disposition must be an object {verdict, reason}');
  } else {
    const keys = Object.keys(disp);
    if (!keys.includes('verdict') || !keys.includes('reason')) errors.push('disposition requires verdict and reason');
    for (const k of keys) if (k !== 'verdict' && k !== 'reason') errors.push(`disposition has unknown field '${k}'`);
    if (!((schema.dispositionVerdicts as string[]).includes(disp.verdict as string))) {
      errors.push(`disposition.verdict must be one of ${(schema.dispositionVerdicts as string[]).join('|')}`);
    }
    const [lo, hi] = ranges['disposition.reason'];
    if (typeof disp.reason !== 'string') errors.push('disposition.reason must be a string');
    else {
      const n = wordCount(disp.reason);
      if (n < lo || n > hi) errors.push(`disposition.reason has ${n} words, required ${lo}-${hi}`);
    }
  }
  return errors;
}

interface AstNode {
  id: string;
  type: string;
  checksum: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  xml?: any;
  metadata: { text?: string };
}

const serializer = new XMLSerializer();

function textParagraphs(ast: { sections: Array<{ nodes: AstNode[] }> }): AstNode[] {
  return ast.sections
    .flatMap((s) => s.nodes)
    .filter((n) => n.type === 'paragraph' && (n.metadata?.text ?? '').trim().length > 0);
}

/**
 * Mutation targets are the first two paragraphs whose runs carry no direct
 * formatting. Rationale, recorded not hidden: replace carries only the first
 * run's properties onto the rewritten text, so targeting a formatted
 * paragraph (especially one with distinct per-run formatting) would conflate
 * replace merge semantics with the preservation under test. Formatted
 * targets stay excluded by rule; the exclusion is documented in the gallery
 * manifests and the README.
 */
function mutationTargets(ast: { sections: Array<{ nodes: AstNode[] }> }): AstNode[] {
  return textParagraphs(ast)
    .filter((n) => {
      if (!n.xml) return false;
      const xml = serializer.serializeToString(n.xml);
      const withoutPPr = xml.replace(/<w:pPr>[\s\S]*?<\/w:pPr>/g, '');
      return !withoutPPr.includes('<w:rPr');
    })
    .slice(0, 2);
}

/**
 * Semantic XML equality for engine-managed package parts (numbering, styles,
 * footers, content types, rels). The engine re-serializes parts it touches,
 * which moves namespace declarations and inter-tag whitespace without changing
 * meaning; byte comparison would flag that canonicalization as corruption.
 * Compares tag names, non-xmlns attributes, and real text exactly; ignores
 * xmlns declaration placement and whitespace-only text nodes.
 */
function significantChildren(node: Node): Node[] {
  const out: Node[] = [];
  for (let i = 0; i < node.childNodes.length; i++) {
    const child = node.childNodes[i];
    if (child.nodeType === 8) continue;
    if (child.nodeType === 3) {
      if ((child.nodeValue ?? '').trim() === '') continue;
      out.push(child);
      continue;
    }
    out.push(child);
  }
  return out;
}

function domEqual(a: Node, b: Node): string | null {
  if (a.nodeType !== b.nodeType) return `nodeType ${a.nodeType} vs ${b.nodeType}`;
  if (a.nodeType === 3) {
    return (a.nodeValue ?? '') === (b.nodeValue ?? '') ? null : `text ${JSON.stringify(a.nodeValue)} vs ${JSON.stringify(b.nodeValue)}`;
  }
  if (a.nodeType !== 1) return null;
  const ea = a as Element;
  const eb = b as Element;
  if (ea.tagName !== eb.tagName) return `tag ${ea.tagName} vs ${eb.tagName}`;
  const attrs = (e: Element): Map<string, string> => {
    const m = new Map<string, string>();
    for (let i = 0; i < e.attributes.length; i++) {
      const attr = e.attributes.item(i);
      if (attr === null) continue;
      if (attr.name === 'xmlns' || attr.name.startsWith('xmlns:')) continue;
      m.set(attr.name, attr.value);
    }
    return m;
  };
  const aa = attrs(ea);
  const ab = attrs(eb);
  if (aa.size !== ab.size) return `attr count on <${ea.tagName}>: ${aa.size} vs ${ab.size}`;
  for (const [k, v] of aa) {
    if (!ab.has(k)) return `attr ${k} missing on <${eb.tagName}>`;
    if (ab.get(k) !== v) return `attr ${k}=${JSON.stringify(v)} vs ${JSON.stringify(ab.get(k))} on <${ea.tagName}>`;
  }
  const ca = significantChildren(ea);
  const cb = significantChildren(eb);
  if (ca.length !== cb.length) return `child count of <${ea.tagName}>: ${ca.length} vs ${cb.length}`;
  for (let i = 0; i < ca.length; i++) {
    const diff = domEqual(ca[i], cb[i]);
    if (diff !== null) return `<${ea.tagName}>[${i}]: ${diff}`;
  }
  return null;
}

function xmlPartsEqual(before: Buffer, after: Buffer, partName: string): string | null {
  let da: Document;
  let db: Document;
  try {
    da = new DOMParser().parseFromString(before.toString('utf8'), 'text/xml');
    db = new DOMParser().parseFromString(after.toString('utf8'), 'text/xml');
  } catch (err) {
    return `${partName} unparseable after round trip: ${String(err).slice(0, 120)}`;
  }
  return domEqual(da.documentElement, db.documentElement);
}

function tagHistogram(documentXml: string): Record<string, number> {
  const doc = new DOMParser().parseFromString(documentXml, 'text/xml');
  const all = doc.getElementsByTagName('*');
  const counts: Record<string, number> = {};
  for (let i = 0; i < all.length; i++) {
    const tag = all.item(i)?.tagName ?? '?';
    counts[tag] = (counts[tag] ?? 0) + 1;
  }
  return counts;
}

function zipEntries(buf: Buffer): AdmZip {
  return new AdmZip(buf);
}

function entryMap(buf: Buffer): Map<string, Buffer> {
  const out = new Map<string, Buffer>();
  for (const entry of zipEntries(buf).getEntries()) {
    if (!entry.entryName.endsWith('/')) out.set(entry.entryName, entry.getData());
  }
  return out;
}

interface PipelineResult {
  mutated: Buffer;
  targetIds: string[];
  normalizations: string[];
}

async function runPipeline(src: Buffer, stem: string): Promise<PipelineResult> {
  const parser = new DocxParser(src);
  const ast = (await parser.getAST()) as unknown as { sections: Array<{ nodes: AstNode[] }> };
  const targets = mutationTargets(ast);
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
  const mutated = parser.rezip();
  return {
    mutated,
    targetIds: targets.map((t) => t.id),
    normalizations: (parser as unknown as { lastNormalizations: string[] }).lastNormalizations ?? [],
  };
}

async function snapshotOf(buf: Buffer): Promise<{ nodes: Map<string, { checksum: string; text: string }>; histogram: Record<string, number>; parts: Map<string, Buffer> }> {
  const parser = new DocxParser(buf);
  const ast = (await parser.getAST()) as unknown as { sections: Array<{ nodes: AstNode[] }> };
  const nodes = new Map<string, { checksum: string; text: string }>();
  for (const section of ast.sections) {
    for (const node of section.nodes) {
      nodes.set(node.id, { checksum: node.checksum, text: node.metadata?.text ?? '' });
    }
  }
  const parts = entryMap(buf);
  const documentXml = parts.get('word/document.xml')?.toString('utf8') ?? '';
  return { nodes, histogram: tagHistogram(documentXml), parts };
}

function fixtureStems(): string[] {
  return fs
    .readdirSync(FIX)
    .filter((f) => f.endsWith('.docx'))
    .map((f) => f.replace(/\.docx$/, ''))
    .sort();
}

function twinDescriptors(): Array<{ family: string; twin: string; base: string; corruption: string; expectedRefusal: string }> {
  return fs
    .readdirSync(HOSTILE)
    .filter((f) => f.endsWith('.twin.json'))
    .sort()
    .map((f) => JSON.parse(fs.readFileSync(path.join(HOSTILE, f), 'utf8')));
}

describe('fidelity corpus: fixtures present', () => {
  it('self-heal guard discovers fixtures and hostile twins', () => {
    const stems = fixtureStems();
    expect(
      stems.length,
      'fidelity fixture self-heal failed: zero .docx in fixtures/ after generator step (generate-fixtures.ts)',
    ).toBeGreaterThan(0);
    const twinDocx = countDocx(HOSTILE);
    expect(
      twinDocx,
      'fidelity fixture self-heal failed: zero .docx in hostile/ after generator step (generate-fixtures.ts)',
    ).toBeGreaterThan(0);
  });
});

describe('fidelity corpus: manifest gate', () => {
  for (const stem of fixtureStems()) {
    it(`manifest for ${stem} satisfies the schema`, () => {
      const manifestPath = path.join(FIX, `${stem}.manifest.json`);
      expect(fs.existsSync(manifestPath), `missing ${stem}.manifest.json`).toBe(true);
      const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
      expect(validateManifest(manifest, stem)).toEqual([]);
    });
  }
});

describe('fidelity corpus: structural gate', () => {
  for (const stem of fixtureStems()) {
    it(`${stem} round-trips with only the intended mutation`, async () => {
      const src = fs.readFileSync(path.join(FIX, `${stem}.docx`));
      const before = await snapshotOf(src);

      const first = await runPipeline(src, stem);
      const second = await runPipeline(src, stem);
      expect(second.mutated.equals(first.mutated), 'pipeline must be byte-deterministic').toBe(true);

      const parser = new DocxParser(first.mutated);
      const validation = (await parser.validateOOXML()) as unknown as { ok: boolean; errors: unknown[] };
      expect(validation.ok, `validateOOXML failed: ${JSON.stringify(validation.errors).slice(0, 500)}`).toBe(true);

      const after = await snapshotOf(first.mutated);

      expect([...after.nodes.keys()].sort()).toEqual([...before.nodes.keys()].sort());
      for (const [id, prev] of before.nodes) {
        const next = after.nodes.get(id);
        if (first.targetIds.includes(id)) {
          expect(next?.checksum, `target ${id} must change`).not.toBe(prev.checksum);
          expect(next?.text, `target ${id} must carry the mutation`).toMatch(
            new RegExp(`^Corpus mutation sentence [12] for ${stem.replace(/-/g, '\\-')}\\.$`),
          );
        } else {
          expect(next?.checksum, `untouched node ${id} must be byte-identical`).toBe(prev.checksum);
        }
      }

      expect(after.histogram).toEqual(before.histogram);

      expect([...after.parts.keys()].sort()).toEqual([...before.parts.keys()].sort());
      for (const [name, data] of before.parts) {
        if (name === 'word/document.xml') continue;
        const next = after.parts.get(name);
        expect(next, `part ${name} must survive the round trip`).toBeDefined();
        if (name.endsWith('.xml') || name.endsWith('.rels')) {
          expect(xmlPartsEqual(data, next as Buffer, name), `part ${name} must be semantically identical`).toBeNull();
        } else {
          expect((next as Buffer).equals(data), `part ${name} must be byte-identical`).toBe(true);
        }
      }
    }, 60_000);
  }
});

describe('fidelity corpus: hostile twins are refused', () => {
  for (const twin of twinDescriptors()) {
    it(`twin ${twin.twin} is refused (${twin.expectedRefusal})`, async () => {
      const src = fs.readFileSync(path.join(HOSTILE, `${twin.twin}.docx`));
      let stage = 'pass';
      try {
        const parser = new DocxParser(src);
        try {
          await parser.getAST();
        } catch {
          stage = 'ingest';
        }
        if (stage === 'pass') {
          try {
            parser.rezip();
          } catch {
            stage = 'rezip';
          }
        }
        if (stage === 'pass') {
          const validation = (await parser.validateOOXML()) as unknown as { ok: boolean };
          if (!validation.ok) stage = 'validate';
        }
      } catch {
        stage = 'open';
      }
      expect(stage, `hostile twin ${twin.twin} passed the gate; corruption: ${twin.corruption}`).not.toBe('pass');
      expect(stage, `twin refused at '${stage}', documented '${twin.expectedRefusal}'`).toBe(twin.expectedRefusal);
    }, 60_000);
  }
});

describe('fidelity corpus: budget', () => {
  it('completes in CI-suitable time', () => {
    const elapsed = Date.now() - SUITE_START;
    console.log(`fidelity corpus wall time: ${(elapsed / 1000).toFixed(1)}s`);
    expect(elapsed).toBeLessThan(CI_BUDGET_MS);
  });
});

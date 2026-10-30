/**
 * What opening a document looks like today, frozen, so that changing how the
 * parser groups blocks into sections cannot change anything else.
 *
 * Every document is read and run through the outputs people see: the node list
 * (ids, types, text, order), the inspect outline, Markdown, HTML, statistics,
 * the text wireframe, and the bytes of a save with no edits. Each output is
 * reduced to a hash and compared with `sections-characterization.golden.json`.
 *
 * - Documents with ONE section must not change in any output.
 * - Documents with SEVERAL sections must keep the same node list, Markdown,
 *   HTML, statistics and saved bytes. The inspect outline and wireframe group
 *   blocks by section, so they are recorded as information only and are
 *   expected to change when the parser starts reading the file's sections.
 *
 * The golden file was first written from the code before the parser read the
 * file's sections. One frozen value has changed since, on purpose: for a file
 * whose only sectPr besides the body's sits inside a table cell, the page setup
 * used to come from that stray element and now comes from the body's. To rewrite
 * it after an intended change: ASTRA_UPDATE_GOLDEN=1 npx vitest run tests/unit/sections-characterization.test.ts
 */
import { describe, it, expect, afterAll } from 'vitest';
import * as crypto from 'crypto';
import * as fs from 'fs';
import * as path from 'path';
import {
  DocxParser,
  ToonGenerator,
  computeDocStats,
  convertAstToHtml,
  convertAstToMarkdown,
  generateVisualWireframe,
  createBlankDocxBuffer,
} from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../helpers/reproducible';
import { SECTION_SHAPES } from '../helpers/section-docs';

useReproducibleOutput();

const ROOT = path.resolve(__dirname, '../..');
const GOLDEN_PATH = path.join(ROOT, 'tests/fixtures/sections/characterization.golden.json');
const UPDATE = process.env.ASTRA_UPDATE_GOLDEN === '1';

const sha = (data: string | Buffer): string => crypto.createHash('sha256').update(data).digest('hex').slice(0, 16);

type Kind = 'single' | 'multi';
interface Source {
  name: string;
  kind: Kind;
  load: () => Buffer;
}

const read = (rel: string) => (): Buffer => fs.readFileSync(path.join(ROOT, rel));

const corpus = (): Source[] => {
  const out: Source[] = [];
  for (const file of fs.readdirSync(path.join(ROOT, 'test-gallery')).filter((f) => /^T\d.*\.docx$/.test(f)).sort()) {
    out.push({ name: `gallery/${file}`, kind: 'single', load: read(`test-gallery/${file}`) });
  }
  out.push({ name: 'validator/valid.docx', kind: 'single', load: read('packages/ooxml-validator/tests/fixtures/valid.docx') });
  out.push({ name: 'new/blank', kind: 'single', load: () => createBlankDocxBuffer('blank') });
  for (const [key, shape] of Object.entries(SECTION_SHAPES)) {
    out.push({ name: `shape/${key}`, kind: shape.sections === 1 ? 'single' : 'multi', load: shape.build });
  }
  out.push({ name: 'python-docx/three-sections', kind: 'multi', load: read('tests/fixtures/sections/python-docx-three-sections.docx') });
  out.push({ name: 'libreoffice/two-page-styles', kind: 'multi', load: read('tests/fixtures/sections/libreoffice-two-page-styles.docx') });
  return out;
};

interface Observation {
  frozen: Record<string, string>;
  info: Record<string, unknown>;
}

async function observe(buf: Buffer, kind: Kind): Promise<Observation> {
  const parser = new DocxParser(buf);
  const ast = await parser.getAST();
  const nodes = ast.sections.flatMap((s) => s.nodes);
  const all: Record<string, string> = {
    flat: sha(nodes.map((n) => [n.id, n.type, n.metadata?.text ?? ''].join('\u0001')).join('\n')),
    markdown: sha(convertAstToMarkdown(ast, parser)),
    html: sha(convertAstToHtml(ast, 'Document', parser)),
    stats: sha(JSON.stringify(computeDocStats(ast))),
    noop: sha(parser.rezip()),
    toon: sha(JSON.stringify(new ToonGenerator().toon(ast, { fidelity: 'standard' }))),
    wireframe: sha(generateVisualWireframe(ast, 1)),
    firstSectionPageSetup: sha(JSON.stringify(ast.sections[0]?.pageSetup ?? null)),
    sectionId: ast.sections[0]?.id ?? '',
  };
  const frozenKeys = kind === 'single' ? Object.keys(all) : ['flat', 'markdown', 'html', 'stats', 'noop'];
  const frozen: Record<string, string> = {};
  const informational: Record<string, string> = {};
  for (const [k, v] of Object.entries(all)) (frozenKeys.includes(k) ? frozen : informational)[k] = v;
  return {
    frozen,
    info: {
      ...informational,
      sectionCount: ast.sections.length,
      sectionIds: ast.sections.map((s) => s.id),
      nodesPerSection: ast.sections.map((s) => s.nodes.length),
      nodeCount: nodes.length,
    },
  };
}

interface GoldenEntry {
  kind: Kind;
  inputSha: string;
  frozen: Record<string, string>;
  info: Record<string, unknown>;
}
interface Golden {
  note: string;
  docs: Record<string, GoldenEntry>;
}

const recorded: Record<string, GoldenEntry> = {};
const golden: Golden | null = !UPDATE && fs.existsSync(GOLDEN_PATH) ? (JSON.parse(fs.readFileSync(GOLDEN_PATH, 'utf8')) as Golden) : null;

describe('opening a document gives the same outputs as before the sections change', () => {
  const sources = corpus();

  it('the corpus covers every shape, the generated files and the gallery', () => {
    const names = sources.map((s) => s.name);
    expect(names.filter((n) => n.startsWith('gallery/'))).toHaveLength(9);
    expect(names.filter((n) => n.startsWith('shape/'))).toHaveLength(Object.keys(SECTION_SHAPES).length);
    expect(sources.filter((s) => s.kind === 'multi').length).toBeGreaterThanOrEqual(10);
  });

  for (const source of sources) {
    it(`${source.name}`, async () => {
      const buf = source.load();
      const first = await observe(buf, source.kind);
      // The observation itself must be repeatable, or a hash mismatch would mean nothing.
      expect(await observe(buf, source.kind)).toEqual(first);

      if (UPDATE) {
        recorded[source.name] = { kind: source.kind, inputSha: sha(buf), frozen: first.frozen, info: first.info };
        return;
      }
      expect(golden, `${path.relative(ROOT, GOLDEN_PATH)} is missing; write it with ASTRA_UPDATE_GOLDEN=1`).not.toBeNull();
      const entry = golden!.docs[source.name];
      expect(entry, `no golden entry for ${source.name}; rewrite the golden file`).toBeDefined();
      expect(sha(buf), `${source.name}: the input file changed, so its golden values no longer apply; rewrite the golden file`).toBe(entry.inputSha);
      expect(first.frozen).toEqual(entry.frozen);
    });
  }

  afterAll(() => {
    if (!UPDATE) return;
    const docs = Object.fromEntries(Object.entries(recorded).sort(([a], [b]) => a.localeCompare(b)));
    const note =
      'Hashes of what opening each document produces. "frozen" values must not change; "info" is recorded for reference and is expected to change for multi-section files. Rewrite with ASTRA_UPDATE_GOLDEN=1.';
    fs.writeFileSync(GOLDEN_PATH, JSON.stringify({ note, docs } satisfies Golden, null, 2) + '\n');
  });
});

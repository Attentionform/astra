import { describe, it, expect, vi, afterEach, beforeEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { DocxParser, ToonGenerator } from '@attentionform/astra-engine';
import { run } from '../../packages/cli/src/cli';

/**
 * `--section` takes a section id or a node id. An id that matches neither is an error that lists the section ids,
 * at every fidelity; before, the standard and compact tiers printed no rows and said nothing.
 */
let dir: string;
let file: string;
let seen: string[];

const text = (): string => seen.join('\n');

beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-section-'));
  file = path.join(dir, 'doc.docx');
  fs.writeFileSync(file, DocxParser.createNewDocument('blank').rezip());
  seen = [];
  vi.spyOn(console, 'log').mockImplementation((...a: unknown[]) => void seen.push(a.join(' ')));
  vi.spyOn(console, 'error').mockImplementation((...a: unknown[]) => void seen.push(a.join(' ')));
});

afterEach(() => {
  vi.restoreAllMocks();
  fs.rmSync(dir, { recursive: true, force: true });
});

async function firstNodeId(): Promise<string> {
  const ast = await new DocxParser(fs.readFileSync(file)).getAST();
  return ast.sections[0].nodes[0].id;
}

describe('ToonGenerator section scope', () => {
  it('takes an exact section id', async () => {
    const ast = await new DocxParser(fs.readFileSync(file)).getAST();
    const table = new ToonGenerator().toon(ast, { sectionId: ast.sections[0].id });
    expect(table.rows.length).toBeGreaterThan(0);
  });

  it('takes an exact node id and gives that node only', async () => {
    const ast = await new DocxParser(fs.readFileSync(file)).getAST();
    const id = ast.sections[0].nodes[0].id;
    const table = new ToonGenerator().toon(ast, { sectionId: id });
    expect(table.rows.length).toBe(1);
  });

  it('takes a prefix of at least 4 characters that matches a node', async () => {
    const ast = await new DocxParser(fs.readFileSync(file)).getAST();
    const id = ast.sections[0].nodes[0].id;
    expect(id.length).toBeGreaterThanOrEqual(5);
    const table = new ToonGenerator().toon(ast, { sectionId: id.slice(0, 4) });
    expect(table.rows.length).toBeGreaterThanOrEqual(1);
  });

  it('refuses a prefix of 3 characters or fewer, even when it would match', async () => {
    const ast = await new DocxParser(fs.readFileSync(file)).getAST();
    const id = ast.sections[0].nodes[0].id;
    for (const fidelity of ['compact', 'standard'] as const) {
      expect(() => new ToonGenerator().toon(ast, { sectionId: id.slice(0, 3), fidelity })).toThrow(/matched no section or node/);
    }
  });

  it('refuses an unknown id at every fidelity and lists the section ids', async () => {
    const ast = await new DocxParser(fs.readFileSync(file)).getAST();
    for (const fidelity of ['compact', 'standard', 'full'] as const) {
      let error: Error | undefined;
      try {
        new ToonGenerator().toon(ast, { sectionId: 'no_such_id', fidelity });
      } catch (e) {
        error = e as Error;
      }
      expect(error, fidelity).toBeDefined();
      expect(error!.message).toContain("'no_such_id' matched no section or node");
      expect(error!.message).toContain(ast.sections[0].id);
      expect((error as Error & { code?: string }).code).toBe('E_VALIDATION');
    }
  });
});

describe('astra inspect --section', () => {
  it('fails on an unknown id with exit 1, lists the section ids, and says Remedy once', async () => {
    expect(await run(['inspect', file, '--section', 'bogus'])).toBe(1);
    const out = text();
    expect(out).toContain('matched no section or node');
    expect(out).toContain('sec_root');
    expect(out.match(/Remedy:/g)?.length).toBe(1);
  });

  it('fails the same way at full fidelity, with one Remedy', async () => {
    expect(await run(['inspect', file, '--fidelity', 'full', '--section', 'bogus'])).toBe(1);
    expect(text().match(/Remedy:/g)?.length).toBe(1);
  });

  it('accepts the section id and a node id', async () => {
    expect(await run(['inspect', file, '--section', 'sec_root'])).toBe(0);
    const id = await firstNodeId();
    seen.length = 0;
    expect(await run(['inspect', file, '--section', id])).toBe(0);
    expect(text()).toContain(id);
  });

  it('prints the section ids in the footer', async () => {
    expect(await run(['inspect', file])).toBe(0);
    expect(text()).toMatch(/sections=1 \[sec_root\]/);
  });

  it('carries the section ids in --json', async () => {
    expect(await run(['inspect', file, '--json'])).toBe(0);
    const envelope = JSON.parse(seen.find((l) => l.trim().startsWith('{')) ?? '{}');
    expect(envelope.sections).toBe(1);
    expect(envelope.sectionIds).toEqual(['sec_root']);
  });

  it('requires a value after --section', async () => {
    expect(await run(['inspect', file, '--section'])).toBe(2);
  });

  it('cuts a long list of section ids in the error', async () => {
    const ast = await new DocxParser(fs.readFileSync(file)).getAST();
    const first = ast.sections[0];
    ast.sections = Array.from({ length: 45 }, (_, i) => ({ ...first, id: i === 0 ? 'sec_root' : `sec_${i + 1}`, nodes: [] }));
    let message = '';
    try {
      new ToonGenerator().toon(ast, { sectionId: 'nothing_like_it' });
    } catch (e) {
      message = (e as Error).message;
    }
    expect(message).toContain('sec_root, sec_2');
    expect(message).toContain('+25 more');
    expect(message).not.toContain('sec_45');
  });
});

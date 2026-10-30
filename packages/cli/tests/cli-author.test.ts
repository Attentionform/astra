/** CLI --author: the default author for tracked changes and comments that do not name one. */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { run } from '../src/cli';
import { DocxParser, createBlankDocxBuffer } from '@attentionform/astra-engine';

let dir: string;
beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-cli-author-'));
  vi.spyOn(process.stdout, 'write').mockImplementation(() => true);
  vi.spyOn(process.stderr, 'write').mockImplementation(() => true);
  vi.spyOn(console, 'log').mockImplementation(() => undefined);
  vi.spyOn(console, 'error').mockImplementation(() => undefined);
  delete process.env.ASTRA_AUTHOR;
});
afterEach(() => {
  vi.restoreAllMocks();
  fs.rmSync(dir, { recursive: true, force: true });
  delete process.env.ASTRA_AUTHOR;
});

async function fixture(): Promise<{ docx: string; ops: string; out: string }> {
  const docx = path.join(dir, 'in.docx');
  fs.writeFileSync(docx, createBlankDocxBuffer('blank'));
  const ast: any = await new DocxParser(fs.readFileSync(docx)).getAST();
  const ops = path.join(dir, 'ops.json');
  fs.writeFileSync(ops, JSON.stringify({ operations: [{ action: 'replace', target: ast.sections[0].nodes[0].id, text: 'edited', trackChanges: true }] }));
  return { docx, ops, out: path.join(dir, 'out.docx') };
}
const authors = (file: string) =>
  [...new AdmZip(file).readAsText('word/document.xml').matchAll(/<w:(?:ins|del) [^>]*w:author="([^"]*)"/g)].map((m) => m[1]);

describe('astra batch --author', () => {
  it('attributes tracked changes to the author given, and leaves no environment behind', async () => {
    const f = await fixture();
    expect(await run(['batch', f.docx, f.ops, '--out', f.out, '--author', 'Cli Person'])).toBe(0);
    expect(new Set(authors(f.out))).toEqual(new Set(['Cli Person']));
    expect(process.env.ASTRA_AUTHOR).toBeUndefined();
  });

  it('accepts --author=<name>', async () => {
    const f = await fixture();
    expect(await run(['batch', f.docx, f.ops, '--out', f.out, '--author=Eq Person'])).toBe(0);
    expect(new Set(authors(f.out))).toEqual(new Set(['Eq Person']));
  });

  it('defaults to ASTra without it', async () => {
    const f = await fixture();
    expect(await run(['batch', f.docx, f.ops, '--out', f.out])).toBe(0);
    expect(new Set(authors(f.out))).toEqual(new Set(['ASTra']));
  });

  it('refuses a missing name', async () => {
    const f = await fixture();
    expect(await run(['batch', f.docx, f.ops, '--out', f.out, '--author'])).not.toBe(0);
    expect(fs.existsSync(f.out)).toBe(false);
  });
});

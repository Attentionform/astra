import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run as runCli, EXIT_OK } from '../../packages/cli/src/cli';

/**
 * `astra stats` and the `astra inspect` footer print the estimate as the code computes it (450 words a page, plus
 * fixed weights), say it is a rough estimate, and show how much of it is not counted from words. They never
 * print a 500-words-a-page formula or an accuracy figure.
 */
let dir: string;
let lines: string[];

beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-stats-cli-'));
  lines = [];
  vi.spyOn(console, 'log').mockImplementation((...a: unknown[]) => void lines.push(a.join(' ')));
  vi.spyOn(console, 'error').mockImplementation((...a: unknown[]) => void lines.push(a.join(' ')));
});
afterEach(() => {
  vi.restoreAllMocks();
  fs.rmSync(dir, { recursive: true, force: true });
});

async function make(name: string, markdown: string): Promise<string> {
  const md = path.join(dir, `${name}.md`);
  const docx = path.join(dir, `${name}.docx`);
  fs.writeFileSync(md, markdown);
  expect(await runCli(['create', docx, '--from-md', md])).toBe(EXIT_OK);
  lines = [];
  return docx;
}

const PROSE = '# Title\n\nSome plain prose in a short document.\n';
const WITH_TABLE = '# Title\n\nSome plain prose.\n\n| A | B |\n| - | - |\n| 1 | 2 |\n';

describe('astra stats', () => {
  it('prose only: the rough-estimate label, the 450 recipe, no breakdown line, no 500', async () => {
    const docx = await make('prose', PROSE);
    expect(await runCli(['stats', docx])).toBe(EXIT_OK);
    const out = lines.join('\n');
    expect(out).toMatch(/pages~\d+ \(rough estimate\)/);
    expect(out).toContain('pages~ = max(1, ceil(words/450 + 0.45 per table + 0.5 per image + 1 per page break))');
    expect(out).not.toMatch(/500 wpp|words\/500/);
    expect(out).not.toContain('of which');
  });

  it('with a table: the breakdown line says how much is not counted from words, with no accuracy figure', async () => {
    const docx = await make('table', WITH_TABLE);
    expect(await runCli(['stats', docx])).toBe(EXIT_OK);
    const out = lines.join('\n');
    expect(out).toContain('pages~ of which ~0.45 pages are counted from images, tables and page breaks');
    expect(out).toContain('astra visual');
    expect(out).not.toMatch(/accura|confiden|within/i);
  });

  it('--json carries pageTerms, and the existing fields are unchanged', async () => {
    const docx = await make('json', WITH_TABLE);
    expect(await runCli(['--json', 'stats', docx])).toBe(EXIT_OK);
    const json = JSON.parse(lines.filter((l) => l.trim().startsWith('{')).pop()!);
    expect(json.pageTerms).toEqual({ words: expect.any(Number), tables: 0.45, images: 0, pageBreaks: 0 });
    for (const key of ['ok', 'file', 'words', 'paragraphs', 'pages', 'tables', 'callouts', 'images']) expect(json).toHaveProperty(key);
  });
});

describe('astra inspect footer', () => {
  it('says rough estimate, includes the breakdown for a table, and never 500 wpp', async () => {
    const withTable = await make('inspect-table', WITH_TABLE);
    expect(await runCli(['inspect', withTable])).toBe(EXIT_OK);
    const out = lines.join('\n');
    expect(out).toMatch(/pages~\d+ \(rough estimate, of which ~0\.45 pages/);
    expect(out).not.toMatch(/500 wpp/);

    const prose = await make('inspect-prose', PROSE);
    expect(await runCli(['inspect', prose])).toBe(EXIT_OK);
    const proseOut = lines.join('\n');
    expect(proseOut).toMatch(/pages~\d+ \(rough estimate\)/);
    expect(proseOut).not.toContain('of which');
  });
});

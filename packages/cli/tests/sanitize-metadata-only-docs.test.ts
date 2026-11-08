import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run } from '../src/cli';
import { AdmZip } from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../../../tests/helpers/reproducible';

useReproducibleOutput();

const ROOT = path.resolve(__dirname, '..', '..', '..');
const GUIDE = fs.readFileSync(path.join(ROOT, 'docs', 'cli-guide.md'), 'utf8');
// A Word-made file with a deleted paragraph that carries a comment thread.
const WORD_FIXTURE = path.join(ROOT, 'tests', 'fixtures', 'word-revisions', 'delete-paragraph-comment-thread.docx');

/** Cells of the sanitization-mode table row whose first cell names `mode`, without relying on column widths or spacing. */
function modeRow(mode: string): string[] {
  const line = GUIDE.split('\n').find(l => l.trimStart().startsWith('|') && l.split('|')[1]?.replace(/[`*\s]/g, '') === mode.replace(/[`*\s]/g, ''));
  if (!line) throw new Error(`no row for ${mode} in docs/cli-guide.md`);
  return line.split('|').slice(1, -1).map(c => c.replace(/[*\s]/g, '').toLowerCase());
}

const dirs: string[] = [];
beforeEach(() => {
  vi.spyOn(console, 'log').mockImplementation(() => {});
  vi.spyOn(console, 'error').mockImplementation(() => {});
});
afterEach(() => {
  vi.restoreAllMocks();
  while (dirs.length) fs.rmSync(dirs.pop()!, { recursive: true, force: true });
});

async function sanitize(...flags: string[]): Promise<{ status: number; zip?: AdmZip }> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'san-docs-'));
  dirs.push(dir);
  const out = path.join(dir, 'out.docx');
  const status = await run(['sanitize', WORD_FIXTURE, out, ...flags]);
  return { status, zip: status === 0 ? new AdmZip(fs.readFileSync(out)) : undefined };
}

describe('docs/cli-guide.md describes what --metadata-only does to review history', () => {
  it('the table gives --metadata-only the same comment, tracked-change and body cells as bare sanitize', () => {
    const bare = modeRow('Bare `sanitize`');
    const only = modeRow('`--metadata-only`');
    expect(only.slice(1)).toEqual(bare.slice(1));
    // columns: mode, core properties, RSID, comments, tracked changes, body PII
    expect(only[3]).toBe('scrubbed');
    expect(only[4]).toBe('accepted');
  });

  it('the guide says plainly that it does not keep review history and can fail closed', () => {
    expect(GUIDE).toMatch(/--metadata-only` is the same as bare `sanitize`/);
    expect(GUIDE).toMatch(/does not preserve review history/);
    expect(GUIDE).toMatch(/fails with `E_VALIDATION` and writes no file/);
  });

  it('the command does what the table says: no comments, no tracked changes left, same as bare', async () => {
    const [only, bare] = [await sanitize('--metadata-only'), await sanitize()];
    for (const result of [only, bare]) {
      expect(result.status).toBe(0);
      const names = result.zip!.getEntries().map(e => e.entryName);
      expect(names.filter(n => /comments/i.test(n))).toEqual([]);
      const body = result.zip!.readAsText('word/document.xml');
      expect(body).not.toMatch(/<w:(?:del|ins|moveFrom|moveTo)[ >]/);
      expect(body).not.toContain('commentReference');
    }
    expect(only.zip!.readAsText('word/document.xml')).toBe(bare.zip!.readAsText('word/document.xml'));
  });

  it('a flag-only run that does not accept revisions leaves them in place (the documented way to keep them)', async () => {
    const result = await sanitize('--strip-comments');
    expect(result.status).toBe(0);
    expect(result.zip!.readAsText('word/document.xml')).toMatch(/<w:del[ >]/);
  });
});

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run } from '../src/cli';
import { AdmZip, DocxParser, getHistoryCommitCount, getRepoDir } from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../../../tests/helpers/reproducible';

useReproducibleOutput();

const dirs: string[] = [];

function smartTagDocx(): { dir: string; file: string } {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-cli-normalize-repair-'));
  dirs.push(dir);
  const zip = new AdmZip(DocxParser.createNewDocument('blank').rezip());
  const xml = zip.getEntry('word/document.xml')?.getData().toString('utf8') as string;
  const anchor = xml.indexOf('<w:sectPr') >= 0 ? xml.indexOf('<w:sectPr') : xml.lastIndexOf('</w:body>');
  const withTag =
    xml.slice(0, anchor) +
    '<w:p><w:smartTag w:uri="urn:x" w:element="place"><w:r><w:t>kept text</w:t></w:r></w:smartTag></w:p>' +
    xml.slice(anchor);
  zip.updateFile('word/document.xml', Buffer.from(withTag, 'utf8'));
  const file = path.join(dir, 'in.docx');
  fs.writeFileSync(file, zip.toBuffer());
  return { dir, file };
}

function jsonOut(): Array<Record<string, unknown>> {
  return vi
    .mocked(console.log)
    .mock.calls.map((c) => {
      try {
        return JSON.parse(String(c[0]));
      } catch {
        return undefined;
      }
    })
    .filter((x): x is Record<string, unknown> => x !== undefined && typeof x === 'object');
}

beforeEach(() => {
  vi.spyOn(console, 'log').mockImplementation(() => undefined);
  vi.spyOn(console, 'error').mockImplementation(() => undefined);
});

afterEach(() => {
  vi.restoreAllMocks();
  for (const d of dirs.splice(0)) fs.rmSync(d, { recursive: true, force: true });
});

describe('normalize repairs known validator error kinds', () => {
  it('unwraps a smart tag, keeps the text, reports remaining and rolledBack', async () => {
    const { dir, file } = smartTagDocx();
    const out = path.join(dir, 'out.docx');
    await expect(run(['--json', 'normalize', file, '--out', out])).resolves.toBe(0);
    const result = jsonOut().find((j) => j.ok === true && 'normalized' in j) as Record<string, unknown>;
    expect(result).toBeDefined();
    expect((result.normalized as string[]).some((n) => n.includes('smartTag'))).toBe(true);
    expect(typeof result.remaining).toBe('number');
    expect(result.rolledBack).toEqual([]);
    const xml = new AdmZip(fs.readFileSync(out)).getEntry('word/document.xml')?.getData().toString('utf8') as string;
    expect(xml).not.toContain('smartTag');
    expect(xml).toContain('kept text');
  });

  it('a second normalize of the result changes nothing', async () => {
    const { dir, file } = smartTagDocx();
    const first = path.join(dir, 'first.docx');
    const second = path.join(dir, 'second.docx');
    await expect(run(['normalize', file, '--out', first])).resolves.toBe(0);
    await expect(run(['normalize', first, '--out', second])).resolves.toBe(0);
    expect(fs.readFileSync(second).equals(fs.readFileSync(first))).toBe(true);
  });

  it('--dry-run writes nothing and does not run the repair rules', async () => {
    const { dir, file } = smartTagDocx();
    const out = path.join(dir, 'never.docx');
    await expect(run(['normalize', file, '--out', out, '--dry-run'])).resolves.toBe(0);
    expect(fs.existsSync(out)).toBe(false);
  });
});

const docXml = (file: string): string => new AdmZip(fs.readFileSync(file)).getEntry('word/document.xml')?.getData().toString('utf8') ?? '';
const commits = (file: string): Promise<number> => getHistoryCommitCount(getRepoDir(file));

describe('normalize and the history repository', () => {
  it('records the state before the repairs and the result, and one revert returns the original text and tags', async () => {
    const { dir, file } = smartTagDocx();
    await run(['history', file, '--init']);
    const ops = path.join(dir, 'ops.json');
    fs.writeFileSync(ops, JSON.stringify({ operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'edited outside history' }] }));
    await run(['batch', file, ops, '--no-history']); // the file now differs from the head of the history
    const before = await commits(file);
    await expect(run(['--json', 'normalize', file])).resolves.toBe(0);
    expect(await commits(file)).toBe(before + 2);
    expect(docXml(file)).not.toContain('smartTag');

    await expect(run(['revert', file, '--steps', '1'])).resolves.toBe(0);
    expect(docXml(file)).toContain('smartTag');
    expect(docXml(file)).toContain('edited outside history');
  });

  it('records only the result when the file equals the head of the history', async () => {
    const { file } = smartTagDocx();
    await run(['history', file, '--init']);
    const before = await commits(file);
    await expect(run(['normalize', file])).resolves.toBe(0);
    expect(await commits(file)).toBe(before + 1);
  });

  it('refuses before writing the file while a history initialization lock is held', async () => {
    const { file } = smartTagDocx();
    await run(['history', file, '--init']);
    const before = fs.readFileSync(file);
    const commitsBefore = await commits(file);
    const lock = path.join(getRepoDir(file), '.astra-init.lock');
    fs.writeFileSync(lock, JSON.stringify({ version: 2, pid: 999999, hostname: 'other-host', token: 'held', startedAt: new Date().toISOString() }));
    await expect(run(['normalize', file])).resolves.not.toBe(0);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
    fs.rmSync(lock);
    expect(await commits(file)).toBe(commitsBefore);
  });

  it('refuses before writing the file while an initialization lock is held and .git does not exist yet', async () => {
    const { file } = smartTagDocx();
    const before = fs.readFileSync(file);
    const repo = getRepoDir(file);
    fs.mkdirSync(repo, { recursive: true });
    const lock = path.join(repo, '.astra-init.lock');
    fs.writeFileSync(lock, JSON.stringify({ version: 2, pid: process.pid, hostname: os.hostname(), token: 'held', startedAt: new Date().toISOString() }));
    expect(fs.existsSync(path.join(repo, '.git'))).toBe(false);
    await expect(run(['normalize', file])).resolves.not.toBe(0);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
    fs.rmSync(lock);
    await expect(run(['normalize', file])).resolves.toBe(0);
  });

  it('records nothing with --no-history, with --out, or without a history repository', async () => {
    const a = smartTagDocx();
    await run(['history', a.file, '--init']);
    const n = await commits(a.file);
    await expect(run(['normalize', a.file, '--no-history'])).resolves.toBe(0);
    expect(await commits(a.file)).toBe(n);

    const b = smartTagDocx();
    await run(['history', b.file, '--init']);
    const m = await commits(b.file);
    await expect(run(['normalize', b.file, '--out', path.join(b.dir, 'copy.docx')])).resolves.toBe(0);
    expect(await commits(b.file)).toBe(m);

    const c = smartTagDocx();
    await expect(run(['normalize', c.file])).resolves.toBe(0);
    expect(fs.existsSync(getRepoDir(c.file))).toBe(false);
  });
});

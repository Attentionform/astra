import { describe, it, expect, vi, afterEach, beforeEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { createBlankDocxBuffer } from '@attentionform/astra-engine';
import { run } from '../../packages/cli/src/cli';

/**
 * `validate --json` lists the errors it found, not only the first ten: up to 500 by default or the number that
 * `--max-errors` gives, and `truncated: true` when the list was cut. The human output still shows ten and counts the rest.
 */
let dir: string;
let seen: string[];

function docxWithErrors(count: number): string {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  const xml = zip.readAsText('word/document.xml');
  // w:bogus is not allowed inside a paragraph: one validator error for each paragraph.
  zip.updateFile('word/document.xml', Buffer.from(xml.replace('<w:body>', '<w:body>' + '<w:p><w:bogus/></w:p>'.repeat(count))));
  const file = path.join(dir, `bad-${count}.docx`);
  fs.writeFileSync(file, zip.toBuffer());
  return file;
}

function envelope(): Record<string, any> {
  const line = seen.find((l) => l.trim().startsWith('{'));
  expect(line, 'a JSON line').toBeDefined();
  return JSON.parse(line as string);
}

beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-validate-'));
  seen = [];
  vi.spyOn(console, 'log').mockImplementation((...a: unknown[]) => void seen.push(a.join(' ')));
  vi.spyOn(console, 'error').mockImplementation((...a: unknown[]) => void seen.push(a.join(' ')));
});

afterEach(() => {
  vi.restoreAllMocks();
  fs.rmSync(dir, { recursive: true, force: true });
});

describe('validate --json error list', () => {
  it('lists every error of a document with more than ten', async () => {
    const file = docxWithErrors(30);
    expect(await run(['validate', '--json', file])).toBe(1);
    const e = envelope();
    expect(e.ok).toBe(false);
    expect(e.errors.length).toBe(30);
    expect(e.errorCount).toBe(30);
    expect(e.truncated).toBeUndefined();
  });

  it('cuts the list at --max-errors and says so', async () => {
    const file = docxWithErrors(30);
    expect(await run(['validate', '--json', '--max-errors', '12', file])).toBe(1);
    const e = envelope();
    expect(e.errors.length).toBe(12);
    expect(e.truncated).toBe(true);
    expect(e.maxErrors).toBe(12);
  });

  it('cuts the list at 500 by default and says so', async () => {
    const file = docxWithErrors(620);
    expect(await run(['validate', '--json', file])).toBe(1);
    const e = envelope();
    expect(e.errors.length).toBe(500);
    expect(e.truncated).toBe(true);
    expect(e.maxErrors).toBe(500);
  });

  it('a list that is exactly at the limit is not marked truncated', async () => {
    const file = docxWithErrors(30);
    expect(await run(['validate', '--json', '--max-errors=30', file])).toBe(1);
    const e = envelope();
    expect(e.errors.length).toBe(30);
    expect(e.truncated).toBeUndefined();
  });

  it('the human output shows ten errors, counts the rest and points to --json', async () => {
    const file = docxWithErrors(30);
    expect(await run(['validate', file])).toBe(1);
    const text = seen.join('\n');
    expect(text).toContain('INVALID');
    expect(text.match(/^ {2}- /gm)?.length).toBe(10);
    expect(text).toContain('... and 20 more');
    expect(text).toContain('--json');
  });

  it('the human output of a document with ten errors or fewer has no "more" line', async () => {
    const file = docxWithErrors(4);
    expect(await run(['validate', file])).toBe(1);
    expect(seen.join('\n')).not.toContain('more');
  });

  it('refuses a --max-errors that is not an integer from 1 to 5000', async () => {
    const file = docxWithErrors(3);
    for (const bad of ['0', '-1', '2.5', 'abc', '5001', '']) {
      expect(await run(['validate', '--max-errors', bad, file]), bad).toBe(2);
    }
    expect(await run(['validate', file, '--max-errors'])).toBe(2);
  });

  it('a valid document is unchanged', async () => {
    const file = path.join(dir, 'ok.docx');
    fs.writeFileSync(file, createBlankDocxBuffer('blank'));
    expect(await run(['validate', '--json', file])).toBe(0);
    const e = envelope();
    expect(e.ok).toBe(true);
    expect(e.errors).toEqual([]);
  });

  it('help names --max-errors', async () => {
    await run(['validate', '--help']);
    expect(seen.join('\n')).toContain('--max-errors');
  });
});

import { describe, it, expect, vi, afterEach, beforeEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { DocxParser } from '@attentionform/astra-engine';
import { run } from '../../packages/cli/src/cli';

/**
 * `sanitize` writes a new file. It does not replace a file that is already there unless `--force` is given,
 * and it never writes over its own input.
 */
let dir: string;
let input: string;
let seen: string[];

function capture(): string {
  return seen.join('\n');
}

beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-san-ow-'));
  input = path.join(dir, 'in.docx');
  fs.writeFileSync(input, DocxParser.createNewDocument('blank').rezip());
  seen = [];
  vi.spyOn(console, 'log').mockImplementation((...a: unknown[]) => void seen.push(a.join(' ')));
  vi.spyOn(console, 'error').mockImplementation((...a: unknown[]) => void seen.push(a.join(' ')));
});

afterEach(() => {
  vi.restoreAllMocks();
  fs.rmSync(dir, { recursive: true, force: true });
});

describe('sanitize and an existing output', () => {
  it('writes a new output file', async () => {
    const out = path.join(dir, 'out.docx');
    expect(await run(['sanitize', input, out])).toBe(0);
    expect(fs.existsSync(out)).toBe(true);
  });

  it('refuses an output that exists, with exit 2, and leaves it as it was', async () => {
    const out = path.join(dir, 'out.docx');
    fs.writeFileSync(out, 'keep me');
    expect(await run(['sanitize', input, out])).toBe(2);
    expect(capture()).toContain('already exists (refusing to overwrite without --force)');
    expect(fs.readFileSync(out, 'utf8')).toBe('keep me');
  });

  it('refuses the same way for the redaction path', async () => {
    const out = path.join(dir, 'out.docx');
    fs.writeFileSync(out, 'keep me');
    expect(await run(['sanitize', input, out, '--redact-pii'])).toBe(2);
    expect(fs.readFileSync(out, 'utf8')).toBe('keep me');
  });

  it('--force replaces it', async () => {
    const out = path.join(dir, 'out.docx');
    fs.writeFileSync(out, 'replace me');
    expect(await run(['sanitize', input, out, '--force'])).toBe(0);
    expect(fs.readFileSync(out).subarray(0, 2).toString()).toBe('PK');
  });

  it('--force replaces it on the redaction path', async () => {
    const out = path.join(dir, 'out.docx');
    fs.writeFileSync(out, 'replace me');
    expect(await run(['sanitize', input, out, '--client-safe', '--force'])).toBe(0);
    expect(fs.readFileSync(out).subarray(0, 2).toString()).toBe('PK');
  });
});

describe('sanitize and its own input', () => {
  it('refuses the same path, with or without --force, and leaves the input as it was', async () => {
    const before = fs.readFileSync(input);
    expect(await run(['sanitize', input, input])).toBe(2);
    expect(capture()).toContain('same file');
    expect(await run(['sanitize', input, input, '--force'])).toBe(2);
    expect(fs.readFileSync(input).equals(before)).toBe(true);
  });

  it('refuses the same file through a different path', async () => {
    const before = fs.readFileSync(input);
    const alias = path.join(dir, 'sub', '..', 'in.docx');
    fs.mkdirSync(path.join(dir, 'sub'));
    expect(await run(['sanitize', input, alias, '--force'])).toBe(2);
    expect(fs.readFileSync(input).equals(before)).toBe(true);
  });

  it('refuses the same file through a symbolic link', async () => {
    const before = fs.readFileSync(input);
    const link = path.join(dir, 'link.docx');
    fs.symlinkSync(input, link);
    expect(await run(['sanitize', input, link, '--force'])).toBe(2);
    expect(fs.readFileSync(input).equals(before)).toBe(true);
  });
});

describe('sanitize help', () => {
  it('lists --force', async () => {
    await run(['sanitize', '--help']);
    expect(capture()).toContain('--force');
  });
});

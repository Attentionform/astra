import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { spawnSync } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { PII_COVERAGE, piiCoverageSentence } from '@attentionform/astra-engine';

/**
 * `pii=0 bodyRedactions=0` reads as "no personal data". The sanitize note says what the PII check looks for and
 * what it does not, and these tests check that the claims are true: each family it names is redacted, and the
 * identifiers it says it does not cover (CPF, CNPJ, unlabeled digit runs) are left alone.
 */
const CLI_JS = path.resolve(__dirname, '..', 'dist', 'cli.js');

function run(args: string[]): { status: number; stdout: string; stderr: string } {
  const res = spawnSync(process.execPath, [CLI_JS, ...args], { encoding: 'utf8' });
  return { status: res.status ?? 99, stdout: res.stdout ?? '', stderr: res.stderr ?? '' };
}

const COVERED: Array<[string, string]> = [
  ['an email address', 'jane.doe@example.com'],
  ['a formatted phone number', '(415) 555-2671'],
  ['a phone number after a label', 'phone 5551234567'],
  ['a dashed US SSN', '123-45-6789'],
  ['9 digits after an SSN label', 'SSN 987654321'],
  ['a 16 digit run', '4111 1111 1111 1111'],
];
const NOT_COVERED: Array<[string, string]> = [
  ['a CPF', '123.456.789-09'],
  ['a CNPJ', '12.345.678/0001-95'],
  ['9 digits without a label', '246813579'],
];

describe('sanitize says what the PII check covers', () => {
  let dir: string;
  let input: string;

  beforeAll(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-pii-note-'));
    const paragraphs = [...COVERED, ...NOT_COVERED].map(([label, value]) => `${label}: ${value}`);
    fs.writeFileSync(path.join(dir, 'in.md'), `# Test\n\n${paragraphs.join('\n\n')}\n`);
    input = path.join(dir, 'in.docx');
    expect(run(['create', input, '--from-md', path.join(dir, 'in.md')]).status).toBe(0);
  });

  afterAll(() => fs.rmSync(dir, { recursive: true, force: true }));

  const bodyOf = (file: string): string => new AdmZip(file).readAsText('word/document.xml');

  it('redacts every family that it says it checks', () => {
    const out = path.join(dir, 'redacted.docx');
    const r = run(['sanitize', input, out, '--redact-pii']);
    expect(r.status).toBe(0);
    const body = bodyOf(out);
    for (const [label, value] of COVERED) expect(body, label).not.toContain(value);
  });

  it('leaves the identifiers that it says it does not check', () => {
    const out = path.join(dir, 'redacted2.docx');
    expect(run(['sanitize', input, out, '--redact-pii']).status).toBe(0);
    const body = bodyOf(out);
    for (const [label, value] of NOT_COVERED) expect(body, label).toContain(value);
  });

  it('prints the note on stderr beside the summary, with the counts explained', () => {
    const r = run(['sanitize', input, path.join(dir, 'note.docx'), '--redact-pii']);
    expect(r.stdout).toMatch(/^SANITIZED .*pii=\d+ bodyRedactions=\d+/);
    expect(r.stderr).toContain('sanitize NOTE: The PII check looks for email addresses;');
    expect(r.stderr).toContain('(for example CPF and CNPJ)');
    expect(r.stderr).toMatch(/pii=\d+ counts document properties and bodyRedactions=\d+ counts body text\./);
    expect(r.stderr).toContain(piiCoverageSentence());
  });

  it('a metadata-only run says that the body was not checked', () => {
    const r = run(['sanitize', input, path.join(dir, 'meta.docx')]);
    expect(r.status).toBe(0);
    expect(r.stderr).toContain('counts document properties only; the body text was not checked.');
  });

  it('--json carries the coverage and keeps stderr quiet about it', () => {
    const r = run(['--json', 'sanitize', input, path.join(dir, 'json.docx'), '--redact-pii']);
    const payload = JSON.parse(r.stdout);
    expect(payload.piiCoverage).toEqual(JSON.parse(JSON.stringify(PII_COVERAGE)));
    expect(r.stderr).not.toContain('sanitize NOTE');
  });

  it('the help names the same limits', () => {
    const help = run(['sanitize', '--help']).stdout.replace(/\s+/g, ' ');
    expect(help).toContain('It does not look for other national ID numbers (for example CPF and CNPJ), names, addresses or unlabeled digit runs');
  });
});

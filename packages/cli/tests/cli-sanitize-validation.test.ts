import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run } from '../src/cli';
import { AdmZip, validateDocxBuffer } from '@attentionform/astra-engine';

const GALLERY = path.resolve(__dirname, '../../../test-gallery');
const BOGUS = '<w:p><w:r><w:bogus/></w:r></w:p>';

let dir: string;
let prevTmp: string | undefined;
let logs: string[];

function invalidCopy(name: string): string {
  const zip = new AdmZip(fs.readFileSync(path.join(GALLERY, 'T2_edit_target.docx')));
  const xml = zip.readAsText('word/document.xml');
  const at = xml.lastIndexOf('<w:sectPr');
  zip.updateFile('word/document.xml', Buffer.from(xml.slice(0, at) + BOGUS + xml.slice(at), 'utf8'));
  const file = path.join(dir, name);
  fs.writeFileSync(file, zip.toBuffer());
  return file;
}

beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-san-val-'));
  prevTmp = process.env.ASTRA_TMPDIR;
  process.env.ASTRA_TMPDIR = path.join(dir, 'tmp');
  logs = [];
  vi.spyOn(console, 'log').mockImplementation((...a: unknown[]) => void logs.push(a.join(' ')));
  vi.spyOn(console, 'error').mockImplementation((...a: unknown[]) => void logs.push(a.join(' ')));
});

afterEach(() => {
  vi.restoreAllMocks();
  if (prevTmp === undefined) delete process.env.ASTRA_TMPDIR;
  else process.env.ASTRA_TMPDIR = prevTmp;
  fs.rmSync(dir, { recursive: true, force: true });
});

const MODES: Array<[string, string[]]> = [
  ['metadata-only (default)', []],
  ['body redaction', ['--client-safe']],
];

describe('sanitize validates what it writes', () => {
  for (const [label, flags] of MODES) {
    describe(label, () => {
      it('refuses an invalid input with E_VALIDATION and writes nothing', async () => {
        const input = invalidCopy('bad.docx');
        const output = path.join(dir, 'out.docx');
        expect(await run(['sanitize', ...flags, input, output])).toBe(1);
        expect(logs.join('\n')).toContain('E_VALIDATION');
        expect(fs.existsSync(output)).toBe(false);
      });

      it('leaves an existing output file byte-identical', async () => {
        const input = invalidCopy('bad.docx');
        const output = path.join(dir, 'existing.docx');
        fs.writeFileSync(output, 'previous contents');
        expect(await run(['sanitize', ...flags, input, output])).toBe(1);
        expect(fs.readFileSync(output, 'utf8')).toBe('previous contents');
      });

      it('reports the refusal in the --json envelope', async () => {
        const input = invalidCopy('bad.docx');
        const output = path.join(dir, 'out.docx');
        expect(await run(['sanitize', '--json', ...flags, input, output])).toBe(1);
        const envelope = JSON.parse(logs.find((l) => l.trim().startsWith('{')) ?? '{}');
        expect(envelope.ok).toBe(false);
        expect(String(envelope.error)).toContain('E_VALIDATION');
        expect(fs.existsSync(output)).toBe(false);
      });

      it('still sanitizes a valid input to a valid output', async () => {
        const input = path.join(GALLERY, 'T2_edit_target.docx');
        const output = path.join(dir, 'ok.docx');
        expect(await run(['sanitize', ...flags, input, output])).toBe(0);
        expect((await validateDocxBuffer(fs.readFileSync(output))).ok).toBe(true);
      });

      it('produces a valid output for every gallery document', async () => {
        const files = fs.readdirSync(GALLERY).filter((f) => f.endsWith('.docx'));
        expect(files.length).toBeGreaterThan(0);
        for (const f of files) {
          const output = path.join(dir, `san-${f}`);
          const code = await run(['sanitize', ...flags, path.join(GALLERY, f), output]);
          expect({ file: f, code }).toEqual({ file: f, code: 0 });
          const res = await validateDocxBuffer(fs.readFileSync(output));
          expect({ file: f, ok: res.ok }).toEqual({ file: f, ok: true });
        }
      });
    });
  }
});

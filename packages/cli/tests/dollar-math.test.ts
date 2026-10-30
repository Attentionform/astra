import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { AdmZip } from '@attentionform/astra-engine';
import { run } from '../src/cli';

const SOURCE = 'We have $x^2$ here, and \\(y_1\\) there. It costs $5 and $10.\n';
let dir: string;

beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-cli-dollar-'));
  vi.spyOn(console, 'log').mockImplementation(() => {});
  vi.spyOn(console, 'error').mockImplementation(() => {});
});
afterEach(() => {
  vi.restoreAllMocks();
  fs.rmSync(dir, { recursive: true, force: true });
});

function mathCount(docx: string): number {
  const xml = new AdmZip(fs.readFileSync(docx)).readAsText('word/document.xml');
  return (xml.match(/<m:oMath>/g) ?? []).length;
}

describe('--dollar-math', () => {
  it('create --from-md reads $...$ only with the flag', async () => {
    const md = path.join(dir, 'm.md');
    fs.writeFileSync(md, SOURCE, 'utf8');
    const off = path.join(dir, 'off.docx');
    const on = path.join(dir, 'on.docx');
    await expect(run(['create', off, '--from-md', md])).resolves.toBe(0);
    await expect(run(['create', on, '--from-md', md, '--dollar-math'])).resolves.toBe(0);
    expect(mathCount(off)).toBe(1);
    expect(mathCount(on)).toBe(2);
  });

  it('import reads $...$ only with the flag', async () => {
    const md = path.join(dir, 'm.md');
    fs.writeFileSync(md, SOURCE, 'utf8');
    const off = path.join(dir, 'off.docx');
    const on = path.join(dir, 'on.docx');
    await expect(run(['import', md, '-o', off])).resolves.toBe(0);
    await expect(run(['import', md, '-o', on, '--dollar-math'])).resolves.toBe(0);
    expect(mathCount(off)).toBe(1);
    expect(mathCount(on)).toBe(2);
  });

  it('both help texts name the flag', async () => {
    for (const cmd of ['create', 'import']) {
      vi.mocked(console.log).mockClear();
      await expect(run([cmd, '--help'])).resolves.toBe(0);
      const text = vi.mocked(console.log).mock.calls.map((c) => String(c[0])).join('\n');
      expect(text).toContain('[--dollar-math]');
      expect(text).toContain('--dollar-math        also read $...$');
    }
  });
});

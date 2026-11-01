import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run } from '../src/cli';
import { AdmZip, DocxParser } from '@attentionform/astra-engine';
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

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run } from '../src/cli';
import { AdmZip, DocxParser, checkoutCommit, getHistory, getHistoryCommitCount, getRepoDir } from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../../../tests/helpers/reproducible';

useReproducibleOutput();

const dirs: string[] = [];

function duplicateIdDocx(): { dir: string; file: string; ops: string } {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-cli-dup-paraid-'));
  dirs.push(dir);
  const zip = new AdmZip(DocxParser.createNewDocument('blank').rezip());
  let xml = zip.getEntry('word/document.xml')?.getData().toString('utf8') as string;
  if (!xml.includes('xmlns:w14=')) xml = xml.replace('<w:document ', '<w:document xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ');
  const anchor = xml.indexOf('<w:sectPr') >= 0 ? xml.indexOf('<w:sectPr') : xml.lastIndexOf('</w:body>');
  xml =
    xml.slice(0, anchor) +
    '<w:p w14:paraId="5C5C5C5C"><w:r><w:t>first copy</w:t></w:r></w:p><w:p w14:paraId="5C5C5C5C"><w:r><w:t>second copy</w:t></w:r></w:p>' +
    xml.slice(anchor);
  zip.updateFile('word/document.xml', Buffer.from(xml, 'utf8'));
  const file = path.join(dir, 'in.docx');
  fs.writeFileSync(file, zip.toBuffer());
  const ops = path.join(dir, 'ops.json');
  fs.writeFileSync(ops, JSON.stringify({ operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'added by the batch' }] }));
  return { dir, file, ops };
}

const docXml = (file: string) => new AdmZip(fs.readFileSync(file)).getEntry('word/document.xml')?.getData().toString('utf8') as string;
const countId = (xml: string, id: string) => (xml.match(new RegExp(`w14:paraId="${id}"`, 'g')) ?? []).length;
const logged = () => vi.mocked(console.log).mock.calls.map((c) => String(c[0])).join('\n');

beforeEach(() => {
  vi.spyOn(console, 'log').mockImplementation(() => undefined);
  vi.spyOn(console, 'error').mockImplementation(() => undefined);
});

afterEach(() => {
  vi.restoreAllMocks();
  for (const d of dirs.splice(0)) fs.rmSync(d, { recursive: true, force: true });
});

describe('batch renumbers repeated paragraph ids', () => {
  it('writes unique ids, reports each change and keeps the earlier state in the history', async () => {
    const { file, ops } = duplicateIdDocx();
    await run(['history', file, '--init']);
    const before = await getHistoryCommitCount(getRepoDir(file));
    await expect(run(['batch', file, ops])).resolves.toBe(0);
    expect(logged()).toContain('Renumbered duplicate w14:paraId');
    const xml = docXml(file);
    expect(countId(xml, '5C5C5C5C')).toBe(1);
    expect(xml).toContain('first copy');
    expect(xml).toContain('second copy');
    expect(xml).toContain('added by the batch');
    expect(new DocxParser(fs.readFileSync(file)).getDuplicateParaIds()).toEqual([]);
    expect(await getHistoryCommitCount(getRepoDir(file))).toBe(before + 1);
    const entries = await getHistory(file);
    expect(countId((await checkoutCommit(file, entries[0].oid)).xml, '5C5C5C5C')).toBe(1);
    expect(countId((await checkoutCommit(file, entries[1].oid)).xml, '5C5C5C5C')).toBe(2);
  });

  it('records the file as it was before the batch when it differs from the head of the history', async () => {
    const clean = duplicateIdDocx();
    const other = duplicateIdDocx();
    // The history head is a document without repeated ids; the file then becomes the one with them.
    const cleanBytes = new AdmZip(fs.readFileSync(clean.file));
    cleanBytes.updateFile('word/document.xml', Buffer.from(docXml(clean.file).replace(/<w:p w14:paraId="5C5C5C5C">.*?<\/w:p>/g, ''), 'utf8'));
    fs.writeFileSync(other.file, cleanBytes.toBuffer());
    await run(['history', other.file, '--init']);
    fs.copyFileSync(clean.file, other.file);
    const before = await getHistoryCommitCount(getRepoDir(other.file));
    await expect(run(['batch', other.file, other.ops])).resolves.toBe(0);
    expect(await getHistoryCommitCount(getRepoDir(other.file))).toBe(before + 2);
    const entries = await getHistory(other.file);
    expect(countId((await checkoutCommit(other.file, entries[0].oid)).xml, '5C5C5C5C')).toBe(1);
    expect(countId((await checkoutCommit(other.file, entries[1].oid)).xml, '5C5C5C5C')).toBe(2);
  });

  it('--strict keeps the refusal and leaves the file as it was', async () => {
    const { file, ops } = duplicateIdDocx();
    const before = fs.readFileSync(file);
    const code = await run(['batch', file, ops, '--strict']);
    expect(code).not.toBe(0);
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });
});

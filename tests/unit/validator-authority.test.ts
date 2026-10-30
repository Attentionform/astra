/**
 * The OOXML validator is the only authority on element validity: when it
 * accepts a document ASTra opens, edits and saves it; when it rejects one
 * ASTra reports the same verdict. ASTra keeps no element-name list of its own.
 */
import { describe, it, expect, vi, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { AdmZip, DocxParser, OperationEngine } from '../../packages/engine/src';
import { validateDocxBuffer } from '../../packages/engine/src/ooxml/validator-gate';
import { run as runCli } from '../../packages/cli/src/cli';

const T2 = path.resolve(__dirname, '../../test-gallery/T2_edit_target.docx');
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';

/** Copy of T2 with `xml` inserted before the final sectPr; `ns` is declared on the root. */
function t2With(xml: string, ns = ''): Buffer {
  const zip = new AdmZip(fs.readFileSync(T2));
  let doc = zip.readAsText('word/document.xml');
  if (ns) doc = doc.replace('<w:document ', `<w:document ${ns} `);
  const at = doc.lastIndexOf('<w:sectPr');
  expect(at).toBeGreaterThan(0);
  doc = doc.slice(0, at) + xml + doc.slice(at);
  zip.updateFile('word/document.xml', Buffer.from(doc, 'utf8'));
  return zip.toBuffer();
}

const SDT =
  '<w:sdt><w:sdtPr><w:alias w:val="note"/><w15:appearance w15:val="hidden"/></w:sdtPr>' +
  '<w:sdtContent><w:p><w:r><w:t>sdt text</w:t></w:r></w:p></w:sdtContent></w:sdt>';
const RUBY =
  '<w:p><w:r><w:ruby><w:rubyPr><w:rubyAlign w:val="center"/><w:hps w:val="12"/><w:hpsRaise w:val="22"/><w:hpsBaseText w:val="24"/><w:lid w:val="ja-JP"/></w:rubyPr>' +
  '<w:rt><w:r><w:t>ふりがな</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>base</w:t></w:r></w:rubyBase></w:ruby></w:r></w:p>';
const BOGUS_W = '<w:p><w:r><w:bogus/></w:r></w:p>';
const BOGUS_A =
  '<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
  '<wp:extent cx="100" cy="100"/><wp:docPr id="900" name="d"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bogus/></a:graphic></wp:inline></w:drawing></w:r></w:p>';

const CASES: Array<{ name: string; bytes: () => Buffer }> = [
  { name: 'w15:appearance in sdtPr', bytes: () => t2With(SDT, `xmlns:w15="${W15}"`) },
  { name: 'w:ruby', bytes: () => t2With(RUBY) },
  { name: 'w:permStart / w:permEnd', bytes: () => t2With('<w:permStart w:id="1" w:edGrp="everyone"/><w:permEnd w:id="1"/>') },
  { name: 'w:bogus', bytes: () => t2With(BOGUS_W) },
  { name: 'a:bogus in a drawing', bytes: () => t2With(BOGUS_A) },
];

const tmp: string[] = [];
afterEach(() => {
  vi.restoreAllMocks();
  for (const f of tmp.splice(0)) fs.rmSync(f, { recursive: true, force: true });
});

describe('validator authority over element names', () => {
  it('opens, edits and saves a document the validator accepts (w15:appearance)', async () => {
    const bytes = t2With(SDT, `xmlns:w15="${W15}"`);
    expect((await validateDocxBuffer(bytes)).ok).toBe(true);
    const parser = new DocxParser(bytes);
    expect((await parser.validateOOXML()).ok).toBe(true);
    const engine = new OperationEngine(await parser.getAST(), parser);
    const res = await engine.executeBatch({
      operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'added' }],
      sequential: true,
    });
    expect(res.success).toBe(true);
    const saved = parser.rezip();
    const xml = new AdmZip(saved).readAsText('word/document.xml');
    expect(xml).toContain('w15:appearance');
    expect(xml).toContain('added');
    expect(xml).not.toContain('data-astra-id');
    expect((await validateDocxBuffer(saved)).ok).toBe(true);
  });

  it('opens, edits and saves a document with w:ruby', async () => {
    const parser = new DocxParser(t2With(RUBY));
    const engine = new OperationEngine(await parser.getAST(), parser);
    const res = await engine.executeBatch({
      operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'after ruby' }],
      sequential: true,
    });
    expect(res.success).toBe(true);
    const saved = parser.rezip();
    expect(new AdmZip(saved).readAsText('word/document.xml')).toContain('w:ruby');
    expect((await validateDocxBuffer(saved)).ok).toBe(true);
  });

  it('reports w:bogus and a:bogus as invalid, with an xPath', async () => {
    for (const xml of [BOGUS_W, BOGUS_A]) {
      const res = await new DocxParser(t2With(xml)).validateOOXML();
      expect(res.ok).toBe(false);
      const first = res.errors[0] as unknown as Record<string, unknown>;
      expect(String(first['description'])).toContain('bogus');
      expect(typeof first['xPath']).toBe('string');
    }
  });

  it('agrees with validateDocxBuffer in both directions', async () => {
    for (const c of CASES) {
      const bytes = c.bytes();
      const oracle = (await validateDocxBuffer(bytes)).ok;
      const astra = (await new DocxParser(bytes).validateOOXML()).ok;
      expect(astra, c.name).toBe(oracle);
    }
    // The cases split both ways, so the property is not vacuous.
    const verdicts = await Promise.all(CASES.map(async (c) => (await validateDocxBuffer(c.bytes())).ok));
    expect(verdicts).toContain(true);
    expect(verdicts).toContain(false);
  });

  it('CLI validate prints VALID for the w15:appearance document', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-authority-'));
    tmp.push(dir);
    const file = path.join(dir, 'ext.docx');
    fs.writeFileSync(file, t2With(SDT, `xmlns:w15="${W15}"`));
    const log = vi.spyOn(console, 'log').mockImplementation(() => {});
    vi.spyOn(console, 'error').mockImplementation(() => {});
    await expect(runCli(['validate', file])).resolves.toBe(0);
    expect(log.mock.calls.map((c) => String(c[0])).join('\n')).toContain('VALID');
  });
});

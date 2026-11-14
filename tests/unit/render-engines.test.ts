/**
 * Renderer / rasterizer / network-reference choices through the real render path, with stand-in programs:
 * a fake soffice and pdftoppm (POSIX scripts) and a fake PowerShell (Node). No Word, no LibreOffice, no Poppler needed.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { createHash } from 'crypto';
import { AdmZip, DocxParser, renderDocxToImages, resetRendererBreakersForTesting, convertDocxToPdfDetailed, type VisualTestHooks } from '@attentionform/astra-engine';
import { buildDocx, complexField, documentXml, readPart } from '../helpers/docx-refs';
import { fakePsHooks, withFakePsEnv } from '../helpers/fake-ps';
import { makeFakePdftoppm, makeFakeSoffice } from '../helpers/fake-bins';

const POSIX = process.platform !== 'win32';
/** Does this package still name the evil location, in its relationships or its body? (Parts are deflated: read them.) */
const leaks = (b: Buffer) => ['word/_rels/document.xml.rels', 'word/document.xml'].some((p) => readPart(b, p).includes('evil.example'));
const sha = (b: Buffer) => createHash('sha256').update(b).digest('hex');
const blank = () => DocxParser.createNewDocument('blank').rezip();

const WORD = { exe: 'C:\\Office\\WINWORD.EXE', version: 16 };
const winHooks: VisualTestHooks = { platform: 'win32', word: WORD, windowsRenderer: true, ...fakePsHooks, wordRetryPauseMs: [0, 0] };

let tmpBase: string;
let bins: { soffice: string; pdftoppm: string };
let savedTmp: string | undefined;
const saved: Record<string, string | undefined> = {};
const ENV_KEYS = ['FAKE_PS_LOG', 'ASTRA_RENDERER', 'ASTRA_PDF_TO_IMAGE', 'ASTRA_ALLOW_NETWORK_REFERENCES', 'FAKE_PS_MODE', 'FAKE_LOG', 'FAKE_COPY_INPUT', 'FAKE_COPY_SEQ', 'FAKE_PDF_PAGES', 'FAKE_TOTAL', 'FAKE_SOFFICE_FAIL', 'FAKE_PDFTOPPM_FAIL', 'FAKE_PS_COUNTER'];

beforeEach(() => {
  tmpBase = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-engines-test-'));
  savedTmp = process.env.ASTRA_TMPDIR;
  process.env.ASTRA_TMPDIR = path.join(tmpBase, 'tmp');
  fs.mkdirSync(process.env.ASTRA_TMPDIR, { recursive: true });
  for (const k of ENV_KEYS) {
    saved[k] = process.env[k];
    delete process.env[k];
  }
  if (POSIX) bins = { soffice: makeFakeSoffice(tmpBase), pdftoppm: makeFakePdftoppm(tmpBase) };
  resetRendererBreakersForTesting();
});
afterEach(() => {
  for (const k of ENV_KEYS) {
    if (saved[k] === undefined) delete process.env[k];
    else process.env[k] = saved[k];
  }
  if (savedTmp === undefined) delete process.env.ASTRA_TMPDIR;
  else process.env.ASTRA_TMPDIR = savedTmp;
  fs.rmSync(tmpBase, { recursive: true, force: true });
  resetRendererBreakersForTesting();
});

const leftovers = () => fs.readdirSync(process.env.ASTRA_TMPDIR as string).filter((n) => n.startsWith('astra-'));
function logRows(file: string): Array<Record<string, unknown>> {
  return fs.existsSync(file) ? fs.readFileSync(file, 'utf8').trim().split('\n').filter(Boolean).map((l) => JSON.parse(l)) : [];
}

describe.skipIf(!POSIX)('renderer choice through renderDocxToImages', () => {
  it('auto is LibreOffice even on a Windows with Word installed (0.5.2: Word is opt-in)', async () => {
    process.env.FAKE_LOG = path.join(tmpBase, 'log');
    process.env.FAKE_PS_MODE = 'ok';
    const r = await renderDocxToImages(blank(), { sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm, testHooks: { ...winHooks, windowsRenderer: false } });
    expect(r.backend).toBe('libreoffice');
    expect(r.renderer).toEqual({ requested: 'auto', used: 'libreoffice' });
    expect(logRows(process.env.FAKE_LOG).map((x) => x.bin)).toEqual(['soffice', 'pdftoppm']);
    expect(leftovers()).toEqual([]);
  });

  it('word renders through the PowerShell host and reports itself', async () => {
    process.env.FAKE_PS_MODE = 'ok';
    process.env.FAKE_PDF_PAGES = '3';
    const r = await renderDocxToImages(blank(), { renderer: 'word', pdfToImage: 'poppler', pdftoppmBin: bins.pdftoppm, maxPages: 2, testHooks: winHooks });
    expect(r.backend).toBe('word');
    expect(r.renderer).toEqual({ requested: 'word', used: 'word', version: 16 });
    expect(r.rasterizer).toEqual({ requested: 'poppler', used: 'poppler' });
    expect(r.totalPages).toBe(3);
    expect(r.images.map((i) => i.page)).toEqual([1, 2]);
    expect(leftovers()).toEqual([]);
  });

  it('word and the Windows renderer together use no Office or Poppler binary at all', async () => {
    process.env.FAKE_PS_MODE = 'ok';
    process.env.FAKE_TOTAL = '5';
    const r = await renderDocxToImages(blank(), { renderer: 'word', pdfToImage: 'windows', pages: [4], testHooks: winHooks, sofficeBin: '/nonexistent/soffice', pdftoppmBin: '/nonexistent/pdftoppm' });
    expect(r.renderer?.used).toBe('word');
    expect(r.rasterizer?.used).toBe('windows');
    expect(r.totalPages).toBe(5);
    expect(r.images.map((i) => i.page)).toEqual([4]);
    expect(fs.existsSync(r.images[0].path)).toBe(true);
  });

  it('an explicit word never falls back: a Word failure is the error, with its own code', async () => {
    process.env.FAKE_PS_MODE = 'password';
    await expect(renderDocxToImages(blank(), { renderer: 'word', pdfToImage: 'poppler', pdftoppmBin: bins.pdftoppm, sofficeBin: bins.soffice, testHooks: winHooks })).rejects.toMatchObject({ code: 'E_VALIDATION', failureClass: 'document' });
    process.env.FAKE_LOG = path.join(tmpBase, 'log');
    expect(logRows(process.env.FAKE_LOG).filter((r) => r.bin === 'soffice')).toEqual([]);
    expect(leftovers().length).toBeGreaterThan(0); // forensics kept for a failed render
  });

  it('word is a clear dependency error off Windows, and without Word on Windows', async () => {
    await expect(renderDocxToImages(blank(), { renderer: 'word', sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm })).rejects.toMatchObject({ code: 'E_DEPENDENCY', message: expect.stringMatching(/only available on Windows/) });
    await expect(renderDocxToImages(blank(), { renderer: 'word', testHooks: { ...winHooks, word: null } })).rejects.toMatchObject({ code: 'E_DEPENDENCY', message: expect.stringMatching(/Word is not installed/) });
  });

  it('auto with the Word-first plan falls back to LibreOffice with a visible reason, and the breaker skips Word next time', async () => {
    process.env.FAKE_PS_MODE = 'no-word';
    process.env.FAKE_LOG = path.join(tmpBase, 'log');
    process.env.FAKE_PS_LOG = path.join(tmpBase, 'pslog');
    const hooks = { ...winHooks, autoPrefersWord: true, windowsRenderer: false };
    const first = await renderDocxToImages(blank(), { sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm, testHooks: hooks });
    expect(first.backend).toBe('libreoffice');
    expect(first.renderer).toMatchObject({ requested: 'auto', used: 'libreoffice', fallbackFrom: 'word' });
    expect(first.renderer?.fallbackReason).toMatch(/Word could not be started/);
    const psCalls = () => logRows(process.env.FAKE_PS_LOG as string).filter((r) => r.word).length;
    expect(psCalls()).toBe(1);
    const second = await renderDocxToImages(blank(), { sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm, testHooks: hooks });
    expect(second.renderer).toEqual({ requested: 'auto', used: 'libreoffice' }); // Word skipped: no fallback to report
    expect(psCalls()).toBe(1);
    // an explicit word ignores the breaker
    process.env.FAKE_PS_MODE = 'ok';
    const third = await renderDocxToImages(blank(), { renderer: 'word', pdfToImage: 'poppler', pdftoppmBin: bins.pdftoppm, testHooks: hooks });
    expect(third.renderer?.used).toBe('word');
  });

  it('a document failure in Word falls back too but does not open the breaker', async () => {
    process.env.FAKE_PS_MODE = 'corrupt';
    process.env.FAKE_PS_LOG = path.join(tmpBase, 'pslog');
    const hooks = { ...winHooks, autoPrefersWord: true, windowsRenderer: false };
    for (let i = 0; i < 2; i++) {
      const r = await renderDocxToImages(blank(), { sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm, testHooks: hooks });
      expect(r.renderer?.fallbackFrom).toBe('word');
    }
    expect(logRows(process.env.FAKE_PS_LOG).filter((r) => r.word)).toHaveLength(2);
  });

  it('when both fail in auto, the LibreOffice error is thrown and Word\'s cause is named', async () => {
    process.env.FAKE_PS_MODE = 'no-word';
    process.env.FAKE_SOFFICE_FAIL = '1';
    const hooks = { ...winHooks, autoPrefersWord: true, windowsRenderer: false };
    const err = await renderDocxToImages(blank(), { sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm, testHooks: hooks }).catch((e) => e);
    expect(err.code).toBe('E_VALIDATION');
    expect(err.message).toMatch(/libreoffice PDF conversion failed/);
    expect(err.message).toMatch(/word failed first: .*Word could not be started/);
  });

  it('libreoffice never touches Word', async () => {
    process.env.FAKE_PS_MODE = 'ok';
    process.env.FAKE_LOG = path.join(tmpBase, 'log');
    const r = await renderDocxToImages(blank(), { renderer: 'libreoffice', sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm, testHooks: { ...winHooks, autoPrefersWord: true } });
    expect(r.backend).toBe('libreoffice');
    expect(logRows(process.env.FAKE_LOG).filter((x) => x.word)).toEqual([]);
  });

  it('ASTRA_RENDERER sets the default and an explicit value beats it; a bad value fails before any work', async () => {
    process.env.FAKE_PS_MODE = 'ok';
    process.env.ASTRA_RENDERER = 'word';
    const viaEnv = await renderDocxToImages(blank(), { pdfToImage: 'poppler', pdftoppmBin: bins.pdftoppm, testHooks: winHooks });
    expect(viaEnv.renderer?.used).toBe('word');
    const explicit = await renderDocxToImages(blank(), { renderer: 'libreoffice', sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm, testHooks: winHooks });
    expect(explicit.renderer?.used).toBe('libreoffice');
    process.env.ASTRA_RENDERER = 'nope';
    const before = leftovers().length;
    await expect(renderDocxToImages(blank(), { sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm })).rejects.toMatchObject({ code: 'E_VALIDATION', message: expect.stringMatching(/ASTRA_RENDERER/) });
    expect(leftovers().length).toBe(before);
  });

  it('no usable renderer or rasterizer is E_DEPENDENCY naming what is missing', async () => {
    await expect(renderDocxToImages(blank(), { sofficeBin: '/nonexistent/soffice', pdftoppmBin: bins.pdftoppm })).rejects.toMatchObject({ code: 'E_DEPENDENCY', message: expect.stringMatching(/soffice/) });
    await expect(renderDocxToImages(blank(), { sofficeBin: bins.soffice, pdftoppmBin: '/nonexistent/pdftoppm' })).rejects.toMatchObject({ code: 'E_DEPENDENCY', message: expect.stringMatching(/pdftoppm/) });
    // an explicit pdftoppm binary means "that Poppler": the Windows renderer is not a silent substitute
    await expect(renderDocxToImages(blank(), { sofficeBin: bins.soffice, pdftoppmBin: '/nonexistent/pdftoppm', testHooks: winHooks })).rejects.toMatchObject({ code: 'E_DEPENDENCY' });
  });
});

describe.skipIf(!POSIX)('rasterizer choice', () => {
  const base = () => ({ sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm });

  it('auto on Windows prefers the built-in renderer', async () => {
    process.env.FAKE_PS_MODE = 'ok';
    process.env.FAKE_TOTAL = '2';
    process.env.FAKE_PDF_PAGES = '2';
    const r = await renderDocxToImages(blank(), { sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm, pdfToImage: 'auto', testHooks: winHooks });
    expect(r.rasterizer).toEqual({ requested: 'auto', used: 'windows' });
    expect(r.totalPages).toBe(2);
  });

  it('falls back to Poppler when the Windows renderer cannot run, says why, and the breaker skips it next time', async () => {
    process.env.FAKE_PS_MODE = 'err41';
    process.env.FAKE_PS_LOG = path.join(tmpBase, 'pslog');
    process.env.FAKE_PDF_PAGES = '2';
    const o = { ...base(), pdfToImage: 'auto' as const, testHooks: winHooks };
    const a = await renderDocxToImages(blank(), o);
    expect(a.rasterizer).toMatchObject({ requested: 'auto', used: 'poppler', fallbackFrom: 'windows' });
    expect(a.rasterizer?.fallbackReason).toMatch(/Windows.Data.Pdf/);
    expect(a.totalPages).toBe(2);
    const b = await renderDocxToImages(blank(), o);
    expect(b.rasterizer).toEqual({ requested: 'auto', used: 'poppler' });
    expect(logRows(process.env.FAKE_PS_LOG).filter((r) => r.pdf)).toHaveLength(1);
  });

  it('a damaged PDF falls back too (Poppler may read it) without opening the breaker', async () => {
    process.env.FAKE_PS_MODE = 'err42';
    process.env.FAKE_PS_LOG = path.join(tmpBase, 'pslog');
    const o = { ...base(), pdfToImage: 'auto' as const, testHooks: winHooks };
    expect((await renderDocxToImages(blank(), o)).rasterizer?.fallbackFrom).toBe('windows');
    expect((await renderDocxToImages(blank(), o)).rasterizer?.fallbackFrom).toBe('windows');
    expect(logRows(process.env.FAKE_PS_LOG).filter((r) => r.pdf)).toHaveLength(2);
  });

  it('explicit windows never falls back; explicit poppler never starts PowerShell', async () => {
    process.env.FAKE_PS_MODE = 'err41';
    process.env.FAKE_LOG = path.join(tmpBase, 'log');
    process.env.FAKE_PS_LOG = path.join(tmpBase, 'pslog');
    await expect(renderDocxToImages(blank(), { ...base(), pdfToImage: 'windows', testHooks: winHooks })).rejects.toMatchObject({ code: 'E_DEPENDENCY' });
    expect(logRows(process.env.FAKE_LOG).filter((r) => r.bin === 'pdftoppm')).toEqual([]);
    fs.rmSync(process.env.FAKE_PS_LOG, { force: true });
    const r = await renderDocxToImages(blank(), { ...base(), pdfToImage: 'poppler', testHooks: winHooks });
    expect(r.rasterizer).toEqual({ requested: 'poppler', used: 'poppler' });
    expect(logRows(process.env.FAKE_PS_LOG)).toEqual([]);
  });

  it('ASTRA_PDF_TO_IMAGE works like the flag; windows off Windows is E_DEPENDENCY; a bad value is E_VALIDATION', async () => {
    process.env.FAKE_PS_MODE = 'ok';
    process.env.ASTRA_PDF_TO_IMAGE = 'poppler';
    expect((await renderDocxToImages(blank(), { sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm, testHooks: winHooks })).rasterizer?.used).toBe('poppler');
    expect((await renderDocxToImages(blank(), { sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm, pdfToImage: 'windows', testHooks: winHooks })).rasterizer?.used).toBe('windows');
    await expect(renderDocxToImages(blank(), { ...base(), pdfToImage: 'windows' })).rejects.toMatchObject({ code: 'E_DEPENDENCY', message: expect.stringMatching(/only available on Windows/) });
    process.env.ASTRA_PDF_TO_IMAGE = 'zzz';
    await expect(renderDocxToImages(blank(), base())).rejects.toMatchObject({ code: 'E_VALIDATION', message: expect.stringMatching(/ASTRA_PDF_TO_IMAGE/) });
  });

  it('page selection, totalPages and jpeg match between the two rasterizers', async () => {
    process.env.FAKE_PS_MODE = 'ok';
    process.env.FAKE_TOTAL = '12';
    process.env.FAKE_PDF_PAGES = '12';
    for (const format of ['png', 'jpeg'] as const) {
      const sel = { sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm, format, pages: [99, 3, 3, 0], maxPages: 3, testHooks: winHooks };
      const win = await renderDocxToImages(blank(), { ...sel, pdfToImage: 'windows' });
      const pop = await renderDocxToImages(blank(), { ...sel, pdfToImage: 'poppler' });
      expect(win.totalPages).toBe(pop.totalPages);
      expect(win.images.map((i) => i.page)).toEqual(pop.images.map((i) => i.page));
      expect(win.images.map((i) => path.extname(i.path))).toEqual(pop.images.map((i) => path.extname(i.path)));
      expect(win.images.map((i) => i.page)).toEqual([1, 3, 12]);
    }
    // no pages named: the first maxPages
    const w = await renderDocxToImages(blank(), { sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm, maxPages: 2, pdfToImage: 'windows', testHooks: winHooks });
    const p = await renderDocxToImages(blank(), { sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm, maxPages: 2, pdfToImage: 'poppler', testHooks: winHooks });
    expect(w.images.map((i) => i.page)).toEqual(p.images.map((i) => i.page));
  });
});

describe.skipIf(!POSIX)('network references through the render path', () => {
  const evil = () =>
    buildDocx({
      body: `<w:p>${complexField([' INCLUDEPICTURE "http://evil.example/f.png" \\d'])}</w:p>`,
      rels: [{ id: 'rId1', type: 'image', target: 'http://evil.example/p.png', mode: 'External' }],
    });
  const seen = (seq: number) => fs.readFileSync(`${path.join(tmpBase, 'seen')}.${seq}`);
  const setup = () => {
    process.env.FAKE_COPY_INPUT = path.join(tmpBase, 'seen');
    process.env.FAKE_COPY_SEQ = '0';
  };

  it('by default the renderer gets a copy with the references retargeted, and the report says so', async () => {
    setup();
    const doc = evil();
    const before = sha(doc);
    const r = await renderDocxToImages(doc, { sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm });
    const copy = seen(0);
    expect(leaks(doc)).toBe(true); // control: the original does name it
    expect(leaks(copy)).toBe(false);
    expect(sha(doc)).toBe(before);
    expect(r.references).toMatchObject({ networkReferencesAllowed: false, omissions: true, omittedCount: 2 });
    expect(JSON.stringify(r.references)).not.toContain('evil');
  });

  it('allowNetworkReferences hands the renderer the original bytes and reports what it recognised', async () => {
    setup();
    const doc = evil();
    const r = await renderDocxToImages(doc, { sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm, allowNetworkReferences: true });
    expect(sha(seen(0))).toBe(sha(doc));
    expect(r.references).toMatchObject({ networkReferencesAllowed: true, omissions: false, recognisedCount: 2 });
  });

  it('the environment default allows, and an explicit false overrides it', async () => {
    setup();
    const doc = evil();
    process.env.ASTRA_ALLOW_NETWORK_REFERENCES = '1';
    const viaEnv = await renderDocxToImages(doc, { sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm });
    expect(viaEnv.references?.networkReferencesAllowed).toBe(true);
    process.env.FAKE_COPY_SEQ = '1';
    const explicit = await renderDocxToImages(doc, { sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm, allowNetworkReferences: false });
    expect(explicit.references?.networkReferencesAllowed).toBe(false);
    expect(leaks(seen(1))).toBe(false);
  });

  it('a clean document is rendered from the very same bytes and reports nothing omitted', async () => {
    setup();
    const doc = blank();
    const r = await renderDocxToImages(doc, { sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm });
    expect(sha(seen(0))).toBe(sha(doc));
    expect(r.references).toMatchObject({ omissions: false, omittedCount: 0 });
  });

  it('the fallback engine renders the same neutralised bytes', async () => {
    setup();
    process.env.FAKE_PS_MODE = 'no-word';
    const doc = evil();
    const r = await renderDocxToImages(doc, { sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm, testHooks: { ...winHooks, autoPrefersWord: true, windowsRenderer: false } });
    expect(r.renderer?.fallbackFrom).toBe('word');
    expect(leaks(seen(0))).toBe(false);
    expect(r.references?.networkReferencesAllowed).toBe(false);
  });

  it('a package that cannot be checked fails closed before any renderer runs, and allowed mode still renders', async () => {
    process.env.FAKE_LOG = path.join(tmpBase, 'log');
    const bad = Buffer.from('this is not a zip');
    await expect(renderDocxToImages(bad, { sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm })).rejects.toMatchObject({ code: 'E_SECURITY', remedy: expect.stringMatching(/allow-network-references/) });
    expect(logRows(process.env.FAKE_LOG)).toEqual([]);
    expect(leftovers()).toEqual([]);
    const r = await renderDocxToImages(bad, { sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm, allowNetworkReferences: true });
    expect(r.references).toMatchObject({ networkReferencesAllowed: true, scanFailed: true });
  });

  it('a valid document with UTF-16 XML parts renders by default, and its network field is still neutralised', async () => {
    setup();
    const utf16 = (t: string) => Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(t, 'utf16le')]);
    const sources = '<?xml version="1.0" encoding="UTF-16"?><b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"/>';
    const clean = buildDocx({ extra: { 'customXml/item1.xml': utf16(sources) } });
    const r = await renderDocxToImages(clean, { sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm });
    expect(r.references).toMatchObject({ omissions: false, omittedCount: 0 });
    expect(sha(seen(0))).toBe(sha(clean));
    process.env.FAKE_COPY_SEQ = '1';
    const header = documentXml(`<w:p>${complexField([' INCLUDEPICTURE "http://evil.example/f.png"'])}</w:p>`);
    const evilDoc = buildDocx({ extra: { 'customXml/item1.xml': utf16(sources), 'word/header1.xml': utf16(header) } });
    const r2 = await renderDocxToImages(evilDoc, { sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm });
    expect(r2.references).toMatchObject({ omissions: true, omittedCount: 1 });
    const part = new AdmZip(seen(1)).getEntry('word/header1.xml')!.getData();
    expect(part.subarray(0, 2).equals(Buffer.from([0xff, 0xfe]))).toBe(true);
    expect(part.toString('utf16le')).not.toContain('evil.example');
  });

  it('a bad ASTRA_ALLOW_NETWORK_REFERENCES value is a loud E_VALIDATION', async () => {
    process.env.ASTRA_ALLOW_NETWORK_REFERENCES = 'perhaps';
    await expect(renderDocxToImages(blank(), { sofficeBin: bins.soffice, pdftoppmBin: bins.pdftoppm })).rejects.toMatchObject({ code: 'E_VALIDATION', message: expect.stringMatching(/ASTRA_ALLOW_NETWORK_REFERENCES/) });
  });
});

describe.skipIf(!POSIX)('PDF export through convertDocxToPdfDetailed', () => {
  const out = () => path.join(tmpBase, 'out', 'doc.pdf');
  it('LibreOffice export reports itself and the references', async () => {
    const r = await convertDocxToPdfDetailed(blank(), out(), { binary: makeFakeSoffice(tmpBase) });
    expect(r.renderer).toEqual({ requested: 'auto', used: 'libreoffice' });
    expect(r.references.omissions).toBe(false);
    expect(fs.readFileSync(r.path, 'latin1').startsWith('%PDF-')).toBe(true);
    expect(leftovers()).toEqual([]);
  });
  it('word export writes the Word PDF to the destination', async () => {
    process.env.FAKE_PS_MODE = 'ok';
    process.env.FAKE_PDF_PAGES = '2';
    const r = await convertDocxToPdfDetailed(blank(), out(), { renderer: 'word', testHooks: winHooks });
    expect(r.renderer).toEqual({ requested: 'word', used: 'word', version: 16 });
    expect(fs.readFileSync(r.path, 'latin1').startsWith('%PDF-')).toBe(true);
    expect(fs.statSync(r.path).mode & 0o777).toBe(0o644);
    expect(leftovers()).toEqual([]);
  });
  it('word export maps tagged to the script and ignores LibreOffice entirely', async () => {
    process.env.FAKE_PS_MODE = 'ok';
    process.env.FAKE_PS_LOG = path.join(tmpBase, 'pslog');
    process.env.FAKE_LOG = path.join(tmpBase, 'log');
    await withFakePsEnv({ FAKE_PS_LOG: path.join(tmpBase, 'pslog') }, async () => {
      await convertDocxToPdfDetailed(blank(), out(), { renderer: 'word', tagged: false, testHooks: winHooks, binary: makeFakeSoffice(tmpBase) });
    });
    expect(logRows(path.join(tmpBase, 'pslog'))[0]).toMatchObject({ tagged: '0' });
    expect(logRows(process.env.FAKE_LOG)).toEqual([]);
    delete process.env.FAKE_PS_LOG;
  });
  it('auto with the Word-first plan falls back to LibreOffice, reports it and leaves no temp folders', async () => {
    process.env.FAKE_PS_MODE = 'no-word';
    const r = await convertDocxToPdfDetailed(blank(), out(), { binary: makeFakeSoffice(tmpBase), testHooks: { ...winHooks, autoPrefersWord: true } });
    expect(r.renderer).toMatchObject({ requested: 'auto', used: 'libreoffice', fallbackFrom: 'word' });
    expect(leftovers()).toEqual([]);
  });
  it('an explicit word failure is thrown as it is (code, remedy), with its temp folder kept for forensics', async () => {
    process.env.FAKE_PS_MODE = 'password';
    await expect(convertDocxToPdfDetailed(blank(), out(), { renderer: 'word', testHooks: winHooks })).rejects.toMatchObject({ code: 'E_VALIDATION', failureClass: 'document' });
    expect(leftovers().length).toBe(1);
  });
  it('word is E_DEPENDENCY off Windows; an invalid renderer is E_VALIDATION; missing LibreOffice keeps the old message', async () => {
    await expect(convertDocxToPdfDetailed(blank(), out(), { renderer: 'word', binary: makeFakeSoffice(tmpBase) })).rejects.toMatchObject({ code: 'E_DEPENDENCY' });
    await expect(convertDocxToPdfDetailed(blank(), out(), { renderer: 'x' as never })).rejects.toMatchObject({ code: 'E_VALIDATION' });
    const saved = process.env.PATH;
    process.env.PATH = '';
    process.env.ASTRA_VISUAL_NO_FALLBACK_DIRS = '1';
    try {
      const { resetPdfBackendCache } = await import('@attentionform/astra-engine');
      resetPdfBackendCache();
      await expect(convertDocxToPdfDetailed(blank(), out())).rejects.toMatchObject({ code: 'E_DEPENDENCY', message: expect.stringMatching(/requires LibreOffice/) });
    } finally {
      process.env.PATH = saved;
      delete process.env.ASTRA_VISUAL_NO_FALLBACK_DIRS;
      const { resetPdfBackendCache } = await import('@attentionform/astra-engine');
      resetPdfBackendCache();
    }
  });
  it('export neutralises network references by default and passes the original when allowed', async () => {
    process.env.FAKE_COPY_INPUT = path.join(tmpBase, 'seen');
    const doc = buildDocx({ rels: [{ id: 'rId1', type: 'image', target: 'http://evil.example/p.png', mode: 'External' }] });
    const soffice = makeFakeSoffice(tmpBase);
    const a = await convertDocxToPdfDetailed(doc, out(), { binary: soffice });
    expect(a.references).toMatchObject({ omissions: true, omittedCount: 1 });
    expect(leaks(fs.readFileSync(`${path.join(tmpBase, 'seen')}.0`))).toBe(false);
    process.env.FAKE_COPY_SEQ = '1';
    const b = await convertDocxToPdfDetailed(doc, out(), { binary: soffice, allowNetworkReferences: true });
    expect(b.references).toMatchObject({ networkReferencesAllowed: true, recognisedCount: 1 });
    expect(sha(fs.readFileSync(`${path.join(tmpBase, 'seen')}.1`))).toBe(sha(doc));
  });
});

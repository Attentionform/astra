import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run } from '../src/cli';
import { readRenderFlag, renderOptionsFromFlags, type RenderFlags } from '../src/cli/render-flags';
import { DocxParser, resetPdfBackendCache, resetRendererBreakersForTesting } from '@attentionform/astra-engine';
import { buildDocx, readPart } from '../../../tests/helpers/docx-refs';
import { makeFakePdftoppm, makeFakeSoffice } from '../../../tests/helpers/fake-bins';

const POSIX = process.platform !== 'win32';

describe('readRenderFlag', () => {
  const read = (args: string[], pdfToImage = true) => {
    const flags: RenderFlags = {};
    const r = readRenderFlag('visual', args, 0, flags, { pdfToImage });
    return { r, flags };
  };
  it('reads each value flag in both forms', () => {
    expect(read(['--renderer', 'word'])).toMatchObject({ r: { matched: true, next: 1 }, flags: { renderer: 'word' } });
    expect(read(['--renderer=libreoffice'])).toMatchObject({ r: { matched: true, next: 0 }, flags: { renderer: 'libreoffice' } });
    expect(read(['--pdf-to-image', 'POPPLER'])).toMatchObject({ r: { next: 1 }, flags: { pdfToImage: 'poppler' } });
    expect(read(['--pdf-to-image=windows'])).toMatchObject({ flags: { pdfToImage: 'windows' } });
  });
  it('reads the network flags; the last one wins', () => {
    const flags: RenderFlags = {};
    readRenderFlag('visual', ['--allow-network-references'], 0, flags, { pdfToImage: true });
    expect(flags.allowNetworkReferences).toBe(true);
    readRenderFlag('visual', ['--no-allow-network-references'], 0, flags, { pdfToImage: true });
    expect(flags.allowNetworkReferences).toBe(false);
  });
  it('reads the markup flags; the last one wins; the option is passed only when given', () => {
    const flags: RenderFlags = {};
    expect(readRenderFlag('visual', ['--show-markup'], 0, flags, { pdfToImage: true })).toEqual({ matched: true, next: 0 });
    expect(flags.showMarkup).toBe(true);
    readRenderFlag('visual', ['--no-show-markup'], 0, flags, { pdfToImage: true });
    expect(flags.showMarkup).toBe(false);
    expect(renderOptionsFromFlags({})).toEqual({});
    expect(renderOptionsFromFlags({ showMarkup: false })).toEqual({ showMarkup: false });
    expect(renderOptionsFromFlags({ showMarkup: true, renderer: 'word' })).toEqual({ showMarkup: true, renderer: 'word' });
    expect(read(['--show-markups']).r).toEqual({ matched: false });
  });
  it('rejects a bad or missing value with a message that names the choices', () => {
    expect(read(['--renderer', 'excel']).r).toMatchObject({ matched: true, error: expect.stringMatching(/excel.*auto, word, libreoffice/) });
    expect(read(['--renderer']).r).toMatchObject({ error: expect.stringMatching(/requires a value/) });
    expect(read(['--renderer=']).r).toMatchObject({ error: expect.stringMatching(/requires a value/) });
    expect(read(['--pdf-to-image', 'gdi']).r).toMatchObject({ error: expect.stringMatching(/auto, windows, poppler/) });
  });
  it('does not take --pdf-to-image where it does not apply, and ignores everything else', () => {
    expect(read(['--pdf-to-image', 'poppler'], false).r).toEqual({ matched: false });
    expect(read(['--page', '2']).r).toEqual({ matched: false });
    expect(read(['--rendererx']).r).toEqual({ matched: false });
  });
  it('passes only what was given, so the environment still applies to the rest', () => {
    expect(renderOptionsFromFlags({})).toEqual({});
    expect(renderOptionsFromFlags({ renderer: 'word', allowNetworkReferences: false })).toEqual({ renderer: 'word', allowNetworkReferences: false });
  });
});

describe.skipIf(!POSIX)('render flags through the CLI', () => {
  const ENV = ['PATH', 'ASTRA_SHOW_MARKUP', 'FAKE_SOFFICE_VERSION', 'ASTRA_RENDERER', 'ASTRA_PDF_TO_IMAGE', 'ASTRA_ALLOW_NETWORK_REFERENCES', 'ASTRA_TMPDIR', 'FAKE_COPY_INPUT', 'FAKE_COPY_SEQ', 'FAKE_PDF_PAGES'];
  const saved: Record<string, string | undefined> = {};
  let base: string;

  const out = () => vi.mocked(console.log).mock.calls.map((c) => String(c[0])).join('\n');
  const err = () => vi.mocked(console.error).mock.calls.map((c) => String(c[0])).join('\n');
  async function cli(args: string[]) {
    vi.mocked(console.log).mockClear();
    vi.mocked(console.error).mockClear();
    const status = await run(args);
    return { status, stdout: out(), stderr: err() };
  }
  const writeDoc = (name: string, buf: Buffer) => {
    const f = path.join(base, name);
    fs.writeFileSync(f, buf);
    return f;
  };
  const withRef = () => buildDocx({ rels: [{ id: 'rId9', type: 'image', target: 'http://evil.example/p.png', mode: 'External' }] });
  const seenNamesEvil = (n: number) => readPart(fs.readFileSync(`${path.join(base, 'seen')}.${n}`), 'word/_rels/document.xml.rels').includes('evil.example');

  beforeEach(() => {
    for (const k of ENV) saved[k] = process.env[k];
    for (const k of ENV.slice(1)) delete process.env[k];
    base = fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(), 'astra-cli-flags-')));
    fs.mkdirSync(path.join(base, 'bin'));
    fs.mkdirSync(path.join(base, 'tmp'));
    makeFakeSoffice(path.join(base, 'bin'));
    makeFakePdftoppm(path.join(base, 'bin'));
    process.env.PATH = `${path.join(base, 'bin')}${path.delimiter}${saved.PATH ?? ''}`;
    process.env.ASTRA_TMPDIR = path.join(base, 'tmp');
    resetPdfBackendCache();
    resetRendererBreakersForTesting();
    vi.spyOn(console, 'log').mockImplementation(() => undefined);
    vi.spyOn(console, 'error').mockImplementation(() => undefined);
  });
  afterEach(() => {
    vi.restoreAllMocks();
    for (const k of ENV) {
      if (saved[k] === undefined) delete process.env[k];
      else process.env[k] = saved[k];
    }
    fs.rmSync(base, { recursive: true, force: true });
    resetPdfBackendCache();
  });

  it('visual: usage errors for a bad value, a missing value and (for export) a flag that does not apply', async () => {
    const doc = writeDoc('a.docx', DocxParser.createNewDocument('blank').rezip());
    const o = path.join(base, 'o.png');
    expect((await cli(['visual', doc, '--out', o, '--renderer', 'excel'])).status).toBe(2);
    expect((await cli(['visual', doc, '--out', o, '--renderer'])).status).toBe(2);
    expect((await cli(['export', doc, '--format', 'pdf', '--pdf-to-image', 'poppler'])).status).toBe(2);
    const html = await cli(['export', doc, '--format', 'html', '--renderer', 'libreoffice']);
    expect(html.status).toBe(2);
    expect(html.stderr).toMatch(/PDF export only/);
  });

  it('visual: reports renderer and rasterizer in --json and writes nothing to stdout but JSON', async () => {
    const doc = writeDoc('b.docx', DocxParser.createNewDocument('blank').rezip());
    const o = path.join(base, 'b.png');
    const r = await cli(['--json', 'visual', doc, '--out', o, '--renderer=libreoffice', '--pdf-to-image', 'poppler']);
    expect(r.status).toBe(0);
    const json = JSON.parse(r.stdout);
    expect(json.renderer).toEqual({ requested: 'libreoffice', used: 'libreoffice' });
    expect(json.rasterizer).toEqual({ requested: 'poppler', used: 'poppler' });
    expect(json.references).toMatchObject({ omissions: false });
    expect(fs.existsSync(o)).toBe(true);
  });

  it('visual: network references are not loaded by default; a note on stderr says how many and how to change it', async () => {
    process.env.FAKE_COPY_INPUT = path.join(base, 'seen');
    process.env.FAKE_COPY_SEQ = '0';
    const doc = writeDoc('c.docx', withRef());
    const r = await cli(['visual', doc, '--out', path.join(base, 'c.png')]);
    expect(r.status).toBe(0);
    expect(r.stderr).toMatch(/note: 1 network reference was not loaded .*--allow-network-references/);
    expect(seenNamesEvil(0)).toBe(false);
    expect(readPart(fs.readFileSync(doc), 'word/_rels/document.xml.rels')).toContain('evil.example'); // the file itself is untouched
  });

  it('visual: --allow-network-references hands the original to the renderer; --no-allow beats the environment', async () => {
    process.env.FAKE_COPY_INPUT = path.join(base, 'seen');
    const doc = writeDoc('d.docx', withRef());
    process.env.FAKE_COPY_SEQ = '0';
    expect((await cli(['visual', doc, '--out', path.join(base, 'd1.png'), '--allow-network-references'])).status).toBe(0);
    expect(seenNamesEvil(0)).toBe(true);
    process.env.ASTRA_ALLOW_NETWORK_REFERENCES = '1';
    process.env.FAKE_COPY_SEQ = '1';
    const r = await cli(['visual', doc, '--out', path.join(base, 'd2.png'), '--no-allow-network-references']);
    expect(r.status).toBe(0);
    expect(seenNamesEvil(1)).toBe(false);
  });

  it('visual --show-markup: markup block in --json, the contract, privacy and long-comment notes on stderr, stdout stays JSON', async () => {
    const doc = writeDoc('m1.docx', DocxParser.createNewDocument('blank').rezip());
    const r = await cli(['--json', 'visual', doc, '--out', path.join(base, 'm1.png'), '--show-markup']);
    expect(r.status).toBe(0);
    const json = JSON.parse(r.stdout);
    expect(json.markup).toMatchObject({ requested: true, renderer: 'libreoffice', mode: 'comments-in-margin' });
    expect(json.markup.warnings[0]).toMatch(/cut short in the margin/);
    expect(r.stderr).toMatch(/note: markup view: pages are scaled and may differ .* page numbers refer to this render/);
    expect(r.stderr).toMatch(/note: markup view: the output shows reviewer names, deleted text and comments/);
    expect(r.stderr).toMatch(/note: markup warning: LibreOffice can show a long comment cut short/);
  });

  it('visual: without the switch there is no markup block and no markup note', async () => {
    const doc = writeDoc('m2.docx', DocxParser.createNewDocument('blank').rezip());
    const r = await cli(['--json', 'visual', doc, '--out', path.join(base, 'm2.png')]);
    expect(JSON.parse(r.stdout).markup).toBeUndefined();
    expect(r.stderr).not.toMatch(/markup/i);
  });

  it('visual: ASTRA_SHOW_MARKUP turns it on and --no-show-markup beats it; a bad value is an error naming the variable', async () => {
    const doc = writeDoc('m3.docx', DocxParser.createNewDocument('blank').rezip());
    process.env.ASTRA_SHOW_MARKUP = '1';
    const on = await cli(['--json', 'visual', doc, '--out', path.join(base, 'm3a.png')]);
    expect(JSON.parse(on.stdout).markup).toMatchObject({ requested: true });
    const off = await cli(['--json', 'visual', doc, '--out', path.join(base, 'm3b.png'), '--no-show-markup']);
    expect(JSON.parse(off.stdout).markup).toBeUndefined();
    process.env.ASTRA_SHOW_MARKUP = 'sometimes';
    const bad = await cli(['visual', doc, '--out', path.join(base, 'm3c.png')]);
    expect(bad.status).toBe(1);
    expect(bad.stderr).toMatch(/ASTRA_SHOW_MARKUP/);
  });

  it('visual --show-markup with a LibreOffice below the tested minimum fails E_DEPENDENCY before any render, with a remedy', async () => {
    process.env.FAKE_SOFFICE_VERSION = 'LibreOffice 6.4.7.2 x';
    const doc = writeDoc('m4.docx', DocxParser.createNewDocument('blank').rezip());
    const r = await cli(['visual', doc, '--out', path.join(base, 'm4.png'), '--show-markup']);
    expect(r.status).toBe(1);
    expect(r.stderr).toMatch(/E_DEPENDENCY/);
    expect(r.stderr).toMatch(/has not been tested for showing comments in the margin/);
    expect(fs.existsSync(path.join(base, 'm4.png'))).toBe(false);
  });

  it('inspect --visual --show-markup carries the markup block in --json', async () => {
    const doc = writeDoc('m5.docx', DocxParser.createNewDocument('blank').rezip());
    const r = await cli(['--json', 'inspect', doc, '--visual', '--show-markup']);
    expect(r.status).toBe(0);
    expect(JSON.parse(r.stdout).visual.markup).toMatchObject({ requested: true, renderer: 'libreoffice' });
  });

  it('export --format pdf --show-markup: markup in --json and the notes; html with the flag is a usage error', async () => {
    const doc = writeDoc('m6.docx', DocxParser.createNewDocument('blank').rezip());
    const pdf = path.join(base, 'm6.pdf');
    const r = await cli(['--json', 'export', doc, '--format', 'pdf', '--out', pdf, '--show-markup']);
    expect(r.status).toBe(0);
    expect(JSON.parse(r.stdout).markup).toMatchObject({ requested: true, renderer: 'libreoffice', mode: 'comments-in-margin' });
    expect(r.stderr).toMatch(/note: markup view: the output shows reviewer names/);
    const html = await cli(['export', doc, '--format', 'html', '--show-markup']);
    expect(html.status).toBe(2);
    expect(html.stderr).toMatch(/--show-markup/);
    expect(html.stderr).toMatch(/PDF export only/);
  });

  it('visual: --renderer word off Windows fails E_DEPENDENCY with the engine remedy (exit 1)', async () => {
    const doc = writeDoc('e.docx', DocxParser.createNewDocument('blank').rezip());
    const r = await cli(['visual', doc, '--out', path.join(base, 'e.png'), '--renderer', 'word']);
    expect(r.status).toBe(1);
    expect(r.stderr).toMatch(/E_DEPENDENCY/);
    expect(r.stderr).toMatch(/only available on Windows/);
    expect(r.stderr).not.toMatch(/install LibreOffice \(soffice in PATH\)/);
  });

  it('visual: a bad ASTRA_RENDERER names the variable', async () => {
    process.env.ASTRA_RENDERER = 'nope';
    const doc = writeDoc('f.docx', DocxParser.createNewDocument('blank').rezip());
    const r = await cli(['visual', doc, '--out', path.join(base, 'f.png')]);
    expect(r.status).toBe(1);
    expect(r.stderr).toMatch(/ASTRA_RENDERER/);
  });

  it('inspect --visual takes the same flags and carries the reports in --json', async () => {
    const doc = writeDoc('g.docx', withRef());
    const r = await cli(['--json', 'inspect', doc, '--visual', '--renderer', 'libreoffice', '--pdf-to-image=poppler']);
    expect(r.status).toBe(0);
    const json = JSON.parse(r.stdout);
    expect(json.visual.renderer).toEqual({ requested: 'libreoffice', used: 'libreoffice' });
    expect(json.visual.rasterizer.used).toBe('poppler');
    expect(json.visual.references).toMatchObject({ omissions: true, omittedCount: 1 });
  });

  it('inspect --visual degrades (exit 0, telemetry kept) when the chosen renderer is unavailable, with the engine remedy', async () => {
    const doc = writeDoc('h.docx', DocxParser.createNewDocument('blank').rezip());
    const r = await cli(['--json', 'inspect', doc, '--visual', '--renderer', 'word']);
    expect(r.status).toBe(0);
    expect(JSON.parse(r.stdout).visual).toMatchObject({ skipped: 'E_DEPENDENCY', images: [] });
    expect(r.stderr).toMatch(/\[E_DEPENDENCY\].*only available on Windows/);
    expect(r.stderr).toMatch(/Remedy: .*--renderer auto/);
  });

  it('inspect --visual: a package whose references cannot be checked is not rendered (E_SECURITY), the inspection still works, and allowing network references renders it', async () => {
    // An unrelated relationships part that is not well-formed XML: the inspector reads the document, the reference check refuses to guess.
    const doc = writeDoc('i.docx', buildDocx({ extra: { 'customXml/_rels/item1.xml.rels': '<Relationships><Relationship' } }));
    process.env.FAKE_COPY_INPUT = path.join(base, 'seen');
    process.env.FAKE_COPY_SEQ = '0';
    const refused = await cli(['--json', 'inspect', doc, '--visual']);
    expect(refused.status, refused.stderr + refused.stdout.slice(0, 300)).toBe(0);
    expect(JSON.parse(refused.stdout).visual).toMatchObject({ skipped: 'E_SECURITY', images: [] });
    expect(refused.stderr).toMatch(/\[E_SECURITY\]/);
    expect(refused.stderr).toMatch(/--allow-network-references/);
    expect(fs.existsSync(`${path.join(base, 'seen')}.0`)).toBe(false); // the renderer never saw the file
    const allowed = await cli(['--json', 'inspect', doc, '--visual', '--allow-network-references']);
    expect(JSON.parse(allowed.stdout).visual.references).toMatchObject({ networkReferencesAllowed: true, scanFailed: true });
    expect(allowed.stderr).toMatch(/the scan for them failed/);
  });

  it('export --format pdf: renderer and references in --json; notes on stderr; flags accepted in = form', async () => {
    process.env.FAKE_COPY_INPUT = path.join(base, 'seen');
    process.env.FAKE_COPY_SEQ = '0';
    const doc = writeDoc('j.docx', withRef());
    const pdf = path.join(base, 'j.pdf');
    const r = await cli(['--json', 'export', doc, '--format', 'pdf', '--out', pdf, '--renderer=libreoffice']);
    expect(r.status).toBe(0);
    const json = JSON.parse(r.stdout);
    expect(json.renderer).toEqual({ requested: 'libreoffice', used: 'libreoffice' });
    expect(json.references).toMatchObject({ omissions: true, omittedCount: 1 });
    expect(r.stderr).toMatch(/note: 1 network reference was not loaded/);
    expect(seenNamesEvil(0)).toBe(false);
    process.env.FAKE_COPY_SEQ = '1';
    const again = await cli(['export', doc, '--format', 'pdf', '--out', path.join(base, 'j2.pdf'), '--allow-network-references']);
    expect(again.status).toBe(0);
    expect(seenNamesEvil(1)).toBe(true);
    expect(again.stderr).toMatch(/network references allowed; the renderer may contact 1 location/);
  });

  it('export --renderer word off Windows fails E_DEPENDENCY and writes nothing', async () => {
    const doc = writeDoc('k.docx', DocxParser.createNewDocument('blank').rezip());
    const pdf = path.join(base, 'k.pdf');
    const r = await cli(['export', doc, '--format', 'pdf', '--out', pdf, '--renderer', 'word']);
    expect(r.status).toBe(1);
    expect(r.stderr).toMatch(/E_DEPENDENCY/);
    expect(fs.existsSync(pdf)).toBe(false);
  });

  it('help texts and usage list the new flags', async () => {
    for (const cmd of ['visual', 'inspect', 'export']) {
      const h = await cli([cmd, '--help']);
      expect(h.stdout).toMatch(/--renderer/);
      expect(h.stdout).toMatch(/--allow-network-references/);
      expect(h.stdout).toMatch(/ASTRA_ALLOW_NETWORK_REFERENCES/);
    }
    expect((await cli(['visual', '--help'])).stdout).toMatch(/--pdf-to-image/);
    expect((await cli(['export', '--help'])).stdout).not.toMatch(/--pdf-to-image/);
    expect((await cli(['--help'])).stdout).toMatch(/--renderer auto\|word\|libreoffice/);
  });
});

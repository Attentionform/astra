/**
 * Show markup, engine side: the switch, the LibreOffice filter spec, the version probe, the Word script and
 * protocol, the fallback rules through the real runPlan, and the visual and export paths with stand-in programs.
 * No Word, LibreOffice or Poppler is needed.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
  DocxParser,
  MIN_LIBREOFFICE_FOR_MARKUP,
  LIBREOFFICE_LONG_COMMENT_WARNING,
  MARKUP_CONTRACT_NOTE,
  MARKUP_PRIVACY_NOTE,
  SHOW_MARKUP_ENV,
  WordRenderError,
  buildMarkupReport,
  convertDocxToPdfDetailed,
  effectiveShowMarkup,
  markupNotes,
  parseShowMarkup,
  renderDocxToImages,
  resetRendererBreakersForTesting,
  wordDocxToPdf,
  type VisualTestHooks,
} from '@attentionform/astra-engine';
import { pdfFilterSpec } from '../../packages/engine/src/export/pdf';
import {
  assertLibreOfficeMarkupSupport,
  compareVersion,
  libreOfficeFilterSpec,
  libreOfficeMarkupWarnings,
  markupFailureError,
  parseLibreOfficeVersion,
  probeLibreOfficeVersion,
  resetLibreOfficeVersionCache,
  versionProbeFile,
  type ProbeRun,
} from '../../packages/engine/src/visual/markup';
import { buildWordScript, parseWordProtocol, validateStartState } from '../../packages/engine/src/visual/word';
import { runPlan, failureClassOf } from '../../packages/engine/src/visual/render-plan';
import { CircuitBreaker } from '../../packages/engine/src/visual/renderer-preference';
import { fakePsHooks } from '../helpers/fake-ps';
import { makeFakePdftoppm, makeFakeSoffice } from '../helpers/fake-bins';

const POSIX = process.platform !== 'win32';
const blank = () => DocxParser.createNewDocument('blank').rezip();
const WORD = { exe: 'C:\\Office\\WINWORD.EXE', version: 16 };
const winHooks: VisualTestHooks = { platform: 'win32', word: WORD, windowsRenderer: true, ...fakePsHooks, wordRetryPauseMs: [0, 0] };

let tmpBase: string;
let savedTmp: string | undefined;
const saved: Record<string, string | undefined> = {};
const ENV_KEYS = [
  SHOW_MARKUP_ENV,
  'ASTRA_WORD_MARKUP',
  'ASTRA_RENDERER',
  'FAKE_PS_MODE',
  'FAKE_PS_LOG',
  'FAKE_PS_WARN',
  'FAKE_PS_RESTORE_FAIL',
  'FAKE_LOG',
  'FAKE_PDF_PAGES',
  'FAKE_SOFFICE_FAIL',
  'FAKE_SOFFICE_FAIL_MARKUP',
  'FAKE_SOFFICE_VERSION',
  'FAKE_SOFFICE_VERSION_STREAM',
  'FAKE_SOFFICE_VERSION_EXIT',
  'FAKE_SOFFICE_VERSION_HANG',
];

beforeEach(() => {
  tmpBase = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-markup-test-'));
  savedTmp = process.env.ASTRA_TMPDIR;
  process.env.ASTRA_TMPDIR = path.join(tmpBase, 'tmp');
  fs.mkdirSync(process.env.ASTRA_TMPDIR, { recursive: true });
  for (const k of ENV_KEYS) {
    saved[k] = process.env[k];
    delete process.env[k];
  }
  resetRendererBreakersForTesting();
  resetLibreOfficeVersionCache();
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
  resetLibreOfficeVersionCache();
});

function logRows(file: string): Array<Record<string, unknown>> {
  return fs.existsSync(file) ? fs.readFileSync(file, 'utf8').trim().split('\n').filter(Boolean).map((l) => JSON.parse(l)) : [];
}

describe('the switch', () => {
  it('parses booleans and the usual words, treats empty as unset, rejects the rest', () => {
    expect(parseShowMarkup(true)).toBe(true);
    expect(parseShowMarkup('Yes')).toBe(true);
    expect(parseShowMarkup(' off ')).toBe(false);
    expect(parseShowMarkup('')).toBeUndefined();
    expect(parseShowMarkup(undefined)).toBeUndefined();
    expect(() => parseShowMarkup('maybe', 'flag')).toThrow(/showMarkup must be true or false/);
    expect(() => parseShowMarkup(2)).toThrow(/E_VALIDATION/);
  });

  it('an explicit value, including false, beats the environment; the default is off', () => {
    expect(effectiveShowMarkup(undefined, {})).toBe(false);
    expect(effectiveShowMarkup(undefined, { [SHOW_MARKUP_ENV]: '1' })).toBe(true);
    expect(effectiveShowMarkup(false, { [SHOW_MARKUP_ENV]: '1' })).toBe(false);
    expect(effectiveShowMarkup(true, { [SHOW_MARKUP_ENV]: '0' })).toBe(true);
    expect(() => effectiveShowMarkup(undefined, { [SHOW_MARKUP_ENV]: 'bogus' })).toThrow(/ASTRA_SHOW_MARKUP showMarkup/);
  });
});

describe('LibreOffice filter spec', () => {
  it('is unchanged with no markup argument (the pinned forms)', () => {
    expect(pdfFilterSpec(true)).toBe('pdf:writer_pdf_Export');
    expect(pdfFilterSpec(false)).toBe('pdf:writer_pdf_Export:{"UseTaggedPDF":{"type":"boolean","value":"false"}}');
    expect(pdfFilterSpec(true, {})).toBe('pdf:writer_pdf_Export');
  });

  it('covers the four tagged x markup combinations with JSON that parses and names each key once', () => {
    for (const tagged of [true, false]) {
      const spec = pdfFilterSpec(tagged, { markup: true });
      const json = JSON.parse(spec.slice('pdf:writer_pdf_Export:'.length));
      expect(json.ExportNotesInMargin).toEqual({ type: 'boolean', value: 'true' });
      expect(spec.match(/ExportNotesInMargin/g)).toHaveLength(1);
      expect(Boolean(json.UseTaggedPDF)).toBe(!tagged);
    }
    expect(libreOfficeFilterSpec(true, false)).toBe('pdf:writer_pdf_Export');
  });
});

describe('LibreOffice version', () => {
  it('parses the launcher output and ignores anything else', () => {
    expect(parseLibreOfficeVersion('LibreOffice 26.8.0.3 bce0998a')).toMatchObject({ major: 26, minor: 8, micro: 0, text: '26.8.0.3' });
    expect(parseLibreOfficeVersion('LibreOffice 25.2.7')).toMatchObject({ text: '25.2.7' });
    expect(parseLibreOfficeVersion('Warning: something\nLibreOffice 24.8.1.2 x')).toMatchObject({ text: '24.8.1.2' });
    expect(parseLibreOfficeVersion('')).toBeNull();
    expect(parseLibreOfficeVersion('OpenOffice 4.1.2')).toBeNull();
    expect(parseLibreOfficeVersion('not LibreOffice 25.2.7')).toBeNull();
  });

  it('compares by major, minor, micro', () => {
    const v = (t: string) => parseLibreOfficeVersion(`LibreOffice ${t}`)!;
    expect(compareVersion(v('25.2.7'), v('25.2.7'))).toBe(0);
    expect(compareVersion(v('25.2.6'), v('25.2.7'))).toBeLessThan(0);
    expect(compareVersion(v('26.0.0'), v('25.99.99'))).toBeGreaterThan(0);
    expect(compareVersion(v('9.9.9'), v('10.0.0'))).toBeLessThan(0);
  });

  describe('which file is asked (versionProbeFile)', () => {
    const win = 'win32' as const;
    it('Windows soffice.exe asks the soffice.com beside it, and is unreadable without one', () => {
      expect(versionProbeFile('C:\\Program Files\\LibreOffice\\program\\soffice.exe', win, () => true)).toBe('C:\\Program Files\\LibreOffice\\program\\soffice.com');
      expect(versionProbeFile('C:\\Program Files\\LibreOffice\\program\\SOFFICE.EXE', win, () => true)).toBe('C:\\Program Files\\LibreOffice\\program\\soffice.com');
      expect(versionProbeFile('C:\\Program Files\\LibreOffice\\program\\soffice.exe', win, () => false)).toBeNull();
    });
    it('a configured path in another folder is probed in that folder, never the installed one', () => {
      const asked: string[] = [];
      const file = versionProbeFile('D:\\portable\\lo\\program\\soffice.exe', win, (p) => (asked.push(p), true));
      expect(file).toBe('D:\\portable\\lo\\program\\soffice.com');
      expect(asked).toEqual(['D:\\portable\\lo\\program\\soffice.com']);
    });
    it('a wrapper, a shim, a renamed binary and an already-.com path run as they are', () => {
      for (const bin of ['C:\\tools\\soffice-wrapper.cmd', 'C:\\tools\\lo.exe', 'C:\\Program Files\\LibreOffice\\program\\soffice.com']) {
        expect(versionProbeFile(bin, win, () => false)).toBe(bin);
      }
    });
    it('on Linux and macOS the selected file (a link, a shim) runs as it is', () => {
      for (const platform of ['linux', 'darwin'] as const) {
        expect(versionProbeFile('/usr/local/bin/lo', platform, () => false)).toBe('/usr/local/bin/lo');
        expect(versionProbeFile('/opt/LibreOffice/program/soffice', platform, () => false)).toBe('/opt/LibreOffice/program/soffice');
      }
    });
  });

  describe('the probe, with an injected runner', () => {
    const ok = (stdout: string): ProbeRun => ({ status: 0, stdout, timedOut: false });
    it('reads a version, passes --version, and caches the answer per path', () => {
      let calls = 0;
      let seen: { file: string; args: string[] } | undefined;
      const run = (file: string, args: string[]) => ((calls++, (seen = { file, args }), ok('LibreOffice 25.2.7.2 abc')));
      const a = probeLibreOfficeVersion('/x/soffice', { platform: 'linux', run });
      const b = probeLibreOfficeVersion('/x/soffice', { platform: 'linux', run });
      expect(a).toMatchObject({ ok: true });
      expect(b).toBe(a);
      expect(calls).toBe(1);
      expect(seen).toEqual({ file: '/x/soffice', args: ['--version'] });
      probeLibreOfficeVersion('/y/soffice', { platform: 'linux', run });
      expect(calls).toBe(2);
    });

    it('every unreadable case is unreadable, never a version', () => {
      const cases: Array<[string, ProbeRun]> = [
        ['empty', ok('')],
        ['garbage', ok('hello world')],
        ['wrong format', ok('LibreOffice vNext')],
        ['non-zero exit', { status: 3, stdout: 'LibreOffice 26.8.0.3 x', timedOut: false }],
        ['timeout', { status: null, stdout: '', timedOut: true }],
        ['cannot start', { status: null, stdout: '', timedOut: false, error: 'ENOENT' }],
      ];
      for (const [name, r] of cases) {
        resetLibreOfficeVersionCache();
        const p = probeLibreOfficeVersion('/x/soffice', { platform: 'linux', run: () => r });
        expect(p.ok, name).toBe(false);
      }
    });

    it('Windows soffice.exe without soffice.com runs nothing', () => {
      let calls = 0;
      const p = probeLibreOfficeVersion('C:\\LO\\program\\soffice.exe', { platform: 'win32', exists: () => false, run: () => (calls++, ok('LibreOffice 26.8.0.3 x')) });
      expect(p).toMatchObject({ ok: false });
      expect(calls).toBe(0);
    });

    it('a failed probe is retried after a minute; a readable one is kept for ten', () => {
      let t = 0;
      let calls = 0;
      let out = '';
      const run = () => (calls++, ok(out));
      const o = { platform: 'linux' as const, run, now: () => t };
      expect(probeLibreOfficeVersion('/x/soffice', o).ok).toBe(false);
      t = 30_000;
      probeLibreOfficeVersion('/x/soffice', o);
      expect(calls).toBe(1);
      t = 61_000;
      out = 'LibreOffice 26.8.0.3 x';
      expect(probeLibreOfficeVersion('/x/soffice', o).ok).toBe(true);
      expect(calls).toBe(2);
      t = 61_000 + 9 * 60_000;
      probeLibreOfficeVersion('/x/soffice', o);
      expect(calls).toBe(2);
      t = 61_000 + 11 * 60_000;
      probeLibreOfficeVersion('/x/soffice', o);
      expect(calls).toBe(3);
    });
  });

  describe('the guard messages', () => {
    const at = (text: string) => ({ platform: 'linux' as const, run: (): ProbeRun => ({ status: 0, stdout: text, timedOut: false }) });
    it('equal to and above the minimum pass', () => {
      expect(assertLibreOfficeMarkupSupport('/a', at(`LibreOffice ${MIN_LIBREOFFICE_FOR_MARKUP} x`)).text).toBe(MIN_LIBREOFFICE_FOR_MARKUP);
      resetLibreOfficeVersionCache();
      expect(assertLibreOfficeMarkupSupport('/a', at('LibreOffice 99.1.0.0 x')).major).toBe(99);
    });
    it('a version below the minimum is "has not been tested", E_DEPENDENCY, class request', () => {
      let err: unknown;
      try {
        assertLibreOfficeMarkupSupport('/a', at('LibreOffice 7.0.1.2 x'));
      } catch (e) {
        err = e;
      }
      expect(err).toMatchObject({ code: 'E_DEPENDENCY', failureClass: 'request' });
      expect((err as Error).message).toMatch(/LibreOffice 7\.0\.1\.2 has not been tested for showing comments in the margin/);
      expect((err as Error).message).toContain(MIN_LIBREOFFICE_FOR_MARKUP);
      expect((err as { remedy: string }).remedy).toMatch(/--renderer word/);
    });
    it('every LibreOffice markup render carries the long-comment warning, whatever the version', () => {
      expect(libreOfficeMarkupWarnings()).toEqual([LIBREOFFICE_LONG_COMMENT_WARNING]);
      expect(LIBREOFFICE_LONG_COMMENT_WARNING).toMatch(/cut short in the margin/);
      expect(LIBREOFFICE_LONG_COMMENT_WARNING).toMatch(/Word on Windows/);
    });
    it('an unreadable version says it could not be determined, and never says too old or cannot show', () => {
      let err: unknown;
      try {
        assertLibreOfficeMarkupSupport('/a', at('rubbish'));
      } catch (e) {
        err = e;
      }
      expect(err).toMatchObject({ code: 'E_DEPENDENCY' });
      const m = (err as Error).message;
      expect(m).toMatch(/could not determine whether this LibreOffice supports --show-markup/);
      expect(m).not.toMatch(/too old|cannot show|not support/i);
    });
  });
});

describe.skipIf(!POSIX)('the probe with real stand-in executables', () => {
  const probe = (bin: string, timeoutMs?: number) => probeLibreOfficeVersion(bin, { platform: 'linux', ...(timeoutMs ? { timeoutMs } : {}) });

  it('reads the fake soffice; empty output, stderr-only output, a non-zero exit and a hang are unreadable', () => {
    const bin = makeFakeSoffice(tmpBase);
    expect(probe(bin)).toMatchObject({ ok: true, version: { text: '26.8.0.3' } });
    for (const env of [{ FAKE_SOFFICE_VERSION: '' }, { FAKE_SOFFICE_VERSION_STREAM: 'stderr' }, { FAKE_SOFFICE_VERSION_EXIT: '2' }, { FAKE_SOFFICE_VERSION: 'junk' }]) {
      resetLibreOfficeVersionCache();
      const keys = Object.keys(env) as Array<keyof typeof env>;
      for (const k of keys) process.env[k] = env[k] as string;
      try {
        expect(probe(bin).ok, JSON.stringify(env)).toBe(false);
      } finally {
        for (const k of keys) delete process.env[k];
      }
    }
    resetLibreOfficeVersionCache();
    process.env.FAKE_SOFFICE_VERSION_HANG = '1';
    expect(probe(bin, 400)).toMatchObject({ ok: false, reason: expect.stringMatching(/timed out/) });
  });

  it('a wrapper that passes --version through works; one that swallows it is unreadable; a missing file is unreadable', () => {
    const real = makeFakeSoffice(tmpBase);
    const pass = path.join(tmpBase, 'lo-wrapper');
    fs.writeFileSync(pass, `#!/bin/sh\nexec "${real}" "$@"\n`, { mode: 0o755 });
    expect(probe(pass)).toMatchObject({ ok: true });
    const swallow = path.join(tmpBase, 'lo-quiet');
    fs.writeFileSync(swallow, '#!/bin/sh\nexit 0\n', { mode: 0o755 });
    expect(probe(swallow).ok).toBe(false);
    expect(probe(path.join(tmpBase, 'missing')).ok).toBe(false);
  });

  it('a symbolic link to the binary is probed as the link', () => {
    const real = makeFakeSoffice(tmpBase);
    const link = path.join(tmpBase, 'soffice-link');
    fs.symlinkSync(real, link);
    expect(probe(link)).toMatchObject({ ok: true });
  });
});

describe('the failure message', () => {
  it('plain run succeeded: states what happened in this call, keeps the real error, no diagnosis', () => {
    const e = markupFailureError('LibreOffice exited with code 81', { ok: true });
    expect(e.message).toMatch(/failed to render this document with comments in the margin, and rendered it without/);
    expect(e.message).toContain('exited with code 81');
    expect(e.failureClass).toBe('request');
    expect(e.code).toBe('E_RENDER');
  });
  it('both failed: both errors are kept and no cause is claimed', () => {
    const e = markupFailureError('error one', { ok: false, error: 'error two' });
    expect(e.message).toContain('error one');
    expect(e.message).toContain('error two');
    expect(e.message).toMatch(/the cause was not determined/);
    expect(e.message).not.toMatch(/option/i);
  });
});

describe('the result block and the notes', () => {
  it('names the engine and the mode it was told to use, and carries no applied claim', () => {
    expect(buildMarkupReport('word')).toEqual({ requested: true, renderer: 'word', mode: 'export-with-markup', warnings: [] });
    expect(buildMarkupReport('libreoffice', ['w'])).toEqual({ requested: true, renderer: 'libreoffice', mode: 'comments-in-margin', warnings: ['w'] });
    expect(JSON.stringify(buildMarkupReport('word'))).not.toMatch(/applied/);
  });
  it('the notes carry the contract sentence, the privacy sentence, then the warnings', () => {
    const n = markupNotes(buildMarkupReport('word', ['Word setting X could not be set.']));
    expect(n[0]).toBe(MARKUP_CONTRACT_NOTE);
    expect(n[1]).toBe(MARKUP_PRIVACY_NOTE);
    expect(n[2]).toMatch(/Word setting X/);
    expect(MARKUP_PRIVACY_NOTE).toMatch(/reviewer names, deleted text and comments/);
    expect(MARKUP_CONTRACT_NOTE).toMatch(/page numbers refer to this render/);
  });
});

describe('the Word script', () => {
  it('has no starting-state step and no environment path to one in the production script', () => {
    const s = buildWordScript();
    expect(s).not.toMatch(/STARTSTATE|ENDSTATE/);
    expect(s).not.toMatch(/ASTRA_WORD_TEST/);
    expect(s).not.toMatch(/START_STATE/);
  });
  it('pins the view only when markup is requested, with Item 7, and restores before the document is closed', () => {
    const s = buildWordScript();
    expect(s).toContain("$markup = ($env:ASTRA_WORD_MARKUP -eq '1')");
    expect(s).toContain('$item = 0');
    expect(s).toContain('$item = 7');
    expect(s).toContain('ExportAsFixedFormat($out, 17, $false, 0, 0, 1, 1, $item,');
    expect(s.indexOf('VIEWRESTORE')).toBeGreaterThan(-1);
    expect(s.indexOf('VIEWRESTORE')).toBeLessThan(s.indexOf('$doc.Close(0)'));
    expect(s).toContain('MarkupMode = 2');
    expect(s).toContain('Fail23');
  });
  it('reads no reviewer name and prints none', () => {
    const s = buildWordScript();
    expect(s).not.toMatch(/\.Author/);
    expect(s).not.toMatch(/Reviewers\.Item\(\$?[a-z]*name/i);
  });
  it('the starting-state step is injected only from a validated state, as numbers and booleans', () => {
    const s = buildWordScript({ markupMode: 0, showComments: false, hideFirstReviewer: true, noMarkup: true });
    expect(s).toContain('$view.ShowRevisionsAndComments = $false');
    expect(s).toContain('$view.RevisionsFilter.Markup = 0');
    expect(s).toContain('$view.MarkupMode = 0');
    expect(s).toContain('$view.ShowComments = $false');
    expect(s).toContain('Reviewers.Item(1).Visible = $false');
    expect(s).toContain('STARTSTATE');
    // The starting state is applied before the originals are recorded, so the restore returns to it. The end state is
    // printed after the restore.
    expect(s.indexOf('STARTSTATE')).toBeGreaterThan(s.indexOf('$view = $doc.ActiveWindow.View'));
    expect(s.indexOf('STARTSTATE')).toBeLessThan(s.indexOf('$viewOrig = @{'));
    expect(s.indexOf('ENDSTATE')).toBeGreaterThan(s.indexOf('VIEWRESTORE'));
    expect(s.indexOf('ENDSTATE')).toBeLessThan(s.indexOf('$doc.Close(0)'));
  });
  it('validation refuses anything outside the three keys and their values', () => {
    expect(validateStartState({ markupMode: 1 })).toEqual({ markupMode: 1 });
    expect(validateStartState({ noMarkup: true, hideFirstReviewer: false })).toEqual({ noMarkup: true, hideFirstReviewer: false });
    for (const bad of [{ noMarkup: 'yes' }, { markupMode: 3 }, { markupMode: '2' }, { showComments: 'false' }, { evil: true }, { hideFirstReviewer: 1 }, null, [], 'x']) {
      expect(() => validateStartState(bad as never), JSON.stringify(bad)).toThrow(/startState/);
    }
    expect(() => buildWordScript({ markupMode: '0; calc' } as never)).toThrow();
  });
});

describe('the Word protocol', () => {
  it('reads the markup, warning, restore and start-state lines, and nothing else from them', () => {
    const p = parseWordProtocol(['MARKUP 2', 'WARN RevisionsFilter.Markup', 'WARN RevisionsView', 'VIEWRESTORE failed MarkupMode,reviewer visibility', 'STARTSTATE mode=0 comments=0 hidden=1', 'ENDSTATE mode=0 comments=0 hidden=1', 'OK']);
    expect(p).toMatchObject({ markupMode: 2, warnings: ['RevisionsFilter.Markup', 'RevisionsView'], startState: 'mode=0 comments=0 hidden=1', endState: 'mode=0 comments=0 hidden=1', ok: true });
    expect(p.restore).toEqual({ ok: false, failed: ['MarkupMode', 'reviewer visibility'] });
    expect(parseWordProtocol(['VIEWRESTORE ok']).restore).toEqual({ ok: true, failed: [] });
    expect(parseWordProtocol(['WARN has spaces and a name']).warnings).toEqual([]);
    expect(parseWordProtocol(['STARTSTATE Alice=1']).startState).toBeUndefined();
  });
});

describe('Word through the fake PowerShell host', () => {
  const run = (o: Parameters<typeof wordDocxToPdf>[2]) => {
    const inp = path.join(tmpBase, 'in.docx');
    fs.writeFileSync(inp, blank());
    return wordDocxToPdf(inp, path.join(tmpBase, 'out.pdf'), { ...fakePsHooks, retryPauseMs: [0, 0], ...o });
  };

  it('without markup the variable is passed as 0 and no markup block is reported', async () => {
    process.env.FAKE_PS_MODE = 'ok';
    process.env.FAKE_PS_LOG = path.join(tmpBase, 'ps');
    const r = await run({});
    expect(r.markup).toBeUndefined();
    expect(logRows(process.env.FAKE_PS_LOG)[0].markup).toBe('0');
  });

  it('an inherited ASTRA_WORD_MARKUP=1 cannot turn the review view on when markup is off', async () => {
    process.env.FAKE_PS_MODE = 'ok';
    process.env.FAKE_PS_LOG = path.join(tmpBase, 'ps');
    process.env.ASTRA_WORD_MARKUP = '1';
    try {
      for (const o of [{}, { markup: false }]) {
        const r = await run(o);
        expect(r.markup).toBeUndefined();
      }
      const rows = logRows(process.env.FAKE_PS_LOG);
      expect(rows.map((x) => x.markup)).toEqual(['0', '0']);
    } finally {
      delete process.env.ASTRA_WORD_MARKUP;
    }
  });

  it('with markup the variable is passed and the confirmed mode comes back', async () => {
    process.env.FAKE_PS_MODE = 'ok';
    process.env.FAKE_PS_LOG = path.join(tmpBase, 'ps');
    const r = await run({ markup: true });
    expect(logRows(process.env.FAKE_PS_LOG)[0].markup).toBe('1');
    expect(r.markup).toMatchObject({ mode: 2, warnings: [] });
  });

  it('best-effort settings and a failed restore become warnings; the render still succeeds', async () => {
    process.env.FAKE_PS_MODE = 'ok';
    process.env.FAKE_PS_WARN = 'RevisionsFilter.Markup,RevisionsView';
    process.env.FAKE_PS_RESTORE_FAIL = 'MarkupMode';
    const r = await run({ markup: true });
    expect(r.markup?.warnings).toHaveLength(3);
    expect(r.markup?.warnings[0]).toMatch(/RevisionsFilter\.Markup could not be set/);
    expect(r.markup?.warnings[2]).toMatch(/could not be fully restored \(MarkupMode\)/);
    expect(r.markup?.restore).toEqual({ ok: false, failed: ['MarkupMode'] });
  });

  it('exit 23 is E_RENDER with class request, the setting named, and a remedy', async () => {
    process.env.FAKE_PS_MODE = 'markup-view-fail';
    const err = (await run({ markup: true }).catch((e: unknown) => e)) as WordRenderError;
    expect(err).toBeInstanceOf(WordRenderError);
    expect(err).toMatchObject({ code: 'E_RENDER', failureClass: 'request', exitCode: 23 });
    expect(err.message).toMatch(/could not be set to show comments and tracked changes \(MarkupMode/);
    expect(err.remedy).toMatch(/without --show-markup|--renderer libreoffice/);
    expect(fs.existsSync(path.join(tmpBase, 'out.pdf'))).toBe(false);
  });

  it('a success with no confirmation of the view is refused as a markup render', async () => {
    process.env.FAKE_PS_MODE = 'markup-no-confirm';
    await expect(run({ markup: true })).rejects.toMatchObject({ code: 'E_RENDER', failureClass: 'request', message: expect.stringMatching(/did not confirm the review view/) });
    await expect(run({})).resolves.toBeDefined();
  });

  it('the starting state is ignored without markup, and an invalid one is refused before anything runs', async () => {
    process.env.FAKE_PS_MODE = 'ok';
    process.env.FAKE_PS_LOG = path.join(tmpBase, 'ps');
    await run({ startState: { markupMode: 0 } });
    await expect(run({ markup: true, startState: { markupMode: 7 } as never })).rejects.toThrow(/startState/);
  });
});

describe('the fallback rules, through the real runPlan', () => {
  class Req extends Error {
    failureClass = 'request' as const;
  }
  class Infra extends Error {
    failureClass = 'infrastructure' as const;
  }

  it('the request class is a known class', () => {
    expect(failureClassOf(new Req('x'))).toBe('request');
  });

  it('an explicit renderer fails without switching engine', async () => {
    const asked: string[] = [];
    const wordBreaker = new CircuitBreaker(1000);
    await expect(
      runPlan(['word'] as const, { canFallBack: false, auto: false }, async (e) => {
        asked.push(e);
        throw new Req('no markup view');
      }, { word: wordBreaker })
    ).rejects.toThrow('no markup view');
    expect(asked).toEqual(['word']);
    expect(wordBreaker.isOpen()).toBe(false);
  });

  it('auto falls through on a request failure, keeps the breaker closed, and names the engine used and why', async () => {
    const wordBreaker = new CircuitBreaker(1000);
    const seen: Array<{ engine: string; markup: boolean }> = [];
    const out = await runPlan(['word', 'libreoffice'] as const, { canFallBack: true, auto: true }, async (engine) => {
      seen.push({ engine, markup: true });
      if (engine === 'word') throw new Req('E_RENDER: Word could not be set to show comments and tracked changes (MarkupMode).');
      return 'pdf';
    }, { word: wordBreaker });
    expect(out).toMatchObject({ used: 'libreoffice', fallbackFrom: 'word' });
    expect(out.fallbackReason).toMatch(/could not be set to show comments/);
    expect(seen.every((s) => s.markup)).toBe(true);
    expect(wordBreaker.isOpen()).toBe(false);
  });

  it('an infrastructure failure under auto still opens the breaker', async () => {
    const wordBreaker = new CircuitBreaker(1000);
    await runPlan(['word', 'libreoffice'] as const, { canFallBack: true, auto: true }, async (e) => {
      if (e === 'word') throw new Infra('Word stuck');
      return 1;
    }, { word: wordBreaker });
    expect(wordBreaker.isOpen()).toBe(true);
  });

  it('when both engines fail the last error is thrown with the earlier one named', async () => {
    await expect(
      runPlan(['word', 'libreoffice'] as const, { canFallBack: true, auto: true }, async (e) => {
        throw new Req(`${e} failed`);
      })
    ).rejects.toThrow(/libreoffice failed.*word failed first: word failed/);
  });
});

describe.skipIf(!POSIX)('the visual render path with stand-in programs', () => {
  let soffice: string;
  let pdftoppm: string;
  beforeEach(() => {
    soffice = makeFakeSoffice(tmpBase);
    pdftoppm = makeFakePdftoppm(tmpBase);
    process.env.FAKE_LOG = path.join(tmpBase, 'log');
  });
  const base = () => ({ sofficeBin: soffice, pdftoppmBin: pdftoppm });
  const converts = () => logRows(process.env.FAKE_LOG as string).filter((r) => r.bin === 'soffice' && r.filter !== undefined).map((r) => r.filter as string);

  it('off: the literal pdf filter, no version query, no markup block (the default command line)', async () => {
    const r = await renderDocxToImages(blank(), base());
    expect(converts()).toEqual(['pdf']);
    expect(logRows(process.env.FAKE_LOG as string).some((x) => x.version)).toBe(false);
    expect('markup' in r).toBe(false);
  });

  it('on: asks the version, converts with the margin-comments filter, reports what was asked', async () => {
    const r = await renderDocxToImages(blank(), { ...base(), showMarkup: true });
    expect(converts()).toHaveLength(1);
    expect(converts()[0]).toContain('"ExportNotesInMargin"');
    expect(r.markup).toEqual({ requested: true, renderer: 'libreoffice', mode: 'comments-in-margin', warnings: [LIBREOFFICE_LONG_COMMENT_WARNING] });
  });

  it('a LibreOffice markup render warns that long comments can be cut short, on an old and on a current version', async () => {
    process.env.FAKE_SOFFICE_VERSION = 'LibreOffice 25.2.7.2 x';
    const old = await renderDocxToImages(blank(), { ...base(), showMarkup: true });
    expect(old.markup?.warnings).toEqual([LIBREOFFICE_LONG_COMMENT_WARNING]);
    resetLibreOfficeVersionCache();
    delete process.env.FAKE_SOFFICE_VERSION;
    const current = await renderDocxToImages(blank(), { ...base(), showMarkup: true });
    expect(current.markup?.warnings).toEqual([LIBREOFFICE_LONG_COMMENT_WARNING]);
  });

  it('the environment variable turns it on, and an explicit false beats it', async () => {
    process.env[SHOW_MARKUP_ENV] = '1';
    const on = await renderDocxToImages(blank(), base());
    expect(on.markup?.requested).toBe(true);
    const off = await renderDocxToImages(blank(), { ...base(), showMarkup: false });
    expect(off.markup).toBeUndefined();
  });

  it('an invalid environment value fails loudly before any program runs', async () => {
    process.env[SHOW_MARKUP_ENV] = 'sometimes';
    await expect(renderDocxToImages(blank(), base())).rejects.toMatchObject({ code: 'E_VALIDATION' });
    expect(logRows(process.env.FAKE_LOG as string)).toEqual([]);
  });

  it('a version below the minimum is E_DEPENDENCY and nothing is converted', async () => {
    process.env.FAKE_SOFFICE_VERSION = 'LibreOffice 7.1.0.3 x';
    await expect(renderDocxToImages(blank(), { ...base(), showMarkup: true })).rejects.toMatchObject({ code: 'E_DEPENDENCY', message: expect.stringMatching(/has not been tested/) });
    expect(converts()).toEqual([]);
  });

  it('an unreadable version is "could not determine" and nothing is converted', async () => {
    process.env.FAKE_SOFFICE_VERSION = '';
    await expect(renderDocxToImages(blank(), { ...base(), showMarkup: true })).rejects.toMatchObject({ code: 'E_DEPENDENCY', message: expect.stringMatching(/could not determine/) });
    expect(converts()).toEqual([]);
  });

  it('a failure with the option runs one diagnostic without it, reports it, and never returns the diagnostic output', async () => {
    process.env.FAKE_SOFFICE_FAIL_MARKUP = '1';
    const err = (await renderDocxToImages(blank(), { ...base(), showMarkup: true }).catch((e: unknown) => e)) as Error & { code: string };
    expect(err).toMatchObject({ code: 'E_RENDER' });
    expect(err.message).toMatch(/rendered it without/);
    expect(err.message).toContain('fake soffice markup failure');
    const f = converts();
    expect(f.filter((x) => x.includes('ExportNotesInMargin'))).toHaveLength(3); // the existing 3 attempts
    expect(f.filter((x) => x === 'pdf')).toHaveLength(1); // one diagnostic run
  });

  it('a failure both ways keeps both errors and claims no cause', async () => {
    process.env.FAKE_SOFFICE_FAIL = '1';
    const err = (await renderDocxToImages(blank(), { ...base(), showMarkup: true }).catch((e: unknown) => e)) as Error;
    expect(err.message).toMatch(/also without/);
    expect(err.message).toMatch(/the cause was not determined/);
  });

  it('word with markup: passes the request, reports the engine and warnings; pages come back', async () => {
    process.env.FAKE_PS_MODE = 'ok';
    process.env.FAKE_PS_LOG = path.join(tmpBase, 'ps');
    process.env.FAKE_PS_WARN = 'RevisionsView';
    const r = await renderDocxToImages(blank(), { renderer: 'word', pdfToImage: 'poppler', pdftoppmBin: pdftoppm, showMarkup: true, testHooks: winHooks });
    expect(logRows(process.env.FAKE_PS_LOG)[0].markup).toBe('1');
    expect(r.markup).toMatchObject({ requested: true, renderer: 'word', mode: 'export-with-markup' });
    expect(r.markup?.warnings[0]).toMatch(/RevisionsView/);
  });

  it('explicit word with a view failure: fails, never switches to LibreOffice, no breaker', async () => {
    process.env.FAKE_PS_MODE = 'markup-view-fail';
    await expect(renderDocxToImages(blank(), { renderer: 'word', pdfToImage: 'poppler', ...base(), showMarkup: true, testHooks: winHooks })).rejects.toMatchObject({ code: 'E_RENDER', failureClass: 'request' });
    expect(converts()).toEqual([]);
  });

  it('auto with Word first: a view failure falls through to LibreOffice, which is also asked for markup, and the result says so', async () => {
    process.env.FAKE_PS_MODE = 'markup-view-fail';
    const hooks = { ...winHooks, autoPrefersWord: true, windowsRenderer: false };
    const r = await renderDocxToImages(blank(), { ...base(), showMarkup: true, testHooks: hooks });
    expect(r.renderer).toMatchObject({ requested: 'auto', used: 'libreoffice', fallbackFrom: 'word' });
    expect(r.renderer?.fallbackReason).toMatch(/could not be set to show comments/);
    expect(converts()[0]).toContain('ExportNotesInMargin');
    expect(r.markup).toMatchObject({ renderer: 'libreoffice', mode: 'comments-in-margin' });
    // The Word breaker stays closed: a second call still tries Word first.
    process.env.FAKE_PS_LOG = path.join(tmpBase, 'ps');
    await renderDocxToImages(blank(), { ...base(), showMarkup: true, testHooks: hooks });
    expect(logRows(process.env.FAKE_PS_LOG)).toHaveLength(1);
  });

  it('auto with Word first: a plain render after an infrastructure failure still uses the breaker (unchanged)', async () => {
    process.env.FAKE_PS_MODE = 'no-word';
    const hooks = { ...winHooks, autoPrefersWord: true, windowsRenderer: false };
    await renderDocxToImages(blank(), { ...base(), testHooks: hooks });
    process.env.FAKE_PS_LOG = path.join(tmpBase, 'ps');
    await renderDocxToImages(blank(), { ...base(), testHooks: hooks });
    expect(logRows(process.env.FAKE_PS_LOG)).toEqual([]);
  });
});

describe.skipIf(!POSIX)('the PDF export path with stand-in programs', () => {
  let soffice: string;
  beforeEach(() => {
    soffice = makeFakeSoffice(tmpBase);
    process.env.FAKE_LOG = path.join(tmpBase, 'log');
  });
  const out = () => path.join(tmpBase, 'out.pdf');
  const converts = () => logRows(process.env.FAKE_LOG as string).filter((r) => r.bin === 'soffice' && r.filter !== undefined).map((r) => r.filter as string);

  it('off: the pinned filter and no markup block', async () => {
    const r = await convertDocxToPdfDetailed(blank(), out(), { binary: soffice });
    expect(converts()).toEqual(['pdf:writer_pdf_Export']);
    expect(r.markup).toBeUndefined();
  });

  it('the export result carries the same warning', async () => {
    const r = await convertDocxToPdfDetailed(blank(), out(), { binary: soffice, showMarkup: true });
    expect(r.markup?.warnings).toEqual([LIBREOFFICE_LONG_COMMENT_WARNING]);
  });

  it('on, tagged and untagged: the margin-comments option is combined with the tagging choice', async () => {
    const t = await convertDocxToPdfDetailed(blank(), out(), { binary: soffice, showMarkup: true });
    expect(t.markup).toEqual({ requested: true, renderer: 'libreoffice', mode: 'comments-in-margin', warnings: [LIBREOFFICE_LONG_COMMENT_WARNING] });
    await convertDocxToPdfDetailed(blank(), out(), { binary: soffice, showMarkup: true, tagged: false });
    const [a, b] = converts();
    expect(a).toBe('pdf:writer_pdf_Export:{"ExportNotesInMargin":{"type":"boolean","value":"true"}}');
    expect(JSON.parse(b.slice('pdf:writer_pdf_Export:'.length))).toMatchObject({ UseTaggedPDF: { value: 'false' }, ExportNotesInMargin: { value: 'true' } });
    expect(fs.readFileSync(out()).subarray(0, 5).toString()).toBe('%PDF-');
  });

  it('a failure with the option is not retried silently without it: the error says what the diagnostic run showed, and no PDF is written', async () => {
    process.env.FAKE_SOFFICE_FAIL_MARKUP = '1';
    const err = (await convertDocxToPdfDetailed(blank(), out(), { binary: soffice, showMarkup: true }).catch((e: unknown) => e)) as Error;
    expect(err.message).toMatch(/rendered it without/);
    expect(fs.existsSync(out())).toBe(false);
    expect(converts().filter((x) => x === 'pdf:writer_pdf_Export')).toHaveLength(1);
  });

  it('an old version is refused before any conversion', async () => {
    process.env.FAKE_SOFFICE_VERSION = 'LibreOffice 6.4.7.2 x';
    await expect(convertDocxToPdfDetailed(blank(), out(), { binary: soffice, showMarkup: true })).rejects.toMatchObject({ code: 'E_DEPENDENCY' });
    expect(converts()).toEqual([]);
  });

  it('word with markup reports the engine; a view failure under an explicit word does not switch', async () => {
    process.env.FAKE_PS_MODE = 'ok';
    const ok = await convertDocxToPdfDetailed(blank(), out(), { renderer: 'word', showMarkup: true, testHooks: winHooks });
    expect(ok.markup).toMatchObject({ renderer: 'word', mode: 'export-with-markup' });
    process.env.FAKE_PS_MODE = 'markup-view-fail';
    await expect(convertDocxToPdfDetailed(blank(), out(), { renderer: 'word', binary: soffice, showMarkup: true, testHooks: winHooks })).rejects.toMatchObject({ failureClass: 'request' });
    expect(converts()).toEqual([]);
  });
});




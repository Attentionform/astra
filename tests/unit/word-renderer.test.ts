import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { MIN_WORD_VERSION, WordRenderError, buildWordScript, findWord, parseWordProtocol, wordDocxToPdf } from '../../packages/engine/src/visual/word';
import { MAX_ENCODED_COMMAND_CHARS, encodeCommand, runPowerShell, powershellPath, sanitiseDetail } from '../../packages/engine/src/visual/ps-runner';
import { fakePsHooks, withFakePsEnv } from '../helpers/fake-ps';

const dirs: string[] = [];
function tmp(): string {
  const d = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-word-test-'));
  dirs.push(d);
  return d;
}
afterEach(() => {
  while (dirs.length) fs.rmSync(dirs.pop() as string, { recursive: true, force: true });
});

const CLSID = '{000209FF-0000-0000-C000-000000000046}';
function reg(map: Record<string, string>) {
  return (key: string) => map[key] ?? null;
}
const goodReg = {
  'HKCR\\Word.Application\\CurVer': 'Word.Application.16',
  'HKCR\\Word.Application\\CLSID': CLSID,
  [`HKCR\\CLSID\\${CLSID}\\LocalServer32`]: '"C:\\Program Files\\Microsoft Office\\Root\\Office16\\WINWORD.EXE" /Automation',
};

describe('findWord', () => {
  it('finds a registered Word and strips the quotes and /Automation', () => {
    const w = findWord({ platform: 'win32', regQuery: reg(goodReg), exists: () => true });
    expect(w).toEqual({ exe: 'C:\\Program Files\\Microsoft Office\\Root\\Office16\\WINWORD.EXE', version: 16 });
  });
  it('handles an unquoted path with spaces', () => {
    const w = findWord({
      platform: 'win32',
      regQuery: reg({ ...goodReg, [`HKCR\\CLSID\\${CLSID}\\LocalServer32`]: 'C:\\Program Files (x86)\\Microsoft Office\\Office14\\WINWORD.EXE /Automation', 'HKCR\\Word.Application\\CurVer': 'Word.Application.14' }),
      exists: () => true,
    });
    expect(w?.exe).toBe('C:\\Program Files (x86)\\Microsoft Office\\Office14\\WINWORD.EXE');
    expect(w?.version).toBe(14);
  });
  it('falls back to the 32-bit registry view for a 32-bit Office', () => {
    const noDirect = { ...goodReg };
    delete (noDirect as Record<string, string>)[`HKCR\\CLSID\\${CLSID}\\LocalServer32`];
    const w = findWord({
      platform: 'win32',
      regQuery: reg({ ...noDirect, [`HKCR\\WOW6432Node\\CLSID\\${CLSID}\\LocalServer32`]: '"C:\\Program Files (x86)\\Microsoft Office\\Root\\Office16\\WINWORD.EXE" /Automation' }),
      exists: () => true,
    });
    expect(w?.exe).toContain('Program Files (x86)');
  });
  it('expands an environment variable in the registered path', () => {
    const w = findWord({
      platform: 'win32',
      env: { ProgramFiles: 'D:\\Apps' },
      regQuery: reg({ ...goodReg, [`HKCR\\CLSID\\${CLSID}\\LocalServer32`]: '"%ProgramFiles%\\Microsoft Office\\WINWORD.EXE" /Automation' }),
      exists: () => true,
    });
    expect(w?.exe).toBe('D:\\Apps\\Microsoft Office\\WINWORD.EXE');
  });
  it('a stale registration (executable gone) is not installed', () => {
    expect(findWord({ platform: 'win32', regQuery: reg(goodReg), exists: () => false })).toBeNull();
  });
  it('Office older than 2010 is not usable', () => {
    expect(MIN_WORD_VERSION).toBe(14);
    expect(findWord({ platform: 'win32', regQuery: reg({ ...goodReg, 'HKCR\\Word.Application\\CurVer': 'Word.Application.12' }), exists: () => true })).toBeNull();
  });
  it('nothing registered, garbage values, or a bad CLSID are all "not installed"', () => {
    expect(findWord({ platform: 'win32', regQuery: () => null, exists: () => true })).toBeNull();
    expect(findWord({ platform: 'win32', regQuery: reg({ ...goodReg, 'HKCR\\Word.Application\\CurVer': 'nonsense' }), exists: () => true })).toBeNull();
    expect(findWord({ platform: 'win32', regQuery: reg({ ...goodReg, 'HKCR\\Word.Application\\CLSID': 'not-a-clsid' }), exists: () => true })).toBeNull();
    expect(findWord({ platform: 'win32', regQuery: reg({ ...goodReg, [`HKCR\\CLSID\\${CLSID}\\LocalServer32`]: 'no executable here' }), exists: () => true })).toBeNull();
  });
  it('a registry or filesystem probe that throws is "not installed", never an exception', () => {
    expect(findWord({ platform: 'win32', regQuery: () => { throw new Error('boom'); } })).toBeNull();
    expect(findWord({ platform: 'win32', regQuery: reg(goodReg), exists: () => { throw new Error('boom'); } })).toBeNull();
  });
  it('never runs off Windows', () => {
    for (const platform of ['linux', 'darwin'] as const) {
      expect(findWord({ platform, regQuery: reg(goodReg), exists: () => true })).toBeNull();
    }
  });
});

describe('the Word script', () => {
  const script = buildWordScript();
  it('keeps every promise the plan makes about not disturbing the user', () => {
    expect(script).not.toMatch(/\.Visible\s*=/); // never made visible
    expect(script).toContain('$word.DisplayAlerts = 0');
    expect(script).toContain('$word.AutomationSecurity = 3'); // macros force-disabled
    expect(script).toMatch(/Documents\.Open\(\$in, \$false, \$true, \$false\)/); // no conversion prompt, read-only, not in recent files
    expect(script).not.toMatch(/\.Options\./); // no persistent user options
    expect(script).not.toMatch(/Documents\.Add|Stop-Process|taskkill|\.Kill\(|TerminateProcess/i);
    expect(script).not.toMatch(/ExecutionPolicy|Set-ItemProperty|New-ItemProperty|reg add/i);
    expect(script).toContain("LanguageMode -ne 'FullLanguage'");
    expect(script).toContain('$env:ASTRA_WORD_IN');
  });
  it('reads inputs from the environment only, so a file name cannot inject script text', () => {
    expect(script).not.toMatch(/\$in\s*=\s*['"]/);
  });
  it('quits only after the stacked checks', () => {
    const quit = script.slice(script.indexOf('$mayQuit = $false'));
    for (const guard of ['$stillAuto', '-not $mismatch', '$crossChecked -or $alone', '-not $word.Visible', '$word.UserControl -eq $false', '$word.Documents.Count -eq 0']) {
      expect(quit).toContain(guard);
    }
    expect(script.match(/\.Quit\(\)/g)?.length).toBe(1);
  });
  it('restores the two settings it changed when it does not quit', () => {
    expect(script).toContain('$word.DisplayAlerts = $origAlerts; $word.AutomationSecurity = $origSec');
  });
  it('encodes under the command-line limit and the encoding is UTF-16LE base64', () => {
    const enc = encodeCommand(script);
    expect(enc.length).toBeLessThan(MAX_ENCODED_COMMAND_CHARS);
    expect(Buffer.from(enc, 'base64').toString('utf16le')).toBe(script);
  });
  it('uses the absolute Windows PowerShell path', () => {
    expect(powershellPath({ SystemRoot: 'C:\\Windows' })).toBe('C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe');
    expect(powershellPath({})).toBe('C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe');
  });
});

describe('protocol parser', () => {
  it('reads a full success', () => {
    expect(parseWordProtocol(['CANDIDATE 12', 'VERSION 16.0', 'PAGES 3', 'OK'])).toEqual({ candidatePid: 12, version: '16.0', pages: 3, ok: true });
  });
  it('reads an error with its hex HRESULT and message', () => {
    expect(parseWordProtocol(['ERR 20 0x800a13e9 The password is wrong']).err).toEqual({ code: 20, hresult: '0x800A13E9', message: 'The password is wrong' });
  });
  it('ignores garbage, a truncated line and an empty message', () => {
    expect(parseWordProtocol(['hello', 'PAGES x', 'CANDIDATE', 'ERR 20', 'ERR 20 0x1']).ok).toBe(false);
    expect(parseWordProtocol(['ERR 13 0x00000000']).err?.message).toBe('');
  });
});

describe('sanitiseDetail', () => {
  it('strips control characters and bounds the length', () => {
    expect(sanitiseDetail('a\u0000b\r\n\tc')).toBe('a b c');
    expect(sanitiseDetail('x'.repeat(1000)).length).toBeLessThanOrEqual(301);
  });
});

async function render(mode: string, extra: Record<string, string | undefined> = {}, opts: Partial<Parameters<typeof wordDocxToPdf>[2]> = {}) {
  const d = tmp();
  const inDocx = path.join(d, 'input.docx');
  fs.writeFileSync(inDocx, 'x');
  const out = path.join(d, 'out.pdf');
  const counter = path.join(d, 'counter');
  return withFakePsEnv({ FAKE_PS_MODE: mode, FAKE_PS_COUNTER: counter, ...extra }, async () => ({
    out,
    result: await wordDocxToPdf(inDocx, out, { ...fakePsHooks, timeoutMs: 20000, retryPauseMs: [0, 0], ...opts }),
  }));
}
async function fail(mode: string, extra: Record<string, string | undefined> = {}, opts: Partial<Parameters<typeof wordDocxToPdf>[2]> = {}): Promise<WordRenderError> {
  try {
    await render(mode, extra, opts);
  } catch (e) {
    expect(e).toBeInstanceOf(WordRenderError);
    return e as WordRenderError;
  }
  throw new Error('expected a failure');
}

describe('wordDocxToPdf with a fake PowerShell', () => {
  it('renders and reports pages and version', async () => {
    const { out, result } = await render('ok');
    expect(result).toEqual({ pages: 1, version: '16.0' });
    expect(fs.readFileSync(out, 'latin1').startsWith('%PDF-')).toBe(true);
  });
  it('passes the tagged choice to the script through the environment', async () => {
    const log = path.join(tmp(), 'log');
    await render('ok', { FAKE_PS_LOG: log }, { tagged: false });
    await render('ok', { FAKE_PS_LOG: log }, { tagged: true });
    const rows = fs.readFileSync(log, 'utf8').trim().split('\n').map((l) => JSON.parse(l));
    expect(rows.map((r) => r.tagged)).toEqual(['0', '1']);
  });
  it('retries an instance that could not be confirmed private, then succeeds', async () => {
    const { result } = await render('busy-then-ok', { FAKE_PS_BUSY: '2' });
    expect(result.pages).toBeUndefined();
  });
  it('gives up after three attempts on a busy instance: infrastructure, nothing changed', async () => {
    const d = tmp();
    const counter = path.join(d, 'c');
    const e = await fail('always-busy', { FAKE_PS_COUNTER: counter });
    expect(fs.readFileSync(counter, 'utf8')).toBe('3');
    expect(e.failureClass).toBe('infrastructure');
    expect(e.code).toBe('E_RENDER');
    expect(e.message).toMatch(/Nothing was changed or closed/);
    expect(e.exitCode).toBe(13);
  });
  it('an ambiguous instance is refused the same way', async () => {
    const e = await fail('ambiguous');
    expect(e.exitCode).toBe(14);
    expect(e.failureClass).toBe('infrastructure');
  });
  it('a password-protected document is a document failure with a clear message', async () => {
    const e = await fail('password');
    expect(e.failureClass).toBe('document');
    expect(e.code).toBe('E_VALIDATION');
    expect(e.message).toMatch(/password-protected documents cannot be rendered/);
    expect(e.hresult).toBe('0x800A13E9');
  });
  it('a corrupt document is a document failure', async () => {
    const e = await fail('corrupt');
    expect(e.failureClass).toBe('document');
    expect(e.code).toBe('E_VALIDATION');
  });
  it('an export failure is a document failure only for a 0x800A HRESULT', async () => {
    expect((await fail('export-doc')).failureClass).toBe('document');
    expect((await fail('export-infra')).failureClass).toBe('infrastructure');
  });
  it('Word that cannot start and constrained language are infrastructure E_DEPENDENCY', async () => {
    for (const mode of ['no-word', 'constrained']) {
      const e = await fail(mode);
      expect(e.failureClass).toBe('infrastructure');
      expect(e.code).toBe('E_DEPENDENCY');
    }
  });
  it('success without a PDF, or with a file that is not a PDF, is a failure', async () => {
    expect((await fail('ok-no-pdf')).message).toMatch(/no valid PDF/);
    expect((await fail('ok-bad-pdf')).message).toMatch(/no valid PDF/);
  });
  it('garbage output, a crash and an unknown exit code are infrastructure failures with bounded detail', async () => {
    for (const mode of ['garbage', 'crash', 'odd-exit']) {
      const e = await fail(mode);
      expect(e.failureClass).toBe('infrastructure');
      expect(e.message.length).toBeLessThan(600);
    }
    expect((await fail('crash')).message).toMatch(/boom from powershell/);
  });
  it('a timeout names the hidden Word it may have left and never claims to have stopped it', async () => {
    const e = await fail('hang', {}, { timeoutMs: 1500 });
    expect(e.failureClass).toBe('infrastructure');
    expect(e.message).toMatch(/did not respond/);
    expect(e.message).toMatch(/pid 4242/);
    expect(e.message).toMatch(/did not stop it/);
    expect(e.candidatePid).toBe(4242);
  });
  it('PowerShell that cannot be started is E_DEPENDENCY infrastructure', async () => {
    const d = tmp();
    fs.writeFileSync(path.join(d, 'input.docx'), 'x');
    await expect(wordDocxToPdf(path.join(d, 'input.docx'), path.join(d, 'o.pdf'), { powershellBin: path.join(d, 'no-such-powershell'), powershellArgs: [], timeoutMs: 5000 })).rejects.toMatchObject({
      code: 'E_DEPENDENCY',
      failureClass: 'infrastructure',
    });
  });
  it('a path that is too long for Word is refused before anything starts', async () => {
    await expect(wordDocxToPdf('C:\\' + 'a'.repeat(300) + '.docx', 'C:\\o.pdf', fakePsHooks)).rejects.toMatchObject({ failureClass: 'infrastructure' });
  });
  it('renders one at a time', async () => {
    const d = tmp();
    const log = path.join(d, 'log');
    const inDocx = path.join(d, 'input.docx');
    fs.writeFileSync(inDocx, 'x');
    await withFakePsEnv({ FAKE_PS_MODE: 'ok', FAKE_PS_LOG: log }, async () => {
      await Promise.all([1, 2, 3].map((i) => wordDocxToPdf(inDocx, path.join(d, `o${i}.pdf`), { ...fakePsHooks, timeoutMs: 20000 })));
    });
    expect(fs.readFileSync(log, 'utf8').trim().split('\n')).toHaveLength(3);
  });
});

describe('runPowerShell', () => {
  it('passes inputs through the environment and delivers lines in order', async () => {
    const lines: string[] = [];
    const res = await withFakePsEnv({ FAKE_PS_MODE: 'ok' }, () =>
      runPowerShell({ script: 'unused', env: { ASTRA_WORD_IN: 'a', ASTRA_WORD_OUT: path.join(tmp(), 'o.pdf') }, timeoutMs: 10000, ...fakePsHooks, onLine: (l) => lines.push(l) })
    );
    expect(res.exitCode).toBe(0);
    expect(lines).toEqual(res.lines);
    expect(lines[0]).toBe('CANDIDATE 4321');
  });
  it('refuses a script that would not fit on a command line', async () => {
    const res = await runPowerShell({ script: 'x'.repeat(60000), env: {}, timeoutMs: 5000 });
    expect(res.spawnError?.code).toBe('E2BIG');
  });
});

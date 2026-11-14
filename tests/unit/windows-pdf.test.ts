import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { WINDOWS_PDF_MAX_WIDTH, WindowsPdfError, buildWindowsPdfScript, parseWindowsPdfProtocol, windowsPdfToImages, windowsRendererAvailable } from '../../packages/engine/src/visual/windows-pdf';
import { MAX_ENCODED_COMMAND_CHARS, encodeCommand } from '../../packages/engine/src/visual/ps-runner';
import { fakePsHooks, withFakePsEnv } from '../helpers/fake-ps';

const dirs: string[] = [];
const tmp = () => {
  const d = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-winpdf-test-'));
  dirs.push(d);
  return d;
};
afterEach(() => {
  while (dirs.length) fs.rmSync(dirs.pop() as string, { recursive: true, force: true });
});

describe('windowsRendererAvailable', () => {
  it('needs Windows 10 or later and Windows PowerShell', () => {
    expect(windowsRendererAvailable({ platform: 'win32', release: '10.0.22631', exists: () => true })).toBe(true);
    expect(windowsRendererAvailable({ platform: 'win32', release: '11.0.0', exists: () => true })).toBe(true);
    expect(windowsRendererAvailable({ platform: 'win32', release: '6.3.9600', exists: () => true })).toBe(false); // 8.1
    expect(windowsRendererAvailable({ platform: 'win32', release: '10.0.22631', exists: () => false })).toBe(false);
    expect(windowsRendererAvailable({ platform: 'win32', release: 'garbage', exists: () => true })).toBe(false);
    for (const platform of ['linux', 'darwin'] as const) expect(windowsRendererAvailable({ platform, release: '10.0.1', exists: () => true })).toBe(false);
  });
});

describe('the Windows PDF script', () => {
  const script = buildWindowsPdfScript();
  it('reads inputs from the environment and writes only page images into its folder', () => {
    for (const v of ['ASTRA_PDF_IN', 'ASTRA_IMG_DIR', 'ASTRA_IMG_FORMAT', 'ASTRA_IMG_DPI', 'ASTRA_IMG_QUALITY', 'ASTRA_IMG_MAX', 'ASTRA_IMG_MAXWIDTH', 'ASTRA_IMG_PAGES']) expect(script).toContain(`$env:${v}`);
    expect(script).not.toMatch(/ExecutionPolicy|Invoke-Expression|iex |Start-Process|DownloadFile|Invoke-WebRequest/i);
    expect(script).toContain("LanguageMode -ne 'FullLanguage'");
  });
  it('converts 1/96-inch page size with the requested dpi, capped by the width limit', () => {
    expect(script).toContain('$page.Size.Width * $dpi / 96');
    expect(script).toContain('[Math]::Min($maxW');
    expect(WINDOWS_PDF_MAX_WIDTH).toBe(8000);
  });
  it('clamps and de-duplicates pages like the Poppler path does', () => {
    expect(script).toContain('[Math]::Max(1, [Math]::Min($total, $_))');
    expect(script).toContain('Sort-Object -Unique | Select-Object -First $max');
  });
  it('fits on a command line', () => {
    expect(encodeCommand(script).length).toBeLessThan(MAX_ENCODED_COMMAND_CHARS);
  });
});

describe('protocol parser', () => {
  it('reads total, pages and ok', () => {
    expect(parseWindowsPdfProtocol(['TOTAL 5', 'PAGE 2', 'PAGE 4', 'OK'])).toEqual({ total: 5, pages: [2, 4], ok: true });
  });
  it('ignores garbage and reads errors', () => {
    expect(parseWindowsPdfProtocol(['junk', 'PAGE x', 'ERR 42 0x80070056 bad']).err).toEqual({ code: 42, hresult: '0x80070056', message: 'bad' });
  });
});

async function run(mode: string, extra: Record<string, string | undefined> = {}, o: Partial<Parameters<typeof windowsPdfToImages>[2]> = {}) {
  const d = tmp();
  const pdf = path.join(d, 'in.pdf');
  fs.writeFileSync(pdf, '%PDF-1.4');
  const out = path.join(d, 'out');
  fs.mkdirSync(out);
  return withFakePsEnv({ FAKE_PS_MODE: mode, ...extra }, async () => ({
    out,
    result: await windowsPdfToImages(pdf, out, { format: 'png', dpi: 150, jpegQuality: 90, maxPages: 4, timeoutMs: 20000, ...fakePsHooks, ...o }),
  }));
}
async function fail(mode: string, extra: Record<string, string | undefined> = {}, o: Partial<Parameters<typeof windowsPdfToImages>[2]> = {}): Promise<WindowsPdfError> {
  try {
    await run(mode, extra, o);
  } catch (e) {
    expect(e).toBeInstanceOf(WindowsPdfError);
    return e as WindowsPdfError;
  }
  throw new Error('expected a failure');
}

describe('windowsPdfToImages with a fake PowerShell', () => {
  it('renders the first maxPages pages when none are named', async () => {
    const { result, out } = await run('ok', { FAKE_TOTAL: '6' }, { maxPages: 3 });
    expect(result.totalPages).toBe(6);
    expect(result.images.map((i) => i.page)).toEqual([1, 2, 3]);
    expect(fs.readdirSync(out).sort()).toEqual(['page-1.png', 'page-2.png', 'page-3.png']);
  });
  it('clamps named pages to the real count, de-duplicates and sorts', async () => {
    const { result } = await run('ok', { FAKE_TOTAL: '3' }, { pages: [99, 2, 2, 0, -1, 1.5], maxPages: 4 });
    // 99 -> last page, 0 and -1 and 1.5 -> page 1 (the same clamping the Poppler path applies)
    expect(result.images.map((i) => i.page)).toEqual([1, 2, 3]);
  });
  it('passes dpi, quality, format and the width cap', async () => {
    const { result } = await run('ok', { FAKE_TOTAL: '1' }, { format: 'jpeg', dpi: 300, jpegQuality: 70 });
    expect(path.extname(result.images[0].file)).toBe('.jpg');
    expect(fs.readFileSync(result.images[0].file, 'utf8')).toBe('fake-jpeg-1-dpi300-q70-w8000');
  });
  it('drops non-finite page numbers and clamps huge ones before they reach the script', async () => {
    const { result } = await run('ok', { FAKE_TOTAL: '5' }, { pages: [2, Number.NaN, Number.POSITIVE_INFINITY, 2_000_000], maxPages: 4 });
    expect(result.images.map((i) => i.page)).toEqual([2, 5]);
  });
  it('a missing type (old Windows, Server Core) and constrained language are infrastructure failures', async () => {
    for (const mode of ['err41', 'constrained']) {
      const e = await fail(mode);
      expect(e.failureClass).toBe('infrastructure');
      expect(e.code).toBe('E_DEPENDENCY');
    }
  });
  it('an unreadable or password-protected PDF and a failed page are document failures', async () => {
    const e42 = await fail('err42');
    expect(e42.failureClass).toBe('document');
    expect(e42.exitCode).toBe(42);
    expect(e42.message).toMatch(/password-protected or damaged/);
    expect((await fail('err43')).failureClass).toBe('document');
  });
  it('a hang, a crash and PowerShell that cannot start are infrastructure failures', async () => {
    expect((await fail('hang', {}, { timeoutMs: 1500 })).message).toMatch(/did not respond/);
    expect((await fail('crash')).failureClass).toBe('infrastructure');
    const d = tmp();
    await expect(windowsPdfToImages(path.join(d, 'x.pdf'), d, { format: 'png', dpi: 150, jpegQuality: 90, maxPages: 1, powershellBin: path.join(d, 'nope'), powershellArgs: [] })).rejects.toMatchObject({ code: 'E_DEPENDENCY', failureClass: 'infrastructure' });
  });
  it('pages that were reported but not written are a failure, not a silent short result', async () => {
    expect((await fail('missing-files')).message).toMatch(/not written/);
  });
  it('a PDF with no pages produces a document failure', async () => {
    const e = await fail('zero-pages');
    expect(e.failureClass).toBe('document');
  });
});

/**
 * Microsoft Word as a DOCX -> PDF renderer (Windows only).
 *
 * Detection reads the registry and never starts Word. Rendering runs an embedded PowerShell script that
 * drives Word through COM: start a hidden Word, open the temporary copy read-only, export a PDF, close the
 * document, ask Word to quit. No window, no focus change, no recent-files entry, macros force-disabled.
 *
 * Honest limits (documented for users; `word` is opt-in in 0.5.2 for this reason):
 *  - COM may hand out a Word that another automation client already owns. The script checks before it
 *    changes anything and refuses (exit 13/14, nothing changed, nothing closed) when the instance looks used
 *    or predates the call. It cannot PROVE exclusive ownership, and `Quit()` after the checks is not atomic:
 *    another client that attaches in that moment could lose its Word.
 *  - ASTra never force-kills a Word process. After a timeout the (hidden) Word it started may still be
 *    running; the error names its PID.
 *  - A hidden Word does not exit by itself, so skipping `Quit()` would leak one process per render.
 */

import { spawnSync } from 'child_process';
import * as fs from 'fs';
import {
  POWERSHELL_PRELUDE,
  RunQueue,
  runPowerShell,
  sanitiseDetail,
  systemToolPath,
  type PowerShellRunOptions,
} from './ps-runner';
import { codedError, type FailureClass } from './renderer-preference';

export interface WordInstall {
  exe: string;
  /** Major version: 14 = 2010, 15 = 2013, 16 = 2016/2019/2021/365. */
  version: number;
}

export interface WordProbeOptions {
  platform?: NodeJS.Platform;
  /** Value of a registry key's default value, or null. Injected by tests. */
  regQuery?: (key: string) => string | null;
  exists?: (p: string) => boolean;
  env?: NodeJS.ProcessEnv;
}

const WORD_PROGID = 'Word.Application';
/** The Word versions that can export a PDF through `ExportAsFixedFormat` (2010 is the first). */
export const MIN_WORD_VERSION = 14;

function defaultRegQuery(key: string): string | null {
  // A 32-bit Node on 64-bit Windows (or the reverse) sees a different registry view, so try the default
  // view first, then each explicit view.
  for (const view of [undefined, '/reg:64', '/reg:32'] as const) {
    try {
      const args = ['query', key, '/ve', ...(view ? [view] : [])];
      const run = spawnSync(systemToolPath('reg.exe'), args, { windowsHide: true, timeout: 5000, encoding: 'utf8' });
      if (run.status !== 0 || typeof run.stdout !== 'string') continue;
      const m = /REG_(?:EXPAND_)?SZ\s+(.*?)\s*$/m.exec(run.stdout);
      if (m && m[1]) return m[1];
    } catch {
      /* try the next view */
    }
  }
  return null;
}

function expandEnv(value: string, env: NodeJS.ProcessEnv): string {
  return value.replace(/%([^%]+)%/g, (whole, name: string) => env[name] ?? env[name.toUpperCase()] ?? whole);
}

/**
 * Is Word installed and registered for automation? Reads `HKCR\Word.Application` (version) and the
 * `LocalServer32` command of its CLSID (the executable), and requires the executable to exist. Folder names
 * such as a leftover `Office15` are never used: the registry is the truth.
 */
export function findWord(o: WordProbeOptions = {}): WordInstall | null {
  try {
    if ((o.platform ?? process.platform) !== 'win32') return null;
    const env = o.env ?? process.env;
    const reg = o.regQuery ?? defaultRegQuery;
    const exists = o.exists ?? ((p: string) => fs.existsSync(p));
    const curVer = reg(`HKCR\\${WORD_PROGID}\\CurVer`); // "Word.Application.16"
    const m = /Word\.Application\.(\d+)/i.exec(curVer ?? '');
    if (!m) return null;
    const version = Number(m[1]);
    if (!Number.isFinite(version) || version < MIN_WORD_VERSION) return null;
    const clsid = reg(`HKCR\\${WORD_PROGID}\\CLSID`)?.trim();
    if (!clsid || !/^\{[0-9a-f-]{36}\}$/i.test(clsid)) return null;
    const server = reg(`HKCR\\CLSID\\${clsid}\\LocalServer32`) ?? reg(`HKCR\\WOW6432Node\\CLSID\\${clsid}\\LocalServer32`);
    if (!server) return null;
    // `"C:\...\WINWORD.EXE" /Automation` or an unquoted path with spaces followed by `/Automation`.
    const exeMatch = /^\s*"?([^"]+?\.exe)"?(?:\s|$)/i.exec(expandEnv(server, env));
    const exe = exeMatch?.[1];
    if (!exe || !exists(exe)) return null; // a stale registration is "not installed"
    return { exe, version };
  } catch {
    return null;
  }
}

// ---------------------------------------------------------------------------------------------------------
// The script
// ---------------------------------------------------------------------------------------------------------

/**
 * Exit codes: 0 ok, 11 Word could not be started, 12 constrained language, 13 instance in use or older than
 * this call (refused, untouched), 14 ambiguous instance or window/process mismatch (refused), 20 document
 * could not be opened, 22 PDF export failed.
 */
export function buildWordScript(): string {
  return `${POWERSHELL_PRELUDE}
Add-Type -Namespace N -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);'
function AutomationWord($id) { $c = Get-CimInstance Win32_Process -Filter "ProcessId=$id" -ErrorAction SilentlyContinue; [bool]($c -and $c.CommandLine -match '/Automation.*-Embedding') }
$in = $env:ASTRA_WORD_IN; $out = $env:ASTRA_WORD_OUT; $tagged = ($env:ASTRA_WORD_TAGGED -ne '0')
$mutex = New-Object Threading.Mutex($false, 'Local\\AstraWordLaunch'); $held = $false
$word = $null; $doc = $null; $cand = 0; $crossChecked = $false; $mismatch = $false; $alone = $false; $origAlerts = $null; $origSec = $null; $changed = $false
try {
  $held = $mutex.WaitOne(30000)
  $before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
  $alone = -not ($before | Where-Object { AutomationWord $_ })
  try { $word = New-Object -ComObject Word.Application } catch { Die 11 $_ }
  if ($held) { $mutex.ReleaseMutex(); $held = $false }
  $fresh = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | Where-Object {
      (AutomationWord $_.Id) -and (Get-Process -Id (Get-CimInstance Win32_Process -Filter "ProcessId=$($_.Id)").ParentProcessId -ErrorAction SilentlyContinue).ProcessName -eq 'svchost' } | ForEach-Object Id)
  $quiet = $false; try { $quiet = (-not $word.Visible) -and ($word.Documents.Count -eq 0) } catch {}
  if (-not $quiet)         { $word = $null; Say 'ERR 13 0x00000000 instance in use'; exit 13 }
  if ($fresh.Count -eq 0)  { $word = $null; Say 'ERR 13 0x00000000 instance predates this call'; exit 13 }
  if ($fresh.Count -ne 1)  { $word = $null; Say 'ERR 14 0x00000000 several new Word processes'; exit 14 }
  $cand = $fresh[0]; Say "CANDIDATE $cand"
  try { Say ('VERSION {0}' -f $word.Version) } catch {}
  $origAlerts = $word.DisplayAlerts; $origSec = $word.AutomationSecurity
  $word.DisplayAlerts = 0; $word.AutomationSecurity = 3; $changed = $true
  try { $doc = $word.Documents.Open($in, $false, $true, $false) } catch { Die 20 $_ }
  $p = 0; [void][N.U]::GetWindowThreadProcessId([IntPtr]$doc.ActiveWindow.Hwnd, [ref]$p)
  if ($p -ne $cand) { $mismatch = $true; Say 'ERR 14 0x00000000 window process differs from candidate'; exit 14 }
  $crossChecked = $true
  Say ('PAGES {0}' -f $doc.ComputeStatistics(2))
  try { $doc.ExportAsFixedFormat($out, 17, $false, 0, 0, 1, 1, 0, $true, $true, 0, $tagged, $true, $false) } catch { Die 22 $_ }
  Say 'OK'
}
finally {
  if ($held) { try { $mutex.ReleaseMutex() } catch {} }
  if ($doc) { try { $doc.Close(0) } catch {} }
  if ($word -and $cand) {
    $mayQuit = $false
    try { $stillAuto = AutomationWord $cand
      $mayQuit = $stillAuto -and (-not $mismatch) -and ($crossChecked -or $alone) -and (-not $word.Visible) -and ($word.UserControl -eq $false) -and ($word.Documents.Count -eq 0) } catch {}
    if ($mayQuit) { try { $word.Quit() } catch {} }
    elseif ($changed) { try { $word.DisplayAlerts = $origAlerts; $word.AutomationSecurity = $origSec } catch {} }
  }
}
`;
}

// ---------------------------------------------------------------------------------------------------------
// Protocol and errors
// ---------------------------------------------------------------------------------------------------------

export interface WordProtocol {
  candidatePid?: number;
  version?: string;
  pages?: number;
  ok: boolean;
  err?: { code: number; hresult: string; message: string };
}

export function parseWordProtocol(lines: string[]): WordProtocol {
  const p: WordProtocol = { ok: false };
  for (const line of lines) {
    let m: RegExpExecArray | null;
    if ((m = /^CANDIDATE (\d+)$/.exec(line))) p.candidatePid = Number(m[1]);
    else if ((m = /^VERSION (\S+)$/.exec(line))) p.version = m[1];
    else if ((m = /^PAGES (\d+)$/.exec(line))) p.pages = Number(m[1]);
    else if (line === 'OK') p.ok = true;
    else if ((m = /^ERR (\d+) (0x[0-9A-Fa-f]{8}) ?(.*)$/.exec(line))) p.err = { code: Number(m[1]), hresult: `0x${m[2].slice(2).toUpperCase()}`, message: sanitiseDetail(m[3]) };
  }
  return p;
}

export class WordRenderError extends Error {
  readonly code: 'E_DEPENDENCY' | 'E_VALIDATION' | 'E_RENDER';
  readonly remedy: string;
  readonly failureClass: FailureClass;
  readonly exitCode?: number;
  readonly hresult?: string;
  /** PID of the hidden Word this call started, when it is known and may still be running. */
  readonly candidatePid?: number;
  constructor(
    code: WordRenderError['code'],
    failureClass: FailureClass,
    message: string,
    remedy: string,
    extra: { exitCode?: number; hresult?: string; candidatePid?: number } = {}
  ) {
    super(message);
    this.name = 'WordRenderError';
    this.code = code;
    this.failureClass = failureClass;
    this.remedy = remedy;
    this.exitCode = extra.exitCode;
    this.hresult = extra.hresult;
    this.candidatePid = extra.candidatePid;
  }
}

/** `0x800A13E9` (password) and the other document-level Word errors all live in the 0x800A range. */
function isDocumentHresult(h: string): boolean {
  return /^0X800A/i.test(h);
}

function classifyFailure(p: WordProtocol, exitCode: number | null, stderr: string): WordRenderError {
  const e = p.err;
  const pid = p.candidatePid;
  const orphanNote = pid !== undefined ? ` A hidden Word process (pid ${pid}) may still be running; ASTra did not stop it. End it from Task Manager if it stays.` : '';
  if (!e) {
    const detail = sanitiseDetail(stderr, 200);
    return new WordRenderError(
      'E_RENDER',
      'infrastructure',
      `E_RENDER: Word renderer ended without a result (exit code ${exitCode ?? 'none'})${detail ? `: ${detail}` : ''}.${orphanNote}`,
      'Use --renderer libreoffice, or check that Windows PowerShell and Word start normally.',
      { exitCode: exitCode ?? undefined, candidatePid: pid }
    );
  }
  const common = { exitCode: e.code, hresult: e.hresult, candidatePid: pid };
  switch (e.code) {
    case 11:
      return new WordRenderError('E_DEPENDENCY', 'infrastructure', `E_DEPENDENCY: Word could not be started (${e.hresult}: ${e.message}).`, 'Repair or re-activate Office, or use --renderer libreoffice.', common);
    case 12:
      return new WordRenderError('E_DEPENDENCY', 'infrastructure', 'E_DEPENDENCY: PowerShell is in constrained language mode (policy), so Word cannot be driven.', 'Use --renderer libreoffice.', common);
    case 13:
    case 14:
      return new WordRenderError(
        'E_RENDER',
        'infrastructure',
        `E_RENDER: Word was not used because the Word instance Windows provided could not be confirmed as private to this render (${e.message}). Nothing was changed or closed.`,
        'Close other programs that automate Word and retry, or use --renderer libreoffice.',
        common
      );
    case 20: {
      const password = /0X800A13E9/i.test(e.hresult);
      return new WordRenderError(
        'E_VALIDATION',
        'document',
        password
          ? 'E_VALIDATION: Word could not open the document: it is password-protected, and password-protected documents cannot be rendered.'
          : `E_VALIDATION: Word could not open the document (${e.hresult}: ${e.message}).`,
        password ? 'Remove the password protection and retry.' : 'Verify the .docx opens in Word and retry.',
        common
      );
    }
    case 22:
      return new WordRenderError(
        'E_RENDER',
        isDocumentHresult(e.hresult) ? 'document' : 'infrastructure',
        `E_RENDER: Word could not export the PDF (${e.hresult}: ${e.message}).${orphanNote}`,
        'Verify the document opens and prints in Word, or use --renderer libreoffice.',
        common
      );
    default:
      return new WordRenderError('E_RENDER', 'infrastructure', `E_RENDER: Word renderer failed (code ${e.code}, ${e.hresult}: ${e.message}).${orphanNote}`, 'Use --renderer libreoffice.', common);
  }
}

// ---------------------------------------------------------------------------------------------------------
// Runner
// ---------------------------------------------------------------------------------------------------------

export interface WordRenderOptions {
  tagged?: boolean;
  timeoutMs?: number;
  /** Tests only. */
  powershellBin?: string;
  powershellArgs?: string[];
  /** Pause range (ms) between retries for an instance that could not be confirmed private. Tests set 0. */
  retryPauseMs?: [number, number];
}

export interface WordRenderOutcome {
  pages?: number;
  /** Word's own version string (for example "16.0"). */
  version?: string;
}

const wordQueue = new RunQueue();
const MAX_WORD_ATTEMPTS = 3;
const MAX_PATH_CHARS = 240;

function pdfLooksValid(file: string): boolean {
  try {
    const fd = fs.openSync(file, 'r');
    try {
      const head = Buffer.alloc(5);
      const n = fs.readSync(fd, head, 0, 5, 0);
      return n === 5 && head.toString('latin1') === '%PDF-';
    } finally {
      fs.closeSync(fd);
    }
  } catch {
    return false;
  }
}

/**
 * Render `inDocx` to `outPdf` with Word. One render at a time per process. Throws WordRenderError.
 * `failureClass` on the error says whether the machine (infrastructure) or the document is at fault.
 */
export function wordDocxToPdf(inDocx: string, outPdf: string, options: WordRenderOptions = {}): Promise<WordRenderOutcome> {
  return wordQueue.run(async () => {
    if (inDocx.length > MAX_PATH_CHARS || outPdf.length > MAX_PATH_CHARS) {
      throw new WordRenderError(
        'E_RENDER',
        'infrastructure',
        'E_RENDER: the temporary path is too long for Word (over 240 characters).',
        'Set ASTRA_TMPDIR to a shorter folder, or use --renderer libreoffice.'
      );
    }
    const timeoutMs = options.timeoutMs ?? 60000;
    const deadline = Date.now() + timeoutMs;
    const script = buildWordScript();
    const pause = options.retryPauseMs ?? [100, 400];
    let last: WordRenderError | undefined;
    for (let attempt = 1; attempt <= MAX_WORD_ATTEMPTS; attempt++) {
      const remaining = deadline - Date.now();
      if (remaining < 1000) break;
      const run: PowerShellRunOptions = {
        script,
        env: { ASTRA_WORD_IN: inDocx, ASTRA_WORD_OUT: outPdf, ASTRA_WORD_TAGGED: options.tagged === false ? '0' : '1' },
        timeoutMs: remaining,
        powershellBin: options.powershellBin,
        powershellArgs: options.powershellArgs,
      };
      const res = await runPowerShell(run);
      const proto = parseWordProtocol(res.lines);
      if (res.spawnError) {
        throw new WordRenderError(
          'E_DEPENDENCY',
          'infrastructure',
          `E_DEPENDENCY: Windows PowerShell could not be started (${sanitiseDetail(res.spawnError.code ?? res.spawnError.message, 80)}).`,
          'PowerShell is blocked or missing; use --renderer libreoffice.'
        );
      }
      if (res.timedOut) {
        const pid = proto.candidatePid;
        throw new WordRenderError(
          'E_RENDER',
          'infrastructure',
          `E_RENDER: Word did not respond within ${Math.round(timeoutMs / 1000)} s.${pid !== undefined ? ` A hidden Word process (pid ${pid}) may still be running; ASTra did not stop it. End it from Task Manager if it stays.` : ' A hidden Word process may still be running; ASTra did not stop it.'}`,
          'Check that Word opens normally (a first-run or activation dialog blocks automation), or use --renderer libreoffice.',
          { candidatePid: pid }
        );
      }
      if (res.exitCode === 0 && proto.ok) {
        if (!fs.existsSync(outPdf) || !pdfLooksValid(outPdf)) {
          throw new WordRenderError('E_RENDER', 'infrastructure', 'E_RENDER: Word reported success but produced no valid PDF.', 'Use --renderer libreoffice.', { candidatePid: proto.candidatePid });
        }
        return { pages: proto.pages, version: proto.version };
      }
      last = classifyFailure(proto, res.exitCode, res.stderr);
      if (last.exitCode === 13 || last.exitCode === 14) {
        if (attempt < MAX_WORD_ATTEMPTS) {
          const [lo, hi] = pause;
          const wait = hi > 0 ? lo + Math.floor(Math.random() * Math.max(1, hi - lo)) : 0;
          if (wait > 0) await new Promise((r) => setTimeout(r, wait));
          continue;
        }
      }
      throw last;
    }
    throw last ?? new WordRenderError('E_RENDER', 'infrastructure', 'E_RENDER: Word render ran out of time before it started.', 'Use --renderer libreoffice.');
  });
}

/** Thrown for an invalid request before anything runs. */
export function wordUnavailableError(): Error {
  return codedError('E_DEPENDENCY', 'E_DEPENDENCY: Microsoft Word is not available.', 'Install Word, or use --renderer libreoffice.');
}

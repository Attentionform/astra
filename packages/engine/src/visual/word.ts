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
import { isEncryptedOoxmlPackage } from './cfb';

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
  // The default probe starts up to three `reg.exe` processes per registry key, and long-lived callers (the MCP
  // server) ask on every render, so the answer is remembered briefly. An injected probe is never cached.
  const cacheable = o.platform === undefined && o.regQuery === undefined && o.exists === undefined && o.env === undefined;
  if (cacheable && process.platform === 'win32') {
    const now = Date.now();
    if (wordProbeCache && now - wordProbeCache.at < WORD_PROBE_TTL_MS) return wordProbeCache.value;
    const value = probeWord(o);
    wordProbeCache = { at: now, value };
    return value;
  }
  return probeWord(o);
}

const WORD_PROBE_TTL_MS = 30_000;
let wordProbeCache: { at: number; value: WordInstall | null } | undefined;
/** Forget the remembered probe result (tests, and callers that just installed or removed Word). */
export function resetWordProbeCache(): void {
  wordProbeCache = undefined;
}

function probeWord(o: WordProbeOptions): WordInstall | null {
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
 * A starting state for the review view, applied on the SAME Word instance and document right before the view is
 * pinned. Tests only: a smoke test needs a bad starting state ("No Markup", comments off, a reviewer hidden) that is
 * really in effect on the instance ASTra uses, which a separate interactive Word window cannot provide.
 */
export interface WordStartState {
  /** 0 balloons, 1 inline, 2 mixed. */
  markupMode?: 0 | 1 | 2;
  showComments?: boolean;
  /** Hide the first reviewer, by index (no name is read). */
  hideFirstReviewer?: boolean;
  /** "No Markup": revisions and comments hidden in the review pane. */
  noMarkup?: boolean;
}

/** Strict validation: only the three known keys, only their allowed values. Throws on anything else. */
export function validateStartState(state: WordStartState): WordStartState {
  if (state === null || typeof state !== 'object' || Array.isArray(state)) throw new TypeError('startState must be an object');
  const out: WordStartState = {};
  for (const [k, v] of Object.entries(state)) {
    if (k === 'markupMode' && (v === 0 || v === 1 || v === 2)) out.markupMode = v;
    else if (k === 'showComments' && typeof v === 'boolean') out.showComments = v;
    else if (k === 'hideFirstReviewer' && typeof v === 'boolean') out.hideFirstReviewer = v;
    else if (k === 'noMarkup' && typeof v === 'boolean') out.noMarkup = v;
    else throw new TypeError(`startState: unsupported key or value (${k})`);
  }
  return out;
}

/** The printed form of the view state: numbers only, so no reviewer name or document content can appear. */
function startStateReadout(label: 'STARTSTATE' | 'ENDSTATE'): string {
  return [
    "\n    $hid = 0; try { for ($i = 1; $i -le [int]$view.Reviewers.Count; $i++) { if (-not $view.Reviewers.Item($i).Visible) { $hid++ } } } catch {}",
    `    try { Say ('${label} mode={0} comments={1} hidden={2} revs={3} filter={4}' -f $view.MarkupMode, [int][bool]$view.ShowComments, $hid, [int][bool]$view.ShowRevisionsAndComments, [int]$view.RevisionsFilter.Markup) } catch {}`,
  ].join('\n');
}

/** PowerShell for a validated start state. Built only from the validated numbers and booleans. */
function startStateFragment(state: WordStartState, label: 'STARTSTATE' | 'ENDSTATE'): string {
  const lines: string[] = [];
  if (state.markupMode !== undefined) lines.push(`    try { $view.MarkupMode = ${state.markupMode} } catch {}`);
  if (state.showComments !== undefined) lines.push(`    try { $view.ShowComments = ${state.showComments ? '$true' : '$false'} } catch {}`);
  if (state.hideFirstReviewer) lines.push('    try { $view.Reviewers.Item(1).Visible = $false } catch {}');
  if (state.noMarkup) lines.push('    try { $view.ShowRevisionsAndComments = $false } catch {}', '    try { $view.RevisionsFilter.Markup = 0 } catch {}');
  return `\n${lines.join('\n')}${startStateReadout(label)}`;
}

/**
 * Exit codes: 0 ok, 11 Word could not be started, 12 constrained language, 13 instance in use or older than
 * this call (refused, untouched), 14 ambiguous instance or window/process mismatch (refused), 20 document
 * could not be opened, 22 PDF export failed, 23 the review view could not be set up (markup requested only).
 *
 * `startState` is for tests only. It is a validated description, never script text, and it is not reachable
 * from an environment variable: the production call passes nothing, so the production script has no such step.
 */
export function buildWordScript(startState?: WordStartState): string {
  const state = startState === undefined ? undefined : validateStartState(startState);
  const hook = state === undefined ? '' : startStateFragment(state, 'STARTSTATE');
  const endHook = state === undefined ? '' : startStateReadout('ENDSTATE');
  return `${POWERSHELL_PRELUDE}
Add-Type -Namespace N -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);'
function AutomationWord($id) { $c = Get-CimInstance Win32_Process -Filter "ProcessId=$id" -ErrorAction SilentlyContinue; [bool]($c -and $c.CommandLine -match '/Automation.*-Embedding') }
$in = $env:ASTRA_WORD_IN; $out = $env:ASTRA_WORD_OUT; $tagged = ($env:ASTRA_WORD_TAGGED -ne '0'); $markup = ($env:ASTRA_WORD_MARKUP -eq '1')
function Fail23($what) { Say ('ERR 23 0x00000000 ' + $what); exit 23 }
$view = $null; $viewOrig = $null; $viewTouched = $false; $best = @()
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
  # A password question is a visible dialog that never closes, even for a hidden Word. A made-up password (to open, and
  # to write) makes Word answer "the password is incorrect" at once instead of asking.
  $m = [Type]::Missing; $pw = [Guid]::NewGuid().ToString('N')
  try { $doc = $word.Documents.Open($in, $false, $true, $false, $pw, $m, $m, $pw) } catch { Die 20 $_ }
  $p = 0; [void][N.U]::GetWindowThreadProcessId([IntPtr]$doc.ActiveWindow.Hwnd, [ref]$p)
  if ($p -ne $cand) { $mismatch = $true; Say 'ERR 14 0x00000000 window process differs from candidate'; exit 14 }
  $crossChecked = $true
  Say ('PAGES {0}' -f $doc.ComputeStatistics(2))
  $item = 0
  if ($markup) {
    # The review view. Every original value is recorded first and written back in the finally block. Required
    # settings are set and read back; a failure is exit 23 and nothing is exported. Reviewers are addressed by
    # index, so no author name is read or printed.
    try { $view = $doc.ActiveWindow.View } catch { Fail23 'the document view is not available' }${hook}
    $viewOrig = @{ Mode = $null; Comments = $null; Rev = @() }
    try {
      $viewOrig.Mode = $view.MarkupMode; $viewOrig.Comments = $view.ShowComments
      $n = [int]$view.Reviewers.Count
      for ($i = 1; $i -le $n; $i++) { $viewOrig.Rev += [bool]$view.Reviewers.Item($i).Visible }
    } catch { Fail23 'the current review view could not be read' }
    $rf = $null; try { $rf = $view.RevisionsFilter } catch {}
    foreach ($b in @(
        @{ O = $view; N = 'ShowRevisionsAndComments'; V = $true },
        @{ O = $view; N = 'RevisionsView'; V = 0 },
        @{ O = $view; N = 'ShowInsertionsAndDeletions'; V = $true },
        @{ O = $view; N = 'ShowFormatChanges'; V = $true },
        @{ O = $rf; N = 'Markup'; V = 2; L = 'RevisionsFilter.Markup' },
        @{ O = $rf; N = 'View'; V = 0; L = 'RevisionsFilter.View' })) {
      if ($null -eq $b.O) { continue }
      if (-not $b.L) { $b.L = $b.N }
      $b.Orig = $null; $b.Had = $false
      try { $b.Orig = $b.O.($b.N); $b.Had = $true } catch {}
      $best += $b
    }
    $viewTouched = $true
    try { $view.MarkupMode = 2 } catch {}
    if ($view.MarkupMode -ne 2) { Fail23 'MarkupMode (inline marks with comment balloons)' }
    try { $view.ShowComments = $true } catch {}
    if ($view.ShowComments -ne $true) { Fail23 'ShowComments' }
    $n = [int]$view.Reviewers.Count
    for ($i = 1; $i -le $n; $i++) {
      try { $view.Reviewers.Item($i).Visible = $true } catch {}
      if (-not $view.Reviewers.Item($i).Visible) { Fail23 'a reviewer could not be made visible' }
    }
    foreach ($b in $best) { try { $b.O.($b.N) = $b.V } catch { Say ('WARN ' + $b.L) } }
    Say ('MARKUP {0}' -f $view.MarkupMode)
    $item = 7
  }
  try { $doc.ExportAsFixedFormat($out, 17, $false, 0, 0, 1, 1, $item, $true, $true, 0, $tagged, $true, $false) } catch { Die 22 $_ }
  Say 'OK'
}
finally {
  if ($held) { try { $mutex.ReleaseMutex() } catch {} }
  if ($viewTouched -and $doc -and $view) {
    # Put the view back before the document is closed, then read it back. Best effort: a killed process cannot do this.
    $bad = @()
    try { $view.MarkupMode = $viewOrig.Mode; if ($view.MarkupMode -ne $viewOrig.Mode) { $bad += 'MarkupMode' } } catch { $bad += 'MarkupMode' }
    try { $view.ShowComments = $viewOrig.Comments; if ($view.ShowComments -ne $viewOrig.Comments) { $bad += 'ShowComments' } } catch { $bad += 'ShowComments' }
    try {
      for ($i = 1; $i -le $viewOrig.Rev.Count; $i++) {
        $view.Reviewers.Item($i).Visible = $viewOrig.Rev[$i - 1]
        if ([bool]$view.Reviewers.Item($i).Visible -ne $viewOrig.Rev[$i - 1]) { $bad += 'reviewer visibility'; break }
      }
    } catch { $bad += 'reviewer visibility' }
    foreach ($b in $best) { if ($b.Had) { try { $b.O.($b.N) = $b.Orig; if ($b.O.($b.N) -ne $b.Orig) { $bad += $b.L } } catch { $bad += $b.L } } }
    if ($bad.Count -eq 0) { Say 'VIEWRESTORE ok' } else { Say ('VIEWRESTORE failed ' + (($bad | Select-Object -Unique) -join ',')) }${endHook}
  }
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
  /** `MARKUP n`: the `MarkupMode` that was set and read back (markup requested only). */
  markupMode?: number;
  /** `WARN <setting>`: a best-effort view setting that could not be set. Names come from a fixed list. */
  warnings: string[];
  /** `VIEWRESTORE`: did the view return to what it was? `undefined` when no view was changed. */
  restore?: { ok: boolean; failed: string[] };
  /** `STARTSTATE ...` and `ENDSTATE ...`: numbers only, printed by the test-only step. */
  startState?: string;
  endState?: string;
  ok: boolean;
  err?: { code: number; hresult: string; message: string };
}

export function parseWordProtocol(lines: string[]): WordProtocol {
  const p: WordProtocol = { ok: false, warnings: [] };
  for (const line of lines) {
    let m: RegExpExecArray | null;
    if ((m = /^CANDIDATE (\d+)$/.exec(line))) p.candidatePid = Number(m[1]);
    else if ((m = /^VERSION (\S+)$/.exec(line))) p.version = m[1];
    else if ((m = /^PAGES (\d+)$/.exec(line))) p.pages = Number(m[1]);
    else if ((m = /^MARKUP (\d)$/.exec(line))) p.markupMode = Number(m[1]);
    else if ((m = /^WARN ([A-Za-z.]{1,40})$/.exec(line))) p.warnings.push(m[1]);
    else if (line === 'VIEWRESTORE ok') p.restore = { ok: true, failed: [] };
    else if ((m = /^VIEWRESTORE failed ([A-Za-z.,\s]{1,200})$/.exec(line))) p.restore = { ok: false, failed: m[1].split(',').map((x) => x.trim()).filter(Boolean) };
    else if ((m = /^STARTSTATE ([a-z0-9= ]{1,80})$/.exec(line))) p.startState = m[1];
    else if ((m = /^ENDSTATE ([a-z0-9= ]{1,80})$/.exec(line))) p.endState = m[1];
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
      // 0x800A1520: the made-up password was refused. 0x800A13E9: the older form of the same error.
      const password = /0X800A(?:1520|13E9)/i.test(e.hresult);
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
    case 23:
      // The engine itself is healthy; it cannot honour the review view. `request` keeps the breaker closed and
      // still lets `auto` fall through to an engine that can.
      return new WordRenderError(
        'E_RENDER',
        'request',
        `E_RENDER: Word could not be set to show comments and tracked changes (${e.message}).`,
        'Render without --show-markup, or use --renderer libreoffice.',
        common
      );
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
  /** Draw tracked changes and comments (the review view). Default false: content only, as before. */
  markup?: boolean;
  /** Tests only; see `WordStartState`. Ignored unless `markup` is true. */
  startState?: WordStartState;
}

export interface WordRenderOutcome {
  pages?: number;
  /** Word's own version string (for example "16.0"). */
  version?: string;
  /** Present when markup was requested. */
  markup?: {
    mode: number;
    /** Best-effort view settings that could not be set, and a restore that did not complete. Fixed wording. */
    warnings: string[];
    /** `STARTSTATE` / `ENDSTATE` lines of the test-only step, when it ran. */
    startState?: string;
    endState?: string;
    restore?: { ok: boolean; failed: string[] };
  };
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
    // A password-protected package is known before Word is started: say so at once instead of starting Word.
    try {
      const head = fs.readFileSync(inDocx);
      if (isEncryptedOoxmlPackage(head)) {
        throw new WordRenderError(
          'E_VALIDATION',
          'document',
          'E_VALIDATION: the document is protected with a password, so it cannot be rendered.',
          'Remove the password in Word (File > Info > Protect Document), then retry.'
        );
      }
    } catch (err: unknown) {
      if (err instanceof WordRenderError) throw err;
      /* an unreadable input is reported by the run below */
    }
    const timeoutMs = options.timeoutMs ?? 60000;
    const deadline = Date.now() + timeoutMs;
    const markup = options.markup === true;
    const script = buildWordScript(markup ? options.startState : undefined);
    const pause = options.retryPauseMs ?? [100, 400];
    let last: WordRenderError | undefined;
    for (let attempt = 1; attempt <= MAX_WORD_ATTEMPTS; attempt++) {
      const remaining = deadline - Date.now();
      if (remaining < 1000) break;
      const run: PowerShellRunOptions = {
        script,
        env: { ASTRA_WORD_IN: inDocx, ASTRA_WORD_OUT: outPdf, ASTRA_WORD_TAGGED: options.tagged === false ? '0' : '1', ...(markup ? { ASTRA_WORD_MARKUP: '1' } : {}) },
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
        if (markup && proto.markupMode !== 2) {
          // The script prints this only after the view was set and read back. Without it the review view is unconfirmed.
          throw new WordRenderError(
            'E_RENDER',
            'request',
            'E_RENDER: Word exported a PDF but did not confirm the review view, so the result cannot be called a markup render.',
            'Render without --show-markup, or use --renderer libreoffice.',
            { candidatePid: proto.candidatePid }
          );
        }
        const warnings = proto.warnings.map((w) => `Word setting ${w} could not be set (the export did not need it on the tested Word).`);
        if (proto.restore && !proto.restore.ok) {
          warnings.push(`Word's review view could not be fully restored (${proto.restore.failed.join(', ') || 'unknown setting'}); it applied to the render's own Word window only.`);
        }
        return {
          pages: proto.pages,
          version: proto.version,
          ...(markup ? { markup: { mode: proto.markupMode as number, warnings, ...(proto.startState ? { startState: proto.startState } : {}), ...(proto.endState ? { endState: proto.endState } : {}), ...(proto.restore ? { restore: proto.restore } : {}) } } : {}),
        };
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

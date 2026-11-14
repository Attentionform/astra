/**
 * "Show markup": opt-in display of tracked changes and comments in a render.
 *
 * What this file holds:
 *  - the switch (`ASTRA_SHOW_MARKUP`, `showMarkup`, `--show-markup`) and its parsing, like the other render switches;
 *  - the LibreOffice PDF filter spec, with and without the margin-comments option;
 *  - the LibreOffice version check. LibreOffice silently ignores a filter option it does not know (exit 0, no
 *    comments), so the option is only sent to a version that has been tested. A version is a guard, not proof;
 *  - how a failed markup conversion is reported (the real error is kept, nothing is claimed beyond the evidence);
 *  - the `markup` block of a result and the two sentences every caller shows.
 *
 * Contract (kept narrow on purpose): a markup render is produced by a different layout mode of the renderer. It can
 * scale the content and change pagination, so a page number refers to the render that was produced. Which marks are
 * drawn, and how, is the renderer's own. ASTra does not check that every comment or change appears in the image.
 */

import { spawnSync } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';
import { codedError, type FailureClass } from './renderer-preference';

export const SHOW_MARKUP_ENV = 'ASTRA_SHOW_MARKUP';

/** Parse the switch from a flag, parameter or ASTRA_SHOW_MARKUP. Empty means "not set". */
export function parseShowMarkup(raw: unknown, source: 'flag' | 'parameter' | typeof SHOW_MARKUP_ENV = 'parameter'): boolean | undefined {
  if (raw === undefined || raw === null) return undefined;
  if (typeof raw === 'boolean') return raw;
  if (typeof raw === 'string') {
    const v = raw.trim().toLowerCase();
    if (v === '') return undefined;
    if (['1', 'true', 'yes', 'on'].includes(v)) return true;
    if (['0', 'false', 'no', 'off'].includes(v)) return false;
  }
  throw codedError(
    'E_VALIDATION',
    `E_VALIDATION: ${source} showMarkup must be true or false (1/0, yes/no, on/off).`,
    `Use ${SHOW_MARKUP_ENV}=1 to show tracked changes and comments, ${SHOW_MARKUP_ENV}=0 (or unset it) for the default render.`
  );
}

/** Explicit per-call value (including an explicit `false`) > ASTRA_SHOW_MARKUP > false. */
export function effectiveShowMarkup(explicit: boolean | string | undefined, env: NodeJS.ProcessEnv = process.env): boolean {
  const fromCall = parseShowMarkup(explicit, 'parameter');
  if (fromCall !== undefined) return fromCall;
  return parseShowMarkup(env[SHOW_MARKUP_ENV], SHOW_MARKUP_ENV) ?? false;
}

// ---------------------------------------------------------------------------------------------------------
// What a result reports, and the sentences callers show
// ---------------------------------------------------------------------------------------------------------

export type MarkupMode = 'export-with-markup' | 'comments-in-margin';

/**
 * What was asked and what the renderer was told to do. It is not a check of the output: `mode` names the setting the
 * renderer was instructed to use, and `warnings` carries best-effort Word settings that could not be set.
 */
export interface MarkupReport {
  requested: true;
  renderer: 'word' | 'libreoffice';
  mode: MarkupMode;
  warnings: string[];
}

export function buildMarkupReport(renderer: 'word' | 'libreoffice', warnings: string[] = []): MarkupReport {
  return { requested: true, renderer, mode: renderer === 'word' ? 'export-with-markup' : 'comments-in-margin', warnings: [...warnings] };
}

export const MARKUP_CONTRACT_NOTE = 'markup view: pages are scaled and may differ from a render without markup; page numbers refer to this render.';
export const MARKUP_PRIVACY_NOTE = 'markup view: the output shows reviewer names, deleted text and comments; check it before sharing it.';

/** The notes shown with every markup render (CLI stderr, MCP `notes`), then any warnings. */
export function markupNotes(report: MarkupReport): string[] {
  return [MARKUP_CONTRACT_NOTE, MARKUP_PRIVACY_NOTE, ...report.warnings.map((w) => `markup warning: ${w}`)];
}

// ---------------------------------------------------------------------------------------------------------
// LibreOffice PDF filter spec
// ---------------------------------------------------------------------------------------------------------

/**
 * `pdf:writer_pdf_Export` with optional JSON filter data. The untagged form must be JSON (LibreOffice 25.2 rejects
 * the semicolon form). `ExportNotesInMargin` draws each comment in a box in the right margin; tracked changes are
 * drawn by LibreOffice's default view already. With no options the spec is the plain filter name.
 */
export function libreOfficeFilterSpec(tagged: boolean, markup = false): string {
  const data: string[] = [];
  if (!tagged) data.push('"UseTaggedPDF":{"type":"boolean","value":"false"}');
  if (markup) data.push('"ExportNotesInMargin":{"type":"boolean","value":"true"}');
  return data.length > 0 ? `pdf:writer_pdf_Export:{${data.join(',')}}` : 'pdf:writer_pdf_Export';
}

// ---------------------------------------------------------------------------------------------------------
// LibreOffice version: a tested minimum, used as a guard
// ---------------------------------------------------------------------------------------------------------

/**
 * The lowest LibreOffice version on which the acceptance checks L1, L2a, L3 and L4 (scripts/word-markup-smoke.mjs)
 * passed: the comment text, its author and the deleted text are in the PDF, and the same render without the option has
 * no comments. Tested down to 7.6.7; older builds were not tried. Changing this constant needs a new row in the version
 * table in docs/page-layout.md.
 */
export const MIN_LIBREOFFICE_FOR_MARKUP = '7.6.7';

/**
 * Observed on every tested version (7.6.7, 24.8.7, 25.2.7, 25.8.7, 26.2.7, 26.8.0; check L2b): LibreOffice's margin
 * box does not always show a long comment in full. 24.8, 25.2 and 25.8 let the box run past the page's right edge, so a
 * line longer than about 40 characters loses its tail; 7.6, 26.2 and 26.8 wrap inside the page but shorten the box
 * ("...") when other comments follow close below. No tested version showed a long comment complete beside its
 * neighbours, so every LibreOffice markup render carries this warning.
 */
export const LIBREOFFICE_LONG_COMMENT_WARNING =
  'LibreOffice can show a long comment cut short in the margin (at the page edge on some versions, or at its end when other comments follow close below). Word on Windows shows comments in full balloons.';

/** Warnings for a markup render by LibreOffice. */
export function libreOfficeMarkupWarnings(): string[] {
  return [LIBREOFFICE_LONG_COMMENT_WARNING];
}

export interface LibreOfficeVersion {
  major: number;
  minor: number;
  micro: number;
  text: string;
}

/** `LibreOffice 26.8.0.3 <build id>` at the start of a line. Anything else is not a version. */
export function parseLibreOfficeVersion(output: string): LibreOfficeVersion | null {
  const m = /^LibreOffice\s+(\d+)\.(\d+)\.(\d+)(?:\.(\d+))?/m.exec(output);
  if (!m) return null;
  return { major: Number(m[1]), minor: Number(m[2]), micro: Number(m[3]), text: `${m[1]}.${m[2]}.${m[3]}${m[4] !== undefined ? `.${m[4]}` : ''}` };
}

export function compareVersion(a: LibreOfficeVersion, b: LibreOfficeVersion): number {
  return a.major - b.major || a.minor - b.minor || a.micro - b.micro;
}

function minimumVersion(): LibreOfficeVersion {
  return parseLibreOfficeVersion(`LibreOffice ${MIN_LIBREOFFICE_FOR_MARKUP}`) as LibreOfficeVersion;
}

export interface ProbeRun {
  status: number | null;
  stdout: string;
  timedOut: boolean;
  /** Spawn failure (missing file, not executable, a script that needs a shell). */
  error?: string;
}
export type ProbeRunner = (file: string, args: string[], timeoutMs: number) => ProbeRun;

export interface VersionProbeOptions {
  platform?: NodeJS.Platform;
  exists?: (p: string) => boolean;
  run?: ProbeRunner;
  timeoutMs?: number;
  now?: () => number;
}

export type VersionProbe = { ok: true; version: LibreOfficeVersion } | { ok: false; reason: string };

/**
 * Which file answers `--version` for the selected LibreOffice binary. The selected binary is the one that will do the
 * conversion; nothing else is searched.
 *  - Windows `soffice.exe`: the console launcher `soffice.com` in the same folder (`soffice.exe --version` prints
 *    nothing to a pipe). No `.com` beside it: unreadable. `program\version.ini` is not used (it holds only a build id).
 *  - Anything else (a wrapper, a shim, a link, a renamed or configured binary, any other platform): the file itself.
 */
export function versionProbeFile(bin: string, platform: NodeJS.Platform, exists: (p: string) => boolean): string | null {
  if (platform === 'win32') {
    const p = path.win32;
    if (p.basename(bin).toLowerCase() === 'soffice.exe') {
      const com = p.join(p.dirname(bin), 'soffice.com');
      return exists(com) ? com : null;
    }
  }
  return bin;
}

const defaultRun: ProbeRunner = (file, args, timeoutMs) => {
  const r = spawnSync(file, args, { encoding: 'utf8', timeout: timeoutMs, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'], maxBuffer: 1024 * 1024 });
  const err = r.error as (NodeJS.ErrnoException & { code?: string }) | undefined;
  return {
    status: r.status,
    stdout: typeof r.stdout === 'string' ? r.stdout : '',
    timedOut: err?.code === 'ETIMEDOUT',
    ...(err && err.code !== 'ETIMEDOUT' ? { error: err.code ?? err.message } : {}),
  };
};

const PROBE_TIMEOUT_MS = 10_000;
const CACHE_VERSION_MS = 10 * 60_000;
const CACHE_UNREADABLE_MS = 60_000;
const probeCache = new Map<string, { at: number; value: VersionProbe }>();
export function resetLibreOfficeVersionCache(): void {
  probeCache.clear();
}

/** Ask the selected binary for its version. Cached per binary path (a readable answer for 10 minutes, a failure for one). */
export function probeLibreOfficeVersion(bin: string, o: VersionProbeOptions = {}): VersionProbe {
  const platform = o.platform ?? process.platform;
  const now = (o.now ?? Date.now)();
  const key = `${platform}|${bin}`;
  const hit = probeCache.get(key);
  if (hit && now - hit.at < (hit.value.ok ? CACHE_VERSION_MS : CACHE_UNREADABLE_MS)) return hit.value;
  const value = probeUncached(bin, platform, o);
  probeCache.set(key, { at: now, value });
  return value;
}

function probeUncached(bin: string, platform: NodeJS.Platform, o: VersionProbeOptions): VersionProbe {
  const exists = o.exists ?? ((p: string) => fs.existsSync(p));
  const file = versionProbeFile(bin, platform, exists);
  if (file === null) return { ok: false, reason: 'no soffice.com beside soffice.exe' };
  let r: ProbeRun;
  try {
    r = (o.run ?? defaultRun)(file, ['--version'], o.timeoutMs ?? PROBE_TIMEOUT_MS);
  } catch (err: unknown) {
    return { ok: false, reason: err instanceof Error ? err.message : String(err) };
  }
  if (r.timedOut) return { ok: false, reason: 'the version query timed out' };
  if (r.error) return { ok: false, reason: `the version query could not start (${r.error})` };
  if (r.status !== 0) return { ok: false, reason: `the version query exited with code ${String(r.status)}` };
  const version = parseLibreOfficeVersion(r.stdout);
  return version ? { ok: true, version } : { ok: false, reason: 'the version query printed no recognisable version' };
}

type RequestError = Error & { code: string; remedy: string; failureClass: FailureClass };

function requestError(code: string, message: string, remedy: string): RequestError {
  return Object.assign(codedError(code, message, remedy), { failureClass: 'request' as FailureClass });
}

const UPGRADE_REMEDY = 'Upgrade LibreOffice, or on Windows use --renderer word, or render without --show-markup.';

/**
 * Throw `E_DEPENDENCY` unless the selected LibreOffice is at or above the tested minimum. An unreadable version is
 * not "too old": the message says what is known, which is that the version could not be determined.
 */
export function assertLibreOfficeMarkupSupport(bin: string, o: VersionProbeOptions = {}): LibreOfficeVersion {
  const probe = probeLibreOfficeVersion(bin, o);
  if (!probe.ok) {
    throw requestError(
      'E_DEPENDENCY',
      `E_DEPENDENCY: ASTra could not determine whether this LibreOffice supports --show-markup (the version could not be read: ${probe.reason}).`,
      UPGRADE_REMEDY
    );
  }
  if (compareVersion(probe.version, minimumVersion()) < 0) {
    throw requestError(
      'E_DEPENDENCY',
      `E_DEPENDENCY: LibreOffice ${probe.version.text} has not been tested for showing comments in the margin; ASTra needs ${MIN_LIBREOFFICE_FOR_MARKUP} or later for --show-markup.`,
      UPGRADE_REMEDY
    );
  }
  return probe.version;
}

// ---------------------------------------------------------------------------------------------------------
// A failed markup conversion
// ---------------------------------------------------------------------------------------------------------

export type PlainRunOutcome = { ok: true } | { ok: false; error: string };

/**
 * The error for a LibreOffice conversion that failed with the margin-comments option. `plain` is one diagnostic
 * conversion without the option (its output is discarded). The message says what the two runs showed in this call:
 *  - plain succeeded: LibreOffice failed with the option and rendered without it. Not a diagnosis of the cause;
 *  - plain failed too: the cause is not determined, and both errors are kept.
 */
export function markupFailureError(markupError: string, plain: PlainRunOutcome): RequestError {
  const e1 = markupError.slice(0, 500);
  if (plain.ok) {
    return requestError(
      'E_RENDER',
      `E_RENDER: LibreOffice failed to render this document with comments in the margin, and rendered it without. The error with the option was: ${e1}`,
      'Render without --show-markup, or use --renderer word on Windows.'
    );
  }
  return requestError(
    'E_RENDER',
    `E_RENDER: LibreOffice failed with comments in the margin (${e1}) and also without (${plain.error.slice(0, 500)}); the cause was not determined.`,
    'Repair the document in Word (accept revisions, fix charts and notes) and retry, and check the LibreOffice installation.'
  );
}

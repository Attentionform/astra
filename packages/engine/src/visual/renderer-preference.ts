/**
 * Which program turns a .docx into a PDF (and, separately, which program turns a PDF into page images).
 *
 * Pure logic, no I/O: parsing the user's choice, resolving flag > environment > default, and planning
 * which engines may run for a given machine. Everything here is unit-testable on any platform.
 *
 * Renderer (DOCX -> PDF):  auto | word | libreoffice      flag --renderer       env ASTRA_RENDERER
 * Rasterizer (PDF -> PNG): auto | windows | poppler        flag --pdf-to-image   env ASTRA_PDF_TO_IMAGE
 */

export type RendererPreference = 'auto' | 'word' | 'libreoffice';
export type RendererId = 'word' | 'libreoffice';
export const RENDERER_PREFERENCES: readonly RendererPreference[] = ['auto', 'word', 'libreoffice'];
export const RENDERER_ENV = 'ASTRA_RENDERER';

export type RasterizerPreference = 'auto' | 'windows' | 'poppler';
export type RasterizerId = 'windows' | 'poppler';
export const RASTERIZER_PREFERENCES: readonly RasterizerPreference[] = ['auto', 'windows', 'poppler'];
export const PDF_TO_IMAGE_ENV = 'ASTRA_PDF_TO_IMAGE';

/**
 * 0.5.2: Word is opt-in. `auto` resolves to LibreOffice on every platform, and Word runs only when the
 * user asks for it (`--renderer word` / `ASTRA_RENDERER=word`). Flip to `true` in a later release, on
 * evidence from real use, to let `auto` prefer Word on Windows (the Word-first planning below is built
 * and tested; this constant is the only switch).
 */
export const AUTO_PREFERS_WORD = false;

export type PreferenceSource = 'flag' | 'parameter' | typeof RENDERER_ENV | typeof PDF_TO_IMAGE_ENV;

export interface CodedError extends Error {
  code: string;
  remedy: string;
}

export function codedError(code: string, message: string, remedy: string): CodedError {
  return Object.assign(new Error(message), { code, remedy });
}

function dependency(message: string, remedy: string): CodedError {
  return codedError('E_DEPENDENCY', `E_DEPENDENCY: ${message}`, remedy);
}

function parseChoice<T extends string>(
  raw: unknown,
  valid: readonly T[],
  what: string,
  source: PreferenceSource,
  remedy: string
): T | undefined {
  if (raw === undefined || raw === null) return undefined;
  if (typeof raw !== 'string') {
    throw codedError('E_VALIDATION', `E_VALIDATION: ${source} ${what} must be one of ${valid.join(', ')}.`, remedy);
  }
  const v = raw.trim().toLowerCase();
  if (v === '') return undefined; // an empty value means "not set" (an empty environment variable)
  if ((valid as readonly string[]).includes(v)) return v as T;
  throw codedError(
    'E_VALIDATION',
    `E_VALIDATION: ${source} ${what} "${raw.length > 40 ? `${raw.slice(0, 40)}…` : raw}" is not one of ${valid.join(', ')}.`,
    remedy
  );
}

export function parseRendererPreference(raw: unknown, source: PreferenceSource = 'flag'): RendererPreference | undefined {
  return parseChoice(
    raw,
    RENDERER_PREFERENCES,
    'renderer',
    source,
    'Use auto (default), word (Windows only) or libreoffice.'
  );
}

/** Flag/parameter > ASTRA_RENDERER > auto. An invalid environment value fails loudly, naming the variable. */
export function effectiveRendererPreference(
  explicit: RendererPreference | string | undefined,
  env: NodeJS.ProcessEnv = process.env
): RendererPreference {
  const fromCall = parseRendererPreference(explicit, 'parameter');
  if (fromCall !== undefined) return fromCall;
  return parseRendererPreference(env[RENDERER_ENV], RENDERER_ENV) ?? 'auto';
}

export function parseRasterizerPreference(raw: unknown, source: PreferenceSource = 'flag'): RasterizerPreference | undefined {
  return parseChoice(
    raw,
    RASTERIZER_PREFERENCES,
    'pdf-to-image',
    source,
    'Use auto (default), windows (Windows only) or poppler.'
  );
}

export function effectiveRasterizerPreference(
  explicit: RasterizerPreference | string | undefined,
  env: NodeJS.ProcessEnv = process.env
): RasterizerPreference {
  const fromCall = parseRasterizerPreference(explicit, 'parameter');
  if (fromCall !== undefined) return fromCall;
  return parseRasterizerPreference(env[PDF_TO_IMAGE_ENV], PDF_TO_IMAGE_ENV) ?? 'auto';
}

export interface RendererProbe {
  platform: NodeJS.Platform;
  /** Word is installed and registered for automation (Windows only). */
  word: boolean;
  libreoffice: boolean;
  /** `auto` may start with Word (the circuit breaker clears this for a while after an infrastructure failure). */
  wordUsable?: boolean;
  /** Overrides AUTO_PREFERS_WORD (tests exercise the Word-first plan without waiting for the release that flips it). */
  autoPrefersWord?: boolean;
}

export interface RenderPlan {
  requested: RendererPreference;
  /** Engines to try, in order. Empty means nothing can render (the caller raises E_DEPENDENCY). */
  engines: RendererId[];
  /** Only `auto` may fall through from one engine to the next. */
  canFallBack: boolean;
}

/** Decide which renderers may run. Throws E_DEPENDENCY for an explicit `word` that cannot work. */
export function planRenderers(pref: RendererPreference, probe: RendererProbe): RenderPlan {
  const win = probe.platform === 'win32';
  if (pref === 'word') {
    if (!win) {
      throw dependency(
        'renderer "word" is only available on Windows.',
        'Use --renderer auto or --renderer libreoffice on this platform (ASTra inside WSL is a Linux process: it always uses LibreOffice).'
      );
    }
    if (!probe.word) {
      throw dependency(
        'Microsoft Word is not installed (no registered Word.Application, version 2010 or later).',
        'Install Word, or use --renderer auto / --renderer libreoffice.'
      );
    }
    return { requested: pref, engines: ['word'], canFallBack: false };
  }
  if (pref === 'libreoffice') {
    return { requested: pref, engines: probe.libreoffice ? ['libreoffice'] : [], canFallBack: false };
  }
  const engines: RendererId[] = [];
  // The breaker only skips an engine that has a fallback: when it is the only one, it stays eligible, so a
  // recovered Word is used again at once instead of the render reporting a missing program for minutes.
  if ((probe.autoPrefersWord ?? AUTO_PREFERS_WORD) && win && probe.word && (probe.wordUsable !== false || !probe.libreoffice)) engines.push('word');
  if (probe.libreoffice) engines.push('libreoffice');
  return { requested: pref, engines, canFallBack: engines.length > 1 };
}

export interface RasterizerProbe {
  platform: NodeJS.Platform;
  /** The Windows built-in PDF renderer can be attempted (Windows 10+, PowerShell present). */
  windowsRenderer: boolean;
  poppler: boolean;
  /** `auto` may start with the Windows renderer (the circuit breaker clears this for a while after an infrastructure failure). */
  windowsUsable?: boolean;
}

export interface RasterizerPlan {
  requested: RasterizerPreference;
  engines: RasterizerId[];
  canFallBack: boolean;
}

export function planRasterizers(pref: RasterizerPreference, probe: RasterizerProbe): RasterizerPlan {
  const win = probe.platform === 'win32';
  if (pref === 'windows') {
    if (!win) {
      throw dependency(
        'pdf-to-image "windows" is only available on Windows.',
        'Use --pdf-to-image auto or --pdf-to-image poppler on this platform.'
      );
    }
    if (!probe.windowsRenderer) {
      throw dependency(
        'The Windows PDF renderer is not available (it needs Windows 10 or later and Windows PowerShell).',
        'Use --pdf-to-image auto or --pdf-to-image poppler.'
      );
    }
    return { requested: pref, engines: ['windows'], canFallBack: false };
  }
  if (pref === 'poppler') {
    return { requested: pref, engines: probe.poppler ? ['poppler'] : [], canFallBack: false };
  }
  const engines: RasterizerId[] = [];
  // The breaker only skips an engine that has a fallback: when it is the only one, it stays eligible, so a
  // recovered Windows renderer is used again at once instead of the render reporting a missing pdftoppm.
  if (win && probe.windowsRenderer && (probe.windowsUsable !== false || !probe.poppler)) engines.push('windows');
  if (probe.poppler) engines.push('poppler');
  return { requested: pref, engines, canFallBack: engines.length > 1 };
}

/** Facts about which engine ran, in every render result. */
export interface RendererReport {
  requested: RendererPreference;
  used: RendererId;
  /** Word major version (14 = 2010, 15 = 2013, 16 = 2016/2019/2021/365) when Word rendered. */
  version?: number;
  /** Set when `auto` fell back: the engine that failed and a short reason. */
  fallbackFrom?: RendererId;
  fallbackReason?: string;
}

export interface RasterizerReport {
  requested: RasterizerPreference;
  used: RasterizerId;
  fallbackFrom?: RasterizerId;
  fallbackReason?: string;
}

/** A failure that says nothing about the document: the machine, a policy or the program itself is at fault. */
export type FailureClass = 'infrastructure' | 'document';

/**
 * Short-lived memory of an infrastructure failure so a long-lived process (the MCP server) does not pay a
 * full timeout on every call when Word or PowerShell is stuck. Used by `auto` only; explicit choices ignore it.
 */
export class CircuitBreaker {
  private openUntil = 0;
  constructor(private readonly cooldownMs: number, private readonly now: () => number = Date.now) {}
  isOpen(): boolean {
    return this.now() < this.openUntil;
  }
  trip(): void {
    this.openUntil = this.now() + this.cooldownMs;
  }
  reset(): void {
    this.openUntil = 0;
  }
}

export const BREAKER_COOLDOWN_MS = 5 * 60 * 1000;
export const wordBreaker = new CircuitBreaker(BREAKER_COOLDOWN_MS);
export const windowsRasterizerBreaker = new CircuitBreaker(BREAKER_COOLDOWN_MS);

export function resetRendererBreakersForTesting(): void {
  wordBreaker.reset();
  windowsRasterizerBreaker.reset();
}

/**
 * Headless PDF backend capability gate (Sprint 6 Track B).
 *
 * Pure Node (`fs`/`path` only).
 * Resolves `soffice`/`libreoffice` + `pdftoppm` via `$PATH` at CALL time
 * (never cached) so tests can hide binaries by mutating `process.env.PATH`.
 * On Windows, a bare name also matches `.exe`/`.com`. When PATH has no match,
 * the standard install folders of each platform are tried (LibreOffice on
 * Windows and macOS does not put itself on PATH). Set
 * `ASTRA_VISUAL_NO_FALLBACK_DIRS=1` to turn the folder fallback off.
 * Missing backend => thrown Error with `code: 'E_DEPENDENCY'` + install remedy.
 */

import * as fs from 'fs';
import * as path from 'path';
import {
  effectiveRasterizerPreference,
  effectiveRendererPreference,
  planRasterizers,
  planRenderers,
  type RasterizerId,
  type RendererId,
} from './renderer-preference';
import { findWord } from './word';
import { windowsRendererAvailable } from './windows-pdf';

export type VisualPdfBackend = 'word' | 'libreoffice' | 'none';

export interface VisualBackendStatus {
  /** A renderer and a rasterizer can both run under the effective preferences (flag-less: environment, then auto). */
  available: boolean;
  /** The renderer a default call would use first, or `none`. */
  backend: VisualPdfBackend;
  sofficePath?: string;
  pdftoppmPath?: string;
  /** Binary names that could not be resolved (for remedies). */
  missing: string[];
  /** What was found, independent of the preferences. */
  renderers?: {
    word: { available: boolean; exe?: string; version?: number };
    libreoffice: { available: boolean; path?: string };
  };
  rasterizers?: {
    windows: { available: boolean };
    poppler: { available: boolean; path?: string };
  };
  /** Engines a default call may use, in order. */
  plan?: { renderers: RendererId[]; rasterizers: RasterizerId[] };
}

export interface VisualBackendOverride {
  sofficeBin?: string;
  pdftoppmBin?: string;
  /** Override PATH for tests (defaults to process.env.PATH). An explicit value also turns the install-folder fallback off. */
  pathEnv?: string;
}

export interface ResolveOptions {
  /** Platform to resolve for (defaults to `process.platform`). */
  platform?: NodeJS.Platform;
  /** Windows executable extensions to try, `;`-separated (defaults to a safe `.exe;.com`). */
  pathExt?: string;
  /** Environment used for install-folder fallbacks (defaults to `process.env`). */
  env?: NodeJS.ProcessEnv;
  /** Try the platform's standard install folders when PATH has no match (default true, unless pathEnv is given). */
  fallbackDirs?: boolean;
  /** Replace the platform's standard install folders (tests use this to stay independent of what the host has installed). */
  installDirs?: string[];
}

export const VISUAL_DEPENDENCY_REMEDY =
  'PDF render backend missing. Install LibreOffice and poppler-utils (pdftoppm): ' +
  'Linux: sudo apt install libreoffice poppler-utils; ' +
  'macOS: brew install --cask libreoffice && brew install poppler; ' +
  'Windows: winget install TheDocumentFoundation.LibreOffice (Microsoft Word can render instead with --renderer word, ' +
  'and Windows draws the page images itself; Poppler is only needed as a fallback: scoop install poppler, or choco install poppler, ' +
  'with pdftoppm.exe on PATH). ' +
  'Alternatively inspect the XML/DOM without images (telemetry-only mode).';

export function createDependencyError(missing: string[]): Error & { code: string; remedy: string } {
  const which = missing.length > 0 ? missing.join(', ') : 'soffice, pdftoppm';
  return Object.assign(new Error(`E_DEPENDENCY: headless visual backend unavailable (missing: ${which})`), {
    code: 'E_DEPENDENCY',
    remedy: VISUAL_DEPENDENCY_REMEDY,
  });
}

/** Executable extensions tried for a bare name on Windows. `.exe` first; `.cmd`/`.bat` are excluded because `execFile` cannot spawn them. */
function windowsExtensions(pathExt?: string): string[] {
  const requested: string[] = (pathExt ?? '.EXE;.COM')
    .split(';')
    .map((e) => e.trim().toLowerCase())
    .filter((e) => e === '.exe' || e === '.com');
  const ordered = ['.exe', '.com'].filter((e) => requested.includes(e));
  return ordered.length > 0 ? ordered : ['.exe', '.com'];
}

function isExecutableFile(candidate: string): boolean {
  try {
    fs.accessSync(candidate, fs.constants.X_OK);
    return fs.statSync(candidate).isFile();
  } catch {
    return false;
  }
}

function candidateNames(name: string, platform: NodeJS.Platform, pathExt?: string): string[] {
  if (platform !== 'win32') return [name];
  const lower = name.toLowerCase();
  if (lower.endsWith('.exe') || lower.endsWith('.com')) return [name];
  return windowsExtensions(pathExt).map((ext) => name + ext);
}

/**
 * Standard install folders per platform and binary. Used only when PATH has
 * no match: LibreOffice for Windows and macOS keeps its binaries inside the
 * application folder and does not add it to PATH.
 */
export function installFolders(
  name: string,
  platform: NodeJS.Platform = process.platform,
  env: NodeJS.ProcessEnv = process.env
): string[] {
  const bare = name.toLowerCase().replace(/\.(exe|com)$/, '');
  const isOffice = bare === 'soffice' || bare === 'libreoffice';
  const isPoppler = bare === 'pdftoppm';
  const dirs: string[] = [];
  if (platform === 'win32') {
    const programFiles = env.ProgramFiles ?? 'C:\\Program Files';
    const programFilesX86 = env['ProgramFiles(x86)'] ?? 'C:\\Program Files (x86)';
    const localAppData = env.LOCALAPPDATA;
    const userProfile = env.USERPROFILE;
    if (isOffice) {
      dirs.push(path.join(programFiles, 'LibreOffice', 'program'));
      dirs.push(path.join(programFilesX86, 'LibreOffice', 'program'));
      if (localAppData) dirs.push(path.join(localAppData, 'Programs', 'LibreOffice', 'program'));
    }
    if (isPoppler) {
      if (userProfile) dirs.push(path.join(userProfile, 'scoop', 'shims'));
      dirs.push(path.join(env.ProgramData ?? 'C:\\ProgramData', 'chocolatey', 'bin'));
      if (localAppData) dirs.push(path.join(localAppData, 'Microsoft', 'WinGet', 'Links'));
    }
  } else if (platform === 'darwin') {
    const home = env.HOME;
    if (isOffice) {
      dirs.push('/Applications/LibreOffice.app/Contents/MacOS');
      if (home) dirs.push(path.join(home, 'Applications', 'LibreOffice.app', 'Contents', 'MacOS'));
    }
    // Homebrew (Apple silicon, Intel) and MacPorts, for both binaries.
    dirs.push('/opt/homebrew/bin', '/usr/local/bin', '/opt/local/bin');
  }
  // Linux: package managers put binaries on PATH, so there is no fallback.
  return dirs;
}

/**
 * Resolve `name` against `$PATH` (plus absolute/relative paths verbatim).
 * Returns the absolute path or null. Respects `pathEnv` override for tests.
 * On Windows a bare name also matches `name.exe` and `name.com`.
 */
export function findOnPath(name: string, pathEnv?: string, options: ResolveOptions = {}): string | null {
  try {
    if (!name || name.trim() === '') return null;
    const trimmed = name.trim();
    const platform = options.platform ?? process.platform;
    const env = options.env ?? process.env;
    const pathExt = options.pathExt ?? (platform === 'win32' ? env.PATHEXT : undefined);
    const names = candidateNames(trimmed, platform, pathExt);
    // Explicit path (contains a slash): check verbatim, no PATH search.
    if (trimmed.includes('/') || trimmed.includes('\\')) {
      for (const n of names) {
        const abs = path.resolve(n);
        if (isExecutableFile(abs)) return abs;
      }
      return null;
    }
    const pathValue = pathEnv !== undefined ? pathEnv : (env.PATH ?? env.Path ?? '');
    const dirs = pathValue
      .split(path.delimiter)
      .map((d) => d.trim().replace(/^"(.*)"$/, '$1'))
      .filter((d) => d.length > 0);
    for (const dir of dirs) {
      for (const n of names) {
        const candidate = path.resolve(dir, n);
        if (isExecutableFile(candidate)) return candidate;
      }
    }
    const useFallback =
      options.fallbackDirs ?? (pathEnv === undefined && env.ASTRA_VISUAL_NO_FALLBACK_DIRS !== '1');
    if (useFallback) {
      for (const dir of options.installDirs ?? installFolders(trimmed, platform, env)) {
        for (const n of names) {
          const candidate = path.resolve(dir, n);
          if (isExecutableFile(candidate)) return candidate;
        }
      }
    }
    return null;
  } catch {
    return null;
  }
}

export function resolveSoffice(override?: string, pathEnv?: string): string | null {
  if (override && override.trim()) {
    const direct = findOnPath(override.trim(), pathEnv);
    if (direct) return direct;
    // Explicit override that does not resolve => missing (do not fall back:
    // callers passing a bogus path intend to simulate absence).
    return null;
  }
  return findOnPath('soffice', pathEnv) ?? findOnPath('libreoffice', pathEnv);
}

/** Resolve Poppler's `pdftoppm` (explicit override, else PATH and standard folders). */
export function resolvePoppler(override?: VisualBackendOverride): string | null {
  const pathEnv = override?.pathEnv;
  if (override?.pdftoppmBin && override.pdftoppmBin.trim()) {
    return findOnPath(override.pdftoppmBin.trim(), pathEnv);
  }
  return findOnPath('pdftoppm', pathEnv);
}

/**
 * Capability probe. Never throws. `available` means a renderer AND a rasterizer can both run under the
 * effective preferences (`ASTRA_RENDERER`, `ASTRA_PDF_TO_IMAGE`, else auto); an invalid environment value is
 * ignored here (the render call reports it loudly). `sofficePath` / `pdftoppmPath` are what was found.
 */
export function hasPdfBackend(override?: VisualBackendOverride): VisualBackendStatus {
  const pathEnv = override?.pathEnv;
  const sofficePath = resolveSoffice(override?.sofficeBin, pathEnv);
  const pdftoppmPath = resolvePoppler(override);
  const platform = process.platform;
  const word = platform === 'win32' ? findWord() : null;
  // An explicit pdftoppm binary means "use that Poppler"; the Windows renderer is then not a substitute.
  const windowsRenderer = !(override?.pdftoppmBin && override.pdftoppmBin.trim()) && windowsRendererAvailable();
  let rendererPref: ReturnType<typeof effectiveRendererPreference> = 'auto';
  let rasterPref: ReturnType<typeof effectiveRasterizerPreference> = 'auto';
  try {
    rendererPref = effectiveRendererPreference(undefined);
  } catch {
    rendererPref = 'auto';
  }
  try {
    rasterPref = effectiveRasterizerPreference(undefined);
  } catch {
    rasterPref = 'auto';
  }
  let renderers: RendererId[] = [];
  let rasterizers: RasterizerId[] = [];
  try {
    renderers = planRenderers(rendererPref, { platform, word: word !== null, libreoffice: sofficePath !== null }).engines;
  } catch {
    renderers = [];
  }
  try {
    rasterizers = planRasterizers(rasterPref, { platform, windowsRenderer, poppler: pdftoppmPath !== null }).engines;
  } catch {
    rasterizers = [];
  }
  const missing: string[] = [];
  if (renderers.length === 0) missing.push('soffice');
  if (rasterizers.length === 0) missing.push('pdftoppm');
  const detail = {
    renderers: {
      word: { available: word !== null, ...(word ? { exe: word.exe, version: word.version } : {}) },
      libreoffice: { available: sofficePath !== null, ...(sofficePath ? { path: sofficePath } : {}) },
    },
    rasterizers: {
      windows: { available: windowsRenderer },
      poppler: { available: pdftoppmPath !== null, ...(pdftoppmPath ? { path: pdftoppmPath } : {}) },
    },
    plan: { renderers, rasterizers },
  };
  if (missing.length === 0) {
    return {
      available: true,
      backend: renderers[0],
      ...(sofficePath ? { sofficePath } : {}),
      ...(pdftoppmPath ? { pdftoppmPath } : {}),
      missing: [],
      ...detail,
    };
  }
  return {
    available: false,
    backend: 'none',
    ...(sofficePath ? { sofficePath } : {}),
    ...(pdftoppmPath ? { pdftoppmPath } : {}),
    missing,
    ...detail,
  };
}

export const hasVisualPdfBackend = hasPdfBackend;
export const getVisualBackendStatus = hasPdfBackend;

export function hasLibreOffice(override?: VisualBackendOverride): boolean {
  return Boolean(resolveSoffice(override?.sofficeBin, override?.pathEnv));
}

export function hasPoppler(override?: VisualBackendOverride): boolean {
  return Boolean(resolvePoppler(override));
}

/**
 * Throwing gate kept for callers that need the two binaries themselves (LibreOffice and Poppler). The render
 * paths use the renderer and rasterizer plans instead, which also allow Word and the Windows PDF renderer.
 */
export function assertPdfBackend(override?: VisualBackendOverride): { sofficePath: string; pdftoppmPath: string } {
  const sofficePath = resolveSoffice(override?.sofficeBin, override?.pathEnv);
  const pdftoppmPath = resolvePoppler(override);
  const missing: string[] = [];
  if (!sofficePath) missing.push('soffice');
  if (!pdftoppmPath) missing.push('pdftoppm');
  if (missing.length > 0 || !sofficePath || !pdftoppmPath) throw createDependencyError(missing);
  return { sofficePath, pdftoppmPath };
}

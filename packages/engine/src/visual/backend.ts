/**
 * Headless PDF backend capability gate (Sprint 6 Track B).
 *
 * Pure Node (`fs`/`path` only) — both trees keep IDENTICAL copies.
 * Resolves `soffice`/`libreoffice` + `pdftoppm` via `$PATH` at CALL time
 * (never cached) so tests can hide binaries by mutating `process.env.PATH`.
 * Missing backend => thrown Error with `code: 'E_DEPENDENCY'` + install remedy.
 */

import * as fs from 'fs';
import * as path from 'path';

export type VisualPdfBackend = 'libreoffice' | 'none';

export interface VisualBackendStatus {
  available: boolean;
  backend: VisualPdfBackend;
  sofficePath?: string;
  pdftoppmPath?: string;
  /** Binary names that could not be resolved (for remedies). */
  missing: string[];
}

export interface VisualBackendOverride {
  sofficeBin?: string;
  pdftoppmBin?: string;
  /** Override PATH for tests (defaults to process.env.PATH). */
  pathEnv?: string;
}

export const VISUAL_DEPENDENCY_REMEDY =
  'PDF render backend missing. Install LibreOffice and poppler-utils (pdftoppm): ' +
  'sudo apt install libreoffice poppler-utils, or brew install libreoffice poppler. ' +
  'Alternatively inspect the XML/DOM without images (telemetry-only mode).';

export function createDependencyError(missing: string[]): Error & { code: string; remedy: string } {
  const which = missing.length > 0 ? missing.join(', ') : 'soffice, pdftoppm';
  return Object.assign(new Error(`E_DEPENDENCY: headless visual backend unavailable (missing: ${which})`), {
    code: 'E_DEPENDENCY',
    remedy: VISUAL_DEPENDENCY_REMEDY,
  });
}

/**
 * Resolve `name` against `$PATH` (plus absolute/relative paths verbatim).
 * Returns the absolute path or null. Respects `pathEnv` override for tests.
 */
export function findOnPath(name: string, pathEnv?: string): string | null {
  try {
    if (!name || name.trim() === '') return null;
    const trimmed = name.trim();
    // Explicit path (contains a slash): check verbatim, no PATH search.
    if (trimmed.includes('/') || trimmed.includes('\\')) {
      try {
        const abs = path.resolve(trimmed);
        fs.accessSync(abs, fs.constants.X_OK);
        const stat = fs.statSync(abs);
        if (stat.isFile()) return abs;
        return null;
      } catch {
        return null;
      }
    }
    const pathValue = pathEnv !== undefined ? pathEnv : (process.env.PATH ?? '');
    if (!pathValue) return null;
    const dirs = pathValue.split(path.delimiter).filter((d) => d.length > 0);
    for (const dir of dirs) {
      try {
        const candidate = path.join(dir, trimmed);
        fs.accessSync(candidate, fs.constants.X_OK);
        const stat = fs.statSync(candidate);
        if (stat.isFile()) return candidate;
      } catch {
        continue;
      }
    }
    return null;
  } catch {
    return null;
  }
}

function resolveSoffice(override?: string, pathEnv?: string): string | null {
  if (override && override.trim()) {
    const direct = findOnPath(override.trim(), pathEnv);
    if (direct) return direct;
    // Explicit override that does not resolve => missing (do not fall back:
    // callers passing a bogus path intend to simulate absence).
    return null;
  }
  return findOnPath('soffice', pathEnv) ?? findOnPath('libreoffice', pathEnv);
}

/**
 * Capability probe. Never throws — returns `{available:false}` when any
 * binary is missing. Call `assertPdfBackend` for the throwing gate.
 */
export function hasPdfBackend(override?: VisualBackendOverride): VisualBackendStatus {
  const pathEnv = override?.pathEnv;
  const sofficePath = resolveSoffice(override?.sofficeBin, pathEnv);
  let pdftoppmPath: string | null = null;
  if (override?.pdftoppmBin && override.pdftoppmBin.trim()) {
    pdftoppmPath = findOnPath(override.pdftoppmBin.trim(), pathEnv);
  } else {
    pdftoppmPath = findOnPath('pdftoppm', pathEnv);
  }
  const missing: string[] = [];
  if (!sofficePath) missing.push('soffice');
  if (!pdftoppmPath) missing.push('pdftoppm');
  if (missing.length === 0) {
    return { available: true, backend: 'libreoffice', sofficePath: sofficePath as string, pdftoppmPath: pdftoppmPath as string, missing: [] };
  }
  return { available: false, backend: 'none', missing };
}

export const hasVisualPdfBackend = hasPdfBackend;
export const getVisualBackendStatus = hasPdfBackend;

export function hasLibreOffice(override?: VisualBackendOverride): boolean {
  return Boolean(resolveSoffice(override?.sofficeBin, override?.pathEnv));
}

export function hasPoppler(override?: VisualBackendOverride): boolean {
  const pathEnv = override?.pathEnv;
  if (override?.pdftoppmBin && override.pdftoppmBin.trim()) {
    return Boolean(findOnPath(override.pdftoppmBin.trim(), pathEnv));
  }
  return Boolean(findOnPath('pdftoppm', pathEnv));
}

export function assertPdfBackend(override?: VisualBackendOverride): { sofficePath: string; pdftoppmPath: string } {
  const status = hasPdfBackend(override);
  if (!status.available) {
    throw createDependencyError(status.missing);
  }
  return { sofficePath: status.sofficePath as string, pdftoppmPath: status.pdftoppmPath as string };
}

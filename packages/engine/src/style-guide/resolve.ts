/**
 * Named + file style guide resolution.
 *
 * A guide ref is either a file path (`./brand.md`, `/workspace/g.md`) or a
 * bundled name (`report`) that resolves under `BUNDLED_GUIDE_DIR`
 * (`packages/engine/templates/guides/<name>.md`, pinned by the parser
 * module). Bundled markdown files themselves ship separately — until they
 * do, named resolution fails cleanly with a coded error, never silently.
 *
 * Jail rule (mirrors image ingestion): relative file refs resolve under
 * `allowedDir` when one is configured, else the process working directory;
 * an absolute ref alongside `allowedDir` must sit inside it. Remote URLs
 * fail closed. Parse diagnostics ride along — a resolved guide still
 * builds documents when it parsed with warnings.
 */
import fs from 'fs';
import path from 'path';
import {
  bundledGuidePath,
  isFileGuideRef,
  parseStyleGuide,
} from './parser';
import type { StyleGuideConfig } from './types';
import { canonicalizeForAsset, isPathInsideDir } from '../sandbox/paths';
import { toError } from '../errors';

export interface ResolveStyleGuideOptions {
  /**
   * Jail for relative file refs (same trust as batch image `allowedDir`).
   * Omitted = relative refs anchor at the process working directory.
   */
  allowedDir?: string;
}

export interface ResolvedStyleGuide {
  config: StyleGuideConfig;
  /** Parser diagnostics for the resolved file (defaulted sections, hex fixes). */
  warnings: string[];
  /** Canonical absolute path of the file the guide was read from. */
  filePath: string;
}

function hasScheme(ref: string): boolean {
  return /^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(ref);
}

function failValidation(message: string, remedy: string): never {
  throw toError('E_VALIDATION', message, remedy);
}

/** Absolute path candidates for a bundled guide name, first hit wins. */
export function bundledGuideCandidates(name: string): string[] {
  const rel = bundledGuidePath(name);
  const base = `${rel.split('/').pop()}`;
  const out = [
    path.resolve(process.cwd(), rel),
    path.join(__dirname, '..', 'templates', 'guides', base),
    path.join(__dirname, '..', '..', 'templates', 'guides', base),
    path.join(__dirname, 'templates', 'guides', base),
  ];
  const installed = installedEngineGuide(base);
  if (installed) out.push(installed);
  return out;
}

/**
 * Locates the guide inside an installed engine copy for bundled callers
 * whose own folder is not the engine folder. Walks up from this module
 * looking for a sibling engine install holding the shipped copy.
 */
function installedEngineGuide(base: string): string | undefined {
  const dir = installedEngineDir();
  if (!dir) return undefined;
  return path.join(dir, base);
}

function installedEngineDir(): string | undefined {
  try {
    let dir = __dirname;
    for (let i = 0; i < 6; i++) {
      const cand = path.join(
        dir,
        'node_modules',
        '@attentionform',
        'astra-engine',
        'dist',
        'templates',
        'guides',
      );
      if (fs.existsSync(cand)) return cand;
      const parent = path.dirname(dir);
      if (parent === dir) break;
      dir = parent;
    }
  } catch {
    /* best-effort only */
  }
  return undefined;
}

/** Bundled names currently installed on disk (empty until guides ship). */
export function listBundledGuides(): string[] {
  const out = new Set<string>();
  const dirs = [
    path.resolve(process.cwd(), 'packages/engine/templates/guides'),
    path.join(__dirname, '..', 'templates', 'guides'),
    path.join(__dirname, '..', '..', 'templates', 'guides'),
    path.join(__dirname, 'templates', 'guides'),
  ];
  const installedDir = installedEngineDir();
  if (installedDir) dirs.push(installedDir);
  for (const candidate of dirs) {
    try {
      if (!fs.existsSync(candidate)) continue;
      for (const entry of fs.readdirSync(candidate)) {
        if (entry.toLowerCase().endsWith('.md')) {
          out.add(entry.replace(/\.md$/i, ''));
        }
      }
    } catch {
      /* best-effort listing */
    }
  }
  return [...out].sort();
}

function resolveBundledGuide(name: string): ResolvedStyleGuide {
  const trimmed = String(name ?? '').trim();
  for (const candidate of bundledGuideCandidates(trimmed)) {
    try {
      if (!fs.existsSync(candidate)) continue;
      const markdown = fs.readFileSync(candidate, 'utf8');
      const { config, warnings } = parseStyleGuide(markdown);
      return { config, warnings, filePath: canonicalizeForAsset(candidate) };
    } catch (err) {
      if (err instanceof Error && /E_VALIDATION|E_SANDBOX|E_IO/.test(err.message)) {
        throw err;
      }
      failValidation(
        `style guide '${trimmed}' is not readable (${err instanceof Error ? err.message : String(err)})`,
        `check file permissions for the bundled guide, or pass a file path to a style-guide.md instead`,
      );
    }
  }
  const available = listBundledGuides();
  failValidation(
    `unknown style guide '${trimmed}': no bundled guide found` +
      (available.length > 0 ? ` (available: ${available.join(', ')})` : ' (no bundled guides installed)'),
    available.length > 0
      ? `pass one of: ${available.join(', ')}, or a file path to a style-guide.md`
      : `pass a file path to a style-guide.md holding the guide markdown`,
  );
}

function resolveFileGuide(ref: string, opts?: ResolveStyleGuideOptions): ResolvedStyleGuide {
  const trimmed = String(ref ?? '').trim();
  if (hasScheme(trimmed)) {
    failValidation(
      `style guide '${trimmed}' is a remote URL; guides must be local files or bundled names`,
      'download the guide markdown into the workspace and pass its file path instead',
    );
  }
  const allowed = opts?.allowedDir?.trim() ? opts.allowedDir.trim() : undefined;
  const base = allowed ?? process.cwd();
  const candidate = path.isAbsolute(trimmed)
    ? path.resolve(trimmed)
    : path.resolve(base, trimmed);
  if (allowed) {
    const canonicalAllowed = canonicalizeForAsset(path.resolve(allowed));
    const canonicalCandidate = canonicalizeForAsset(candidate);
    if (
      !isPathInsideDir(canonicalCandidate, canonicalAllowed) &&
      canonicalCandidate !== canonicalAllowed
    ) {
      throw toError(
        'E_SANDBOX',
        `style guide '${trimmed}' resolves outside the allowed directory`,
        `keep guide files inside the allowed directory (${canonicalAllowed}) and use a relative path`,
      );
    }
  }
  let markdown: string;
  try {
    markdown = fs.readFileSync(candidate, 'utf8');
  } catch {
    failValidation(
      `style guide file '${trimmed}' not found (resolved '${candidate}')`,
      'pass a readable .md file path, or a bundled guide name',
    );
  }
  const { config, warnings } = parseStyleGuide(markdown);
  return { config, warnings, filePath: canonicalizeForAsset(candidate) };
}

/**
 * Resolves a guide ref to a parsed config. File-backed refs (`.md`
 * suffix, path separators) read from disk under the jail; anything else
 * is a bundled name. Unknown names and missing files fail closed with a
 * coded `E_VALIDATION` (never null, never silent defaults).
 */
export function resolveStyleGuide(
  ref: string,
  opts?: ResolveStyleGuideOptions,
): ResolvedStyleGuide {
  if (typeof ref !== 'string' || ref.trim() === '') {
    failValidation(
      'style guide ref must be a non-empty string',
      'pass a bundled guide name or a file path to a style-guide.md',
    );
  }
  const trimmed = ref.trim();
  if (hasScheme(trimmed) || isFileGuideRef(trimmed)) {
    return resolveFileGuide(trimmed, opts);
  }
  return resolveBundledGuide(trimmed);
}

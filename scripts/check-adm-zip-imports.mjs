#!/usr/bin/env node
/**
 * L6 BUILD gate: fail the build on raw adm-zip imports.
 *
 * Only packages-engine-src-adm-zip-safe.ts (the CVE-2026-76845 hardened
 * wrapper) may import raw adm-zip. Any other production import under
 * packages src trees, backend, electron, or renderer reintroduces disk
 * extraction outside the instance hardening and must fail the build
 * loudly (replacing the dep is out of scope; this gate is the
 * containment). Static `from`/`require`, subpaths (`adm-zip/*`),
 * side-effect imports, and dynamic `import()` of `adm-zip` all fail —
 * including backtick-interpolated specifiers (dynamic import of
 * `adm-zip${x}`), whose specifier START is still statically `adm-zip`.
 *
 * Scans real import and require forms only (prose mentions in comments are
 * stripped first so docs stay grep-clean without tripping the gate).
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, '..');
const srcRoots = [
  'packages/engine/src',
  'packages/agent/src',
  'packages/mcp-server/src',
  'packages/cli/src',
  'backend',
  'electron',
  'renderer',
];
const ALLOW = new Set([path.join('packages', 'engine', 'src', 'adm-zip-safe.ts')]);

const SCAN_EXTS = new Set(['.ts', '.tsx', '.js', '.jsx', '.mjs', '.cjs']);
const SKIP_DIRS = new Set(['node_modules', 'dist']);

function walk(dir) {
  const out = [];
  let entries;
  try {
    entries = fs.readdirSync(dir, { withFileTypes: true });
  } catch {
    return out;
  }
  for (const e of entries) {
    const full = path.join(dir, e.name);
    if (e.isDirectory()) {
      if (SKIP_DIRS.has(e.name)) continue;
      out.push(...walk(full));
    } else if (e.isFile() && !full.endsWith('.d.ts')) {
      if (SCAN_EXTS.has(path.extname(full))) out.push(full);
    }
  }
  return out;
}

// Quote/backtick opens an adm-zip specifier; the lookahead accepts the plain
// close quote, a `/subpath`, or `${` interpolation (backtick templates whose
// specifier START is adm-zip, e.g. dynamic import of `adm-zip${x}`). A
// trailing `-suffix` (different package, e.g. 'adm-zip-evil') still does NOT
// match.
const importRe = /(?:from\s+|import\s*\(\s*|require\(\s*|import\s+)(['"`])adm-zip(?=[/'"`$])/;
const failures = [];
for (const rel of srcRoots) {
  const abs = path.join(root, rel);
  for (const file of walk(abs)) {
    const relFile = path.relative(root, file);
    if (ALLOW.has(relFile)) continue;
    const raw = fs.readFileSync(file, 'utf8');
    const stripped = raw
      .replace(/\/\*[\s\S]*?\*\//g, '')
      .replace(/(^|[^:"'`])\/\/.*$/gm, '$1');
    if (importRe.test(stripped)) failures.push(relFile);
  }
}
if (failures.length > 0) {
  console.error('[adm-zip-gate] BUILD REFUSED: raw adm-zip import(s) outside adm-zip-safe.ts:');
  for (const f of failures) console.error(`  - ${f}`);
  console.error("Remedy: import from './adm-zip-safe' (engine) instead of 'adm-zip'.");
  process.exit(1);
}
console.log('[adm-zip-gate] ok: no raw adm-zip imports outside the hardened wrapper.');

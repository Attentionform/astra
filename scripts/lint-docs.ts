/**
 * Docs gate: user-facing markdown plus registry descriptions stay free of
 * unfinished markers, legacy env references, internal terms, and stale labels.
 *
 * Usage: npx tsx scripts/lint-docs.ts
 *
 * Scans: README.md, docs/*.md, packages asterisk README.md, registries JSON,
 * integrations README plus JSON. Prints file:line on hit, exits non-zero.
 *
 * Prose-only check for internal terms: fenced blocks plus inline code are
 * stripped first, so file paths plus code samples do not count. Markers plus
 * the legacy env literal are checked on full text (even inside code).
 *
 * Allowlist: scripts/lint-docs-allowlist.json. Each entry names a file
 * fragment plus a regex; a hit is skipped when both match.
 */

import * as fs from 'fs';
import * as path from 'path';

type AllowEntry = { file: string; match: string; why?: string };

function repoRoot(): string {
  return process.cwd();
}

function listFiles(root: string): string[] {
  const out: string[] = [];
  const pushIfExists = (p: string) => {
    if (fs.existsSync(p) && fs.statSync(p).isFile()) out.push(p);
  };
  pushIfExists(path.join(root, 'README.md'));
  const docsDir = path.join(root, 'docs');
  if (fs.existsSync(docsDir)) {
    for (const f of fs.readdirSync(docsDir)) {
      if (f.endsWith('.md')) out.push(path.join(docsDir, f));
    }
  }
  const packagesDir = path.join(root, 'packages');
  if (fs.existsSync(packagesDir)) {
    for (const pkg of fs.readdirSync(packagesDir)) {
      const cand = path.join(packagesDir, pkg, 'README.md');
      pushIfExists(cand);
    }
  }
  const regDir = path.join(root, 'registries');
  if (fs.existsSync(regDir)) {
    for (const f of fs.readdirSync(regDir)) {
      if (f.endsWith('.json')) out.push(path.join(regDir, f));
    }
  }
  const integDir = path.join(root, 'integrations');
  const walk = (dir: string) => {
    if (!fs.existsSync(dir)) return;
    for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
      const full = path.join(dir, e.name);
      if (e.isDirectory()) walk(full);
      else if (e.name === 'README.md' || e.name.endsWith('.json')) out.push(full);
    }
  };
  walk(integDir);
  return [...new Set(out)].sort();
}

function loadAllowlist(root: string): AllowEntry[] {
  const p = path.join(root, 'scripts', 'lint-docs-allowlist.json');
  if (!fs.existsSync(p)) return [];
  try {
    const raw = JSON.parse(fs.readFileSync(p, 'utf8')) as { entries?: AllowEntry[] };
    if (!Array.isArray(raw.entries)) return [];
    return raw.entries.filter((e) => typeof e?.file === 'string' && typeof e?.match === 'string');
  } catch {
    return [];
  }
}

function isAllowed(rel: string, line: string, list: AllowEntry[]): boolean {
  for (const e of list) {
    if (!rel.includes(e.file) && !e.file.includes(rel)) {
      if (rel !== e.file && !rel.endsWith(e.file)) continue;
    }
    let re: RegExp | null = null;
    try {
      re = new RegExp(e.match, 'i');
    } catch {
      continue;
    }
    if (re.test(line)) return true;
  }
  return false;
}

function stripCode(text: string): string {
  const noFence = text.replace(/```[\s\S]*?```/g, (m) => m.replace(/[^\n]/g, ' '));
  return noFence.replace(/`[^`\n]*`/g, (m) => ' '.repeat(m.length));
}

function buildPatterns(): { markers: RegExp[]; env: RegExp[]; terms: RegExp[]; stale: RegExp[] } {
  const MARK_A = 'TO-' + 'VERIFY';
  const MARK_B = 'TO' + 'DO';
  const MARK_C = 'FIX' + 'ME';
  const ENV_A = 'ASTRA_' + 'ALLOW' + 'ED_DIR';
  const T_A = 'lo' + 'op';
  const T_B = 'revie' + 'wer';
  const T_C = 'tes' + 'ter';
  const T_D = 'suba' + 'gent';
  const T_E = 'mu' + 'se';
  const T_F = 'antigra' + 'vity';
  const T_G = 'hos' + 'tile';
  const T_H = 'miles' + 'tone';
  const T_I = 'spr' + 'int';
  const T_J = 'pha' + 'se';
  const T_K = 'wa' + 've';

  const esc = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const word = (s: string) => new RegExp('\\b' + esc(s) + '\\b', 'i');

  const markers = [new RegExp(esc(MARK_A)), word(MARK_B), word(MARK_C)];
  const env = [new RegExp(esc(ENV_A))];
  const terms = [T_A, T_B, T_C, T_D, T_E, T_F, T_G, T_H, T_I, T_J, T_K].map((t) => word(t));

  const S_A = 'Spr' + 'int-';
  const S_B = 'Pha' + 'se ';
  const S_C = 'Wa' + 've ';
  const stale = [
    new RegExp('\\b' + esc(S_A) + '\\d', 'i'),
    new RegExp('\\b' + esc(S_B) + '\\d', 'i'),
    new RegExp('\\b' + esc(S_C) + '\\d', 'i'),
    /\bTBD\b/,
    /\bXXX\b/,
  ];
  return { markers, env, terms, stale };
}

function main(): number {
  const root = repoRoot();
  const files = listFiles(root);
  const allow = loadAllowlist(root);
  const pats = buildPatterns();
  const hits: string[] = [];

  for (const abs of files) {
    const rel = path.relative(root, abs);
    let text = '';
    try {
      text = fs.readFileSync(abs, 'utf8');
    } catch {
      continue;
    }
    const lines = text.split('\n');
    const prose = stripCode(text).split('\n');

    for (let i = 0; i < lines.length; i++) {
      const raw = lines[i];
      const num = i + 1;
      const checkAllowed = (s: string) => isAllowed(rel, s, allow);

      for (const re of pats.markers) {
        if (re.test(raw) && !checkAllowed(raw)) {
          hits.push(`${rel}:${num}: unfinished marker (${re.source.slice(0, 40)}) :: ${raw.trim().slice(0, 160)}`);
        }
      }
      for (const re of pats.env) {
        if (re.test(raw) && !checkAllowed(raw)) {
          hits.push(`${rel}:${num}: legacy env reference :: ${raw.trim().slice(0, 160)}`);
        }
      }
      for (const re of pats.stale) {
        if (re.test(raw) && !checkAllowed(raw)) {
          hits.push(`${rel}:${num}: stale label (${re.source.slice(0, 40)}) :: ${raw.trim().slice(0, 160)}`);
        }
      }

      const proseLine = prose[i] ?? '';
      for (const re of pats.terms) {
        if (re.test(proseLine) && !checkAllowed(raw) && !checkAllowed(proseLine)) {
          hits.push(`${rel}:${num}: internal term (${re.source}) :: ${raw.trim().slice(0, 160)}`);
        }
      }
    }
  }

  if (hits.length > 0) {
    for (const h of hits) console.error(`[lint-docs] ${h}`);
    console.error(`[lint-docs] ${hits.length} hit(s) in ${files.length} file(s). See allowlist scripts/lint-docs-allowlist.json for grandfathered entries.`);
    return 1;
  }
  console.log(`[lint-docs] OK: ${files.length} file(s) clean.`);
  return 0;
}

const code = main();
process.exitCode = code;

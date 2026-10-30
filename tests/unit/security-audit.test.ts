import { describe, it, expect } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import RawAdmZip from 'adm-zip';

// Hardened canonical entrypoint under test.
import AdmZipEngine from '../../packages/engine/src/adm-zip-safe';

// Production importers go through the wrapper — importing them here proves
// they resolve and keep working with hardened instances.
import '../../packages/engine/src/parser';

const DISK_METHODS = ['extractAllTo', 'extractAllToAsync', 'extractEntryTo'] as const;

describe('CVE-2026-76845 adm-zip disk-extraction hardening (REAL instances)', () => {
  it('documents why the old prototype neuter was ineffective (factory owns methods)', () => {
    const raw = new RawAdmZip();
    expect(
      Object.prototype.hasOwnProperty.call(raw, 'extractAllTo'),
      'adm-zip@0.6.1 instances must carry OWN extractAllTo (prototype patch bypass)'
    ).toBe(true);
    expect(
      Object.prototype.hasOwnProperty.call(raw, 'extractEntryTo'),
      'adm-zip@0.6.1 instances must carry OWN extractEntryTo (prototype patch bypass)'
    ).toBe(true);
  });

  it('engine wrapper: new AdmZip().method throws [E_SECURITY]', () => {
    for (const method of DISK_METHODS) {
      const zip = new (AdmZipEngine as unknown as new () => Record<string, unknown>)();
      expect(typeof zip[method]).toBe('function');
      expect(() => (zip[method] as (...a: unknown[]) => unknown)('/tmp/astra-cve-2026-76845')).toThrow(
        /\[E_SECURITY\].*CVE-2026-76845/
      );
    }
  });

  it('disk methods are own, frozen stubs (non-writable, non-configurable)', () => {
    const zip = new (AdmZipEngine as unknown as new () => Record<string, unknown>)();
    for (const method of DISK_METHODS) {
      expect(Object.prototype.hasOwnProperty.call(zip, method)).toBe(true);
      const desc = Object.getOwnPropertyDescriptor(zip, method);
      expect(desc).toBeDefined();
      expect(desc!.writable).toBe(false);
      expect(desc!.configurable).toBe(false);
    }
  });

  it('live temp-dir disk probe: blocked extractAllTo writes zero files', () => {
    const zip = new (AdmZipEngine as unknown as new () => Record<string, unknown>)() as unknown as {
      addFile(n: string, b: Buffer): void;
      toBuffer(): Buffer;
      extractAllTo(p: string): void;
    };
    zip.addFile('word/document.xml', Buffer.from('hello', 'utf8'));
    const buf = zip.toBuffer();
    const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-cve-2026-76845-'));
    try {
      const victim = new (AdmZipEngine as unknown as new (b: Buffer) => {
        extractAllTo(p: string): void;
      })(buf);
      expect(() => victim.extractAllTo(tmp)).toThrow(/\[E_SECURITY\].*CVE-2026-76845/);
      expect(fs.readdirSync(tmp)).toEqual([]);
    } finally {
      fs.rmSync(tmp, { recursive: true, force: true });
    }
  });

  it('live temp-dir disk probe: blocked extractEntryTo writes zero files', () => {
    const zip = new (AdmZipEngine as unknown as new () => Record<string, unknown>)() as unknown as {
      addFile(n: string, b: Buffer): void;
      toBuffer(): Buffer;
    };
    zip.addFile('word/document.xml', Buffer.from('hello', 'utf8'));
    const buf = zip.toBuffer();
    const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-cve-2026-76845-'));
    try {
      const victim = new (AdmZipEngine as unknown as new (b: Buffer) => {
        extractEntryTo(e: string, p: string): void;
      })(buf);
      expect(() => victim.extractEntryTo('word/document.xml', tmp)).toThrow(
        /\[E_SECURITY\].*CVE-2026-76845/
      );
      expect(fs.readdirSync(tmp)).toEqual([]);
    } finally {
      fs.rmSync(tmp, { recursive: true, force: true });
    }
  });

  it('in-memory paths are unaffected (addFile/getEntry/getData/toBuffer/readAsText)', () => {
    const Ctor = AdmZipEngine as unknown as new (b?: Buffer) => {
      addFile(n: string, b: Buffer): void;
      getEntry(n: string): { getData(): Buffer } | null;
      readAsText(n: string): string;
      toBuffer(): Buffer;
    };
    const zip = new Ctor();
    zip.addFile('word/document.xml', Buffer.from('hello', 'utf8'));
    const entry = zip.getEntry('word/document.xml');
    expect(entry).not.toBeNull();
    expect(entry!.getData().toString('utf8')).toBe('hello');
    expect(zip.readAsText('word/document.xml')).toBe('hello');
    const buf = zip.toBuffer();
    expect(Buffer.isBuffer(buf)).toBe(true);
    const roundTrip = new Ctor(buf);
    expect(roundTrip.getEntry('word/document.xml')!.getData().toString('utf8')).toBe('hello');
  });

  it('backend/adm-zip-safe.ts is cleanly deleted in favor of canonical engine package', () => {
    const repoRoot = process.cwd();
    expect(fs.existsSync(path.resolve(repoRoot, 'backend/adm-zip-safe.ts'))).toBe(false);
    expect(fs.existsSync(path.resolve(repoRoot, 'packages/engine/src/adm-zip-safe.ts'))).toBe(true);
  });

  it('zero raw adm-zip imports in production outside the wrapper (static grep-in-test)', () => {
    const repoRoot = process.cwd();
    const prodDirs = [
      'backend',
      'packages/engine/src',
      'packages/cli/src',
      'packages/agent/src', // P0 extract: agency moved here from packages/cli/src
      'packages/mcp-server/src',
      'electron',
      'renderer',
    ];
    const violations: string[] = [];

    const walk = (dir: string): string[] => {
      const out: string[] = [];
      let entries: string[];
      try {
        entries = fs.readdirSync(dir);
      } catch {
        return out;
      }
      for (const name of entries) {
        // Skip build output, deps, and VCS — they bundle adm-zip itself.
        if (name === 'node_modules' || name === 'dist' || name === '.git') continue;
        const full = path.join(dir, name);
        let stat: fs.Stats;
        try {
          stat = fs.statSync(full);
        } catch {
          continue;
        }
        if (stat.isDirectory()) {
          out.push(...walk(full));
        } else if (/\.(ts|tsx|js|jsx|mjs|cjs)$/.test(name) && !/\.d\.ts$/.test(name)) {
          out.push(full);
        }
      }
      return out;
    };

    // Static from/require, subpaths (adm-zip/*), side-effect imports, and
    // dynamic import() of adm-zip all bypass the wrapper and must fail —
    // including backtick-interpolated specifiers (`adm-zip${...}`), whose
    // specifier START is still statically adm-zip. Same shape as
    // scripts/check-adm-zip-imports.mjs: quote/backtick opens the specifier,
    // then a close quote, `/subpath`, or `${` must follow.
    const RAW_IMPORT_RE =
      /from\s+['"`]adm-zip(?=[/'"`$])|require\(\s*['"`]adm-zip(?=[/'"`$])|import\(\s*['"`]adm-zip(?=[/'"`$])|import\s+['"`]adm-zip(?=[/'"`$])/;

    for (const rel of prodDirs) {
      const abs = path.resolve(repoRoot, rel);
      if (!fs.existsSync(abs)) continue;
      for (const file of walk(abs)) {
        // Allowlist: the hardened wrapper itself is the ONLY file allowed to
        // import raw adm-zip in production.
        if (path.basename(file) === 'adm-zip-safe.ts') continue;
        let content: string;
        try {
          content = fs.readFileSync(file, 'utf8');
        } catch {
          continue;
        }
        const lines = content
          .replace(/\/\*[\s\S]*?\*\//g, (m) => m.replace(/[^\n]/g, ''))
          .split('\n');
        lines.forEach((line, idx) => {
          const code = line.replace(/(^|[^:"'`])\/\/.*$/g, '$1');
          if (!RAW_IMPORT_RE.test(code)) return;
          violations.push(`${path.relative(repoRoot, file)}:${idx + 1}: ${line.trim().slice(0, 160)}`);
        });
      }
    }

    expect(violations, `raw adm-zip imports outside adm-zip-safe must be zero; found:\n${violations.join('\n')}`).toEqual([]);
  });
});

/**
 * H2 jail hardening: env-drop + shared-jail parity.
 *
 * - `ASTRA_ALLOWED_DIR` is ignored (D3 drop): pointing it at an outside dir
 *   neither advertises the root nor admits files under it.
 * - `ASTRA_EXTRA_DOC_ROOTS` entries must exist as directories (phantom
 *   entries never widen either jail).
 * - MCP (`session-store.ts`) and engine (`document-manager.ts`) jails give
 *   identical allow/refuse verdicts on an adversarial corpus when rooted at
 *   the same explicit extra root (wide/narrow parity — both consume the Q0a
 *   shared primitives, so semantics cannot disagree).
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
  getMcpAllowedRoots,
  validateDocPath,
} from '../../packages/mcp-server/src/lib/session-store';
import {
  getSandboxAllowedRoots,
  validateSandboxDocumentPath,
  clearTrustedDocPaths,
} from '../../packages/engine/src/document-manager';

const EXTRA_KEY = 'ASTRA_EXTRA_DOC_ROOTS';
const ALLOWED_KEY = 'ASTRA_ALLOWED_DIR';

let savedExtra: string | undefined;
let savedAllowed: string | undefined;

beforeEach(() => {
  savedExtra = process.env[EXTRA_KEY];
  savedAllowed = process.env[ALLOWED_KEY];
  delete process.env[EXTRA_KEY];
  delete process.env[ALLOWED_KEY];
  clearTrustedDocPaths();
});

afterEach(() => {
  if (savedExtra === undefined) delete process.env[EXTRA_KEY];
  else process.env[EXTRA_KEY] = savedExtra;
  if (savedAllowed === undefined) delete process.env[ALLOWED_KEY];
  else process.env[ALLOWED_KEY] = savedAllowed;
  clearTrustedDocPaths();
});

function codeOf(fn: () => unknown): string {
  try {
    fn();
  } catch (err: unknown) {
    return (err as { code?: string }).code ?? `THROW:${err instanceof Error ? err.message : String(err)}`;
  }
  return 'ALLOW';
}

describe('ASTRA_ALLOWED_DIR is ignored (D3 drop)', () => {
  it('env widening admits nothing: root unadvertised, file refused', () => {
    const outsideBase = fs.mkdtempSync(path.join(os.homedir(), 'astra-env-probe-'));
    const outside = path.join(outsideBase, 'vault');
    try {
      fs.mkdirSync(outside, { recursive: true });
      const evil = path.join(outside, 'evil.docx');
      fs.writeFileSync(evil, 'x');
      process.env[ALLOWED_KEY] = outside;
      const roots = getMcpAllowedRoots();
      expect(roots).not.toContain(path.resolve(outside));
      expect(roots).not.toContain(outside);
      // Refusal here is pure containment (no dots, .docx ext, absolute):
      // the ONLY way this file passes is the env root.
      expect(codeOf(() => validateDocPath(evil))).toBe('E_SANDBOX');
      const engineRoots = getSandboxAllowedRoots();
      expect(engineRoots).not.toContain(path.resolve(outside));
      expect(() => validateSandboxDocumentPath(evil)).toThrow();
    } finally {
      fs.rmSync(outsideBase, { recursive: true, force: true });
    }
  });

  it('pointing the env var at a tmpdir subdir still admits nothing extra', () => {
    const sub = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-allowed-probe-'));
    try {
      process.env[ALLOWED_KEY] = sub;
      // Subdir of the tmpdir base root passes only via the BASE root, never
      // via the env var: roots list is unchanged with and without the var.
      const withEnv = getMcpAllowedRoots();
      delete process.env[ALLOWED_KEY];
      const withoutEnv = getMcpAllowedRoots();
      expect(withEnv).toEqual(withoutEnv);
    } finally {
      fs.rmSync(sub, { recursive: true, force: true });
    }
  });
});

describe('extra roots are existence-checked', () => {
  it('phantom entries never join either jail', () => {
    // Outside every base root (homedir is never a root) so refusal can only
    // come from the phantom extra entry being correctly excluded.
    const phantom = path.join(os.homedir(), `astra-phantom-${process.pid}-nope`);
    expect(fs.existsSync(phantom)).toBe(false);
    process.env[EXTRA_KEY] = phantom;
    expect(getMcpAllowedRoots()).not.toContain(path.resolve(phantom));
    expect(getSandboxAllowedRoots()).not.toContain(path.resolve(phantom));
    expect(codeOf(() => validateDocPath(path.join(phantom, 'x.docx')))).toBe('E_SANDBOX');
    expect(() => validateSandboxDocumentPath(path.join(phantom, 'x.docx'))).toThrow();
  });

  it('a file (not a directory) never joins either jail', () => {
    const file = path.join(os.tmpdir(), `astra-rootfile-${process.pid}.txt`);
    fs.writeFileSync(file, 'x');
    try {
      process.env[EXTRA_KEY] = file;
      expect(getMcpAllowedRoots()).not.toContain(path.resolve(file));
      expect(getSandboxAllowedRoots()).not.toContain(path.resolve(file));
    } finally {
      fs.rmSync(file, { force: true });
    }
  });

  it('an existing directory joins both jails and admits files', () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-extra-'));
    try {
      process.env[EXTRA_KEY] = dir;
      expect(getMcpAllowedRoots()).toContain(path.resolve(dir));
      expect(getSandboxAllowedRoots()).toContain(path.resolve(dir));
      const inside = path.join(dir, 'ok.docx');
      fs.writeFileSync(inside, 'x');
      expect(validateDocPath(inside)).toBe(path.resolve(inside));
      expect(validateSandboxDocumentPath(inside)).toBe(path.resolve(inside));
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });
});

describe('MCP/engine wide-narrow parity on one explicit root', () => {
  it('identical allow/refuse verdicts across an adversarial corpus', () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-parity-'));
    try {
      process.env[EXTRA_KEY] = dir;
      fs.writeFileSync(path.join(dir, 'ok.docx'), 'x');
      fs.mkdirSync(path.join(dir, 'sub'));
      const corpus = [
        path.join(dir, 'ok.docx'),
        path.join(dir, 'sub', 'new.docx'),
        path.join(dir, '..', 'evil.docx'),
        '/etc/passwd.docx',
        path.join(dir, '.ssh', 'x.docx'),
        path.join(dir, '.evil.docx'),
        'https://evil/x.docx',
        'report.pdf',
      ];
      for (const candidate of corpus) {
        const mcp = codeOf(() => validateDocPath(candidate));
        let engine: string;
        try {
          validateSandboxDocumentPath(candidate);
          engine = 'ALLOW';
        } catch {
          engine = 'REFUSE';
        }
        const mcpAllows = mcp === 'ALLOW';
        // .pdf diverges by design (MCP: E_VALIDATION on extension; engine
        // rejects non-docx too) — both refuse, codes may differ. What must
        // never happen: one side allows what the other refuses.
        expect(mcpAllows, `parity allow for ${candidate}`).toBe(engine === 'ALLOW');
      }
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });

  it('symlink escapes refuse on both sides', () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-parity-link-'));
    try {
      process.env[EXTRA_KEY] = dir;
      const outsideTarget = fs.existsSync('/etc/hosts') ? '/etc/hosts' : '/etc/passwd';
      const link = path.join(dir, 'evil.docx');
      try {
        fs.symlinkSync(outsideTarget, link);
      } catch {
        return; // no symlink privilege — lexical parity above suffices
      }
      expect(codeOf(() => validateDocPath(link))).toBe('E_SANDBOX');
      expect(() => validateSandboxDocumentPath(link)).toThrow();
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });
});

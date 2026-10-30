/**
 * Q0a jail-equivalence: central engine jail gives identical verdicts to the
 * removed local copies on an adversarial corpus.
 *
 * Reference implementations below are byte-identical copies of the deleted
 * locals (normalize + win32-aware prefix; realpath + walk-up). The test
 * asserts old-vs-new agreement on: absolute, `..`, backslash, NUL, scheme,
 * hidden, and symlinked-dir cases — plus assertInside E_SANDBOX+remedy shape.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
  isPathInsideDir,
  canonicalizeForAsset,
  assertInside,
} from '@attentionform/astra-engine';

// ── Reference copies of the removed locals ──────────────────────────────
function oldIsPathInsideDir(
  candidate: string,
  dir: string,
  platform: string = process.platform
): boolean {
  const normC = path.normalize(candidate);
  const normD = path.normalize(dir);
  if (platform === 'win32') {
    const lc = normC.toLowerCase();
    const ld = normD.toLowerCase();
    return lc === ld || lc.startsWith(ld + path.sep.toLowerCase());
  }
  return normC === normD || normC.startsWith(normD + path.sep);
}

function oldCanonicalize(resolved: string): string {
  try {
    return fs.realpathSync(resolved);
  } catch {
    /* not existing — walk up */
  }
  let cur = resolved;
  const stack: string[] = [];
  for (;;) {
    const parent = path.dirname(cur);
    if (parent === cur) return resolved;
    stack.unshift(path.basename(cur));
    cur = parent;
    try {
      return path.join(fs.realpathSync(cur), ...stack);
    } catch {
      /* keep walking up */
    }
  }
}

function oldResolveCanonicalPath(p: string): string {
  try {
    if (fs.existsSync(p)) {
      return fs.realpathSync(p);
    }
  } catch {
    return p;
  }
  const parent = path.dirname(p);
  if (parent === p) {
    return p;
  }
  const realParent = oldResolveCanonicalPath(parent);
  return path.join(realParent, path.basename(p));
}

function oldIsContainedWithin(parent: string, child: string): boolean {
  const normParent = process.platform === 'win32' ? path.normalize(parent).toLowerCase() : path.normalize(parent);
  const normChild = process.platform === 'win32' ? path.normalize(child).toLowerCase() : path.normalize(child);
  const rel = path.relative(normParent, normChild);
  return rel === '' || (!rel.startsWith('..' + path.sep) && rel !== '..' && !path.isAbsolute(rel));
}

describe('Q0a jail equivalence (old-vs-new identical verdicts)', () => {
  const dir = path.resolve(os.tmpdir(), 'q0a-jail-corpus-root');

  const lexicalCorpus: Array<[string, string]> = [
    // absolute
    [path.join(dir, 'a.docx'), dir],
    [path.join(dir, 'sub', 'b.docx'), dir],
    [path.join(os.tmpdir(), 'outside.docx'), dir],
    ['/etc/passwd', dir],
    [dir, dir],
    // .. traversal
    [path.join(dir, '..', 'evil.docx'), dir],
    [dir + '/../evil.docx', dir],
    [path.join(dir, 'sub', '..', 'ok.docx'), dir],
    // backslash (POSIX treats as literal chars; win32 branch lowercases)
    [dir + '\\evil.docx', dir],
    [dir + '\\..\\evil', dir],
    ['C:\\Windows\\evil.docx', dir],
    // NUL
    [dir + '/a\0b.docx', dir],
    [dir + '/%00evil.docx', dir],
    // scheme
    ['https://evil/x.docx', dir],
    ['file:///etc/passwd', dir],
    ['ftp:evil.docx', dir],
    // hidden
    [path.join(dir, '.ssh', 'x.docx'), dir],
    [path.join(dir, '.evil.docx'), dir],
    [path.join(dir, 'sub', '.git', 'x'), dir],
    // prefix siblings (must stay outside)
    [dir + '2/x.docx', dir],
    [dir + '-suffix', dir],
    [dir + '/sub2/../sub/ok.docx', dir],
    // trailing-slash dir form
    [path.join(dir, 'a'), dir + path.sep],
  ];

  for (const platform of [process.platform, 'win32']) {
    it(`lexical verdicts identical (platform=${platform})`, () => {
      for (const [candidate, d] of lexicalCorpus) {
        expect(
          isPathInsideDir(candidate, d, platform),
          `candidate=${JSON.stringify(candidate)} dir=${JSON.stringify(d)}`
        ).toBe(oldIsPathInsideDir(candidate, d, platform));
      }
    });
  }

  it('canonicalize outputs identical (existing + walk-up)', () => {
    const targets = [
      dir, // likely non-existing: walk-up path
      path.join(dir, 'nope', 'nested.docx'),
      __filename, // existing file
      path.dirname(__filename), // existing dir
      '/',
    ];
    for (const t of targets) {
      expect(canonicalizeForAsset(t), `target=${t}`).toBe(oldCanonicalize(t));
    }
  });

  it('existsSync-gated canonicalize matches central (existing + walk-up)', () => {
    const targets = [
      dir,
      path.join(dir, 'nope', 'nested.docx'),
      __filename,
      path.dirname(__filename),
      '/',
    ];
    for (const t of targets) {
      expect(canonicalizeForAsset(t), `target=${t}`).toBe(oldResolveCanonicalPath(t));
    }
  });

  // Callers resolve both sides before checking, so compare resolved forms
  // (raw dir with trailing slash diverges between prefix and relative
  // checks, but resolve strips it before the jail ever sees it).
  it('relative-based containment matches central on lexical corpus', () => {
    for (const [candidate, d] of lexicalCorpus) {
      const rc = path.resolve(candidate);
      const rd = path.resolve(d);
      expect(
        isPathInsideDir(rc, rd),
        `candidate=${JSON.stringify(candidate)} dir=${JSON.stringify(d)}`
      ).toBe(oldIsContainedWithin(rd, rc));
    }
  });

  describe('symlinked dirs (live fs)', () => {
    let root = '';
    let outside = '';
    beforeEach(() => {
      root = fs.mkdtempSync(path.join(os.tmpdir(), 'q0a-jail-'));
      outside = fs.mkdtempSync(path.join(os.tmpdir(), 'q0a-outside-'));
      fs.writeFileSync(path.join(outside, 'secret.docx'), 'x');
      fs.mkdirSync(path.join(root, 'real'));
      fs.writeFileSync(path.join(root, 'real', 'ok.docx'), 'x');
      try {
        fs.symlinkSync(outside, path.join(root, 'link-out'), 'dir');
      } catch {
        /* symlink privilege missing — tests below skip */
      }
      try {
        fs.symlinkSync(
          path.join(outside, 'secret.docx'),
          path.join(root, 'evil.docx')
        );
      } catch {
        /* ignore */
      }
    });
    afterEach(() => {
      try {
        fs.rmSync(root, { recursive: true, force: true });
      } catch {
        /* ignore */
      }
      try {
        fs.rmSync(outside, { recursive: true, force: true });
      } catch {
        /* ignore */
      }
    });

    it('canonicalize + verdicts identical on symlink escapes', () => {
      const linkOut = path.join(root, 'link-out');
      expect(fs.existsSync(linkOut), 'symlink creation unavailable — cannot prove jail containment').toBe(true);
      const cases = [
        path.join(root, 'real', 'ok.docx'), // inside
        path.join(root, 'real', 'new.docx'), // inside, not-yet-created
        path.join(linkOut, 'secret.docx'), // symlink-parent escape
        path.join(linkOut, 'new.docx'), // symlink-parent escape, not-yet-created
        path.join(root, 'evil.docx'), // symlinked file escape
      ];
      for (const c of cases) {
        const nCanon = canonicalizeForAsset(c);
        const oCanon = oldCanonicalize(c);
        const gCanon = oldResolveCanonicalPath(c);
        expect(nCanon, `canon ${c}`).toBe(oCanon);
        expect(nCanon, `gated canon ${c}`).toBe(gCanon);
        expect(isPathInsideDir(nCanon, root), `verdict ${c}`).toBe(
          oldIsPathInsideDir(oCanon, root)
        );
        expect(isPathInsideDir(nCanon, root), `gated verdict ${c}`).toBe(
          oldIsContainedWithin(root, gCanon)
        );
      }
      // escape verdicts are refusals on both sides
      expect(isPathInsideDir(canonicalizeForAsset(path.join(linkOut, 'secret.docx')), root)).toBe(false);
      expect(isPathInsideDir(canonicalizeForAsset(path.join(root, 'real', 'ok.docx')), root)).toBe(true);
    });

    it('assertInside throws E_SANDBOX+remedy on escape, returns canonical inside', () => {
      const linkOut = path.join(root, 'link-out');
      expect(fs.existsSync(linkOut), 'symlink creation unavailable — cannot prove jail containment').toBe(true);
      const inside = path.join(root, 'real', 'ok.docx');
      expect(assertInside(inside, root)).toBe(canonicalizeForAsset(path.normalize(inside)));
      for (const esc of [path.join(linkOut, 'secret.docx'), '/etc/passwd']) {
        let err: unknown;
        try {
          assertInside(esc, root);
        } catch (e) {
          err = e;
        }
        expect(err).toBeInstanceOf(Error);
        expect(String((err as Error).message)).toContain('E_SANDBOX');
        expect((err as { code?: unknown }).code).toBe('E_SANDBOX');
        expect(typeof (err as { remedy?: unknown }).remedy).toBe('string');
      }
    });
  });
});

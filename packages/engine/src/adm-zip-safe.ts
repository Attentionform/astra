/**
 * CVE-2026-76845 hardened adm-zip entrypoint (REAL fix).
 *
 * Background: adm-zip@0.6.1 (and 0.6.0) exports a factory function whose instances carry
 * OWN `extractAllTo` / `extractAllToAsync` / `extractEntryTo` properties. A
 * prototype patch (commit b2417a6) therefore never touches live instances and
 * is bypassed — proven by a live disk-write probe. Do NOT reintroduce a
 * prototype neuter; it is ineffective for 0.6.x.
 *
 * This module is the ONLY allowed production entrypoint for adm-zip. It wraps
 * the raw factory and returns instances whose three disk-extraction methods
 * are replaced with own, non-writable, non-configurable stubs that throw
 * `[E_SECURITY] ... (CVE-2026-76845)`. All in-memory APIs (`getEntry` /
 * `getEntries`, `getData` / `readAsText` / `readFile`, `toBuffer`, `addFile` /
 * `updateFile` / `deleteFile`, etc.) are untouched and keep working.
 *
 * Fail-open disclaimer (honest): protection applies ONLY to instances created
 * via this wrapper. A direct `import ... from 'adm-zip'` bypasses it entirely
 * (fail-open). `tests/unit/security-audit.test.ts` statically greps production
 * trees for raw `adm-zip` imports outside this wrapper and fails on any, but
 * that is a lint-time guard — not a runtime sandbox. Any future file that
 * imports raw `adm-zip` reintroduces the CVE. With adm-zip bumped to 0.6.1,
 * this wrapper remains as defense-in-depth instance hardening.
 */
import RawAdmZip from 'adm-zip';

const DISK_METHODS = ['extractAllTo', 'extractAllToAsync', 'extractEntryTo'] as const;
type DiskMethod = (typeof DISK_METHODS)[number];

function blockedDiskMethod(method: DiskMethod): (...args: unknown[]) => never {
  const fn = function (..._args: unknown[]): never {
    void _args;
    throw new Error(
      `[E_SECURITY] adm-zip ${method} blocked (CVE-2026-76845): disk extraction is disabled; use in-memory getEntry/getData/toBuffer only.`
    );
  };
  return fn;
}

function hardenInstance<T extends object>(zip: T): T {
  const rec = zip as unknown as Record<string, unknown>;
  for (const method of DISK_METHODS) {
    try {
      Object.defineProperty(rec, method, {
        value: blockedDiskMethod(method),
        writable: false,
        configurable: false,
        enumerable: false,
      });
    } catch {
      try {
        (rec as unknown as Record<string, (...args: unknown[]) => never>)[method] =
          blockedDiskMethod(method);
      } catch {
        /* ignore: best-effort hardening */
      }
    }
  }
  return zip;
}

// Factory wrapper: works with `new AdmZip(...)` and `AdmZip(...)` — either
// way an explicitly-returned hardened object wins over `this`.
function AdmZipImpl(this: unknown, ...args: unknown[]): unknown {
  const Raw = RawAdmZip as unknown as new (...a: unknown[]) => object;
  const inst = new Raw(...args);
  return hardenInstance(inst);
}

// Instance type (so `private zip: AdmZip` keeps working after migration).
// Merged with the const below: default import carries BOTH value (factory)
// and type (instance), exactly like `import AdmZip from 'adm-zip'` did.
type RawAdmZipType = InstanceType<typeof RawAdmZip>;
// eslint-disable-next-line @typescript-eslint/no-empty-object-type -- intentional: declaration merging with the const below (value+type parity with `import AdmZip from 'adm-zip'`)
interface AdmZip extends RawAdmZipType {}

const AdmZip = AdmZipImpl as unknown as typeof RawAdmZip;

export default AdmZip;
export { AdmZip };



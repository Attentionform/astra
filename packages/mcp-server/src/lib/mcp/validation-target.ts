/**
 * The Office schema version this server validates against. The operator sets it once, before the
 * server starts (`--office-version <version>` or `ASTRA_OFFICE_VERSION`; the flag wins). It is read
 * at startup and then fixed: every document session uses it for the check made at open and for every
 * later check, so a save compares like with like. Tools take no such parameter, and an agent cannot
 * change it. Every call site passes `serverOfficeVersion()` to the validator explicitly.
 */
import { resolveOfficeVersion, type OfficeVersion } from '@attentionform/astra-engine';

let configured: OfficeVersion | undefined;

/** Set by the `--office-version` flag at startup. */
export function configureOfficeVersion(version: OfficeVersion | undefined): void {
  configured = version;
}

/** The version every validation call of this server uses: the flag, then `ASTRA_OFFICE_VERSION`, then `Office2019`. */
export function serverOfficeVersion(): OfficeVersion {
  return resolveOfficeVersion(configured);
}

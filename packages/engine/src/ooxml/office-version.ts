/**
 * The Office schema version the validator checks a document against.
 *
 * The version is always resolved to one explicit value before the validator is called, so the value
 * that is reported and the value that was used are the same. The library keeps no mutable setting:
 * a caller passes `officeVersion`, and a server or CLI resolves its own flag once and passes it on.
 */
import type { OfficeVersion } from '@attentionform/ooxml-validator';

export type { OfficeVersion };

/**
 * The Office versions that validation can target, oldest first.
 */
export const OFFICE_VERSIONS: readonly OfficeVersion[] = [
  'Office2007',
  'Office2010',
  'Office2013',
  'Office2016',
  'Office2019',
  'Office2021',
  'Microsoft365',
];

/** The target ASTra validates against when nothing else is chosen. */
export const DEFAULT_OFFICE_VERSION: OfficeVersion = 'Office2019';

/** The environment variable read when no explicit version is passed. */
export const OFFICE_VERSION_ENV = 'ASTRA_OFFICE_VERSION';

/** True when `value` is one of the seven supported Office versions. */
export function isOfficeVersion(value: unknown): value is OfficeVersion {
  return typeof value === 'string' && (OFFICE_VERSIONS as readonly string[]).includes(value);
}

/** Text that names the valid choices, for usage errors. */
export function officeVersionChoices(): string {
  return OFFICE_VERSIONS.join(', ');
}

/**
 * Checks a value that came from outside (a flag, an environment variable, a JavaScript caller).
 * Throws a `TypeError` that names `source` and lists the seven valid versions.
 */
export function parseOfficeVersion(value: unknown, source = 'officeVersion'): OfficeVersion {
  if (isOfficeVersion(value)) return value;
  throw new TypeError(`Invalid ${source} ${JSON.stringify(value)}. Expected one of: ${officeVersionChoices()}`);
}

/**
 * Picks the version for one validation call: the explicit option, then the `ASTRA_OFFICE_VERSION`
 * environment variable (read only, never written), then `Office2019`.
 */
export function resolveOfficeVersion(
  explicit?: OfficeVersion,
  env: NodeJS.ProcessEnv = process.env,
): OfficeVersion {
  if (explicit !== undefined) return parseOfficeVersion(explicit);
  const fromEnv = env[OFFICE_VERSION_ENV];
  if (fromEnv !== undefined && fromEnv.trim() !== '') return parseOfficeVersion(fromEnv.trim(), OFFICE_VERSION_ENV);
  return DEFAULT_OFFICE_VERSION;
}

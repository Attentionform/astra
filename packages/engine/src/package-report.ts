/**
 * What the engine tells a caller about the package it opened, beyond the document itself (DN3, 9.2).
 * Present only when there is something to say: a canonical, unsigned package has an empty report.
 */
export interface PackageReport {
  /** Parts whose namespace prefixes were rewritten to the usual ones at open. `from[i]` was rewritten to `to[i]` (prefixes; '' is the default namespace). */
  canonicalised: Array<{ part: string; from: string[]; to: string[] }>;
  /** True when the package carries a digital signature (entries under `_xmlsignatures/`). */
  signed: boolean;
}

export const SIGNED_WARNING = 'This package is digitally signed. Saving it changes signed parts and invalidates the signature.';

const label = (prefix: string): string => (prefix === '' ? 'the default namespace' : prefix);

/** Fixed warning texts for a report, in a stable order: signature first, then the rewrite. Empty when there is nothing to say. */
export function packageWarnings(report: PackageReport): string[] {
  const warnings: string[] = [];
  if (report.signed) warnings.push(SIGNED_WARNING);
  if (report.canonicalised.length > 0) {
    const first = report.canonicalised[0];
    const example = `${first.part}: ${first.from.map((f, i) => `${label(f)} to ${first.to[i]}`).join(', ')}`;
    const n = report.canonicalised.length;
    warnings.push(
      `${n} part${n === 1 ? '' : 's'} used namespace prefixes other than the usual ones and ${n === 1 ? 'was' : 'were'} rewritten in memory (for example ${example}). Saving writes the usual prefixes.`
    );
  }
  return warnings;
}

/** The `packageNotes` field of an open result: absent for a canonical, unsigned package so its output does not change. */
export function packageNotes(report: PackageReport): (PackageReport & { warnings: string[] }) | undefined {
  const warnings = packageWarnings(report);
  return warnings.length === 0 ? undefined : { ...report, warnings };
}

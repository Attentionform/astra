/**
 * What the engine found out about the package that it opened, beyond the document itself. An unsigned package that follows the usual namespace prefixes has an empty report.
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
    const shown = report.canonicalised.filter((c) => c.from.length > 0);
    const first = shown[0] ?? report.canonicalised[0];
    const pairs = first.from.map((f, i) => `${label(f)} to ${first.to[i]}`).join(', ');
    const example = pairs === '' ? first.part : `${first.part}: ${pairs}`;
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

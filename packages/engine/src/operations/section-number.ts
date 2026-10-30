import { toError } from '../errors';

/**
 * Sections are numbered from 1 in every operation (1 is the first section),
 * like rows, columns and chart series. The document model counts from 0, so an
 * operation converts here once and works with the zero-based position after.
 * Returns undefined when the caller named no section.
 */
export function sectionPosition(
  opName: string,
  raw: unknown,
  sectionCount: number,
  whenOmitted = 'omit it'
): number | undefined {
  if (raw === undefined || raw === null || raw === '') return undefined;
  const n = typeof raw === 'number' ? raw : Number(String(raw));
  if (!Number.isInteger(n) || n < 1 || n > sectionCount) {
    throw toError(
      'E_VALIDATION',
      `${opName} sectionIndex ${String(raw)} is out of range (document has ${sectionCount} section${sectionCount === 1 ? '' : 's'})`,
      `pass sectionIndex in [1, ${sectionCount}], counting from 1, or ${whenOmitted}; re-read the structure with inspect_document`
    );
  }
  return n - 1;
}

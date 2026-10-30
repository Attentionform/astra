import { toError, getEngineErrorCode } from '../errors';

/**
 * Fuzz-corpus coded-error contract: raw archive/XML toolkit failures on the
 * parse/save ingress surface E_VALIDATION naming the part (never raw
 * ADM-ZIP/xmldom/zlib text). Already-coded errors (E_LIMIT ceilings,
 * E_SECURITY/XXE gates, macro refusals) pass through untouched.
 */
export function rethrowCoded(err: unknown): void {
  if (err instanceof Error && getEngineErrorCode(err)) throw err;
}

/** Wrap a raw archive open failure (ADM-ZIP format/LOC rejects) as E_VALIDATION. */
export function invalidPackageError(err: unknown): Error {
  rethrowCoded(err);
  const msg = err instanceof Error ? err.message : String(err);
  return toError(
    'E_VALIDATION',
    `Invalid DOCX package: ${msg}`,
    're-export the file from Word (or pick a valid .docx) and retry',
  );
}

/** Wrap a raw part-data read failure (CRC/inflate/zlib short-read) as E_VALIDATION naming the part. */
export function unreadablePartError(partName: string, err: unknown): Error {
  rethrowCoded(err);
  const msg = err instanceof Error ? err.message : String(err);
  return toError(
    'E_VALIDATION',
    `Invalid DOCX part '${partName}': ${msg}`,
    're-export the file from Word (or pick a valid .docx) and retry',
  );
}

/** Coded part read: entry.getData() CRC/inflate failures surface E_VALIDATION naming the part. */
export function readPartData(entry: { getData(): Buffer }, partName: string): Buffer {
  try {
    return entry.getData();
  } catch (err) {
    throw unreadablePartError(partName, err);
  }
}

/** Wrap a raw XML toolkit failure ([xmldom error/warning]) as E_VALIDATION. */
export function invalidXmlError(err: unknown): Error {
  rethrowCoded(err);
  const msg = err instanceof Error ? err.message : String(err);
  return toError(
    'E_VALIDATION',
    `Invalid DOCX XML: ${msg}`,
    'repair the part XML or re-save the document in Word before retrying',
  );
}

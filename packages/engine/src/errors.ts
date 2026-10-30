/**
 * Q0b central engine error table (single source of truth).
 *
 * Full EngineErrorCode set + code→remedy map + toError helper. Downstream
 * CLI / agent / MCP error tables import this module instead of maintaining
 * local copies. Previously-passing codes keep identical behavior; previously-
 * dropped codes (E_SESSION/E_CONFLICT/E_NOT_FOUND/E_SECURITY/E_STATE and, per
 * surface, E_LIMIT/E_IO/E_DEPENDENCY/E_AUTH) now pass through verbatim with
 * code intact.
 */

export const ENGINE_ERROR_CODES = [
  'E_VALIDATION',
  'E_SESSION',
  'E_QUOTA',
  'E_IO',
  'E_CONFLICT',
  'E_NOT_FOUND',
  'E_SANDBOX',
  'E_LIMIT',
  'E_SECURITY',
  'E_STATE',
  'E_DEPENDENCY',
  'E_AUTH',
] as const;

export type EngineErrorCode = (typeof ENGINE_ERROR_CODES)[number];

const ENGINE_ERROR_CODES_SET: ReadonlySet<string> = new Set<string>([
  ...ENGINE_ERROR_CODES,
]);

/** Default human remedy per code (engine errors may carry their own). */
export const ENGINE_ERROR_REMEDIES: Record<EngineErrorCode, string> = {
  // E_VALIDATION default stays node-agnostic on purpose: it fires only for
  // remedy-less foreign errors (every engine throw carries its own), and
  // "re-inspect ids" would lie for container-level failures (corrupt bytes,
  // unreadable parts) that have no node to re-inspect.
  E_VALIDATION: 're-inspect the document via `astra inspect` and retry, or re-export the file from Word when the package itself is unreadable',
  E_SESSION: 'call open_document first to establish an active session, then retry',
  E_QUOTA: 'split into batches of <= 500 operations (sequential:true) and retry',
  E_IO: 'check the path exists and is readable, then retry',
  E_CONFLICT: 're-inspect live state, then retry with a fresh id or force:true',
  E_NOT_FOUND: 'verify the id or path exists, then retry',
  E_SANDBOX: 'keep the path inside the allowed directory and retry',
  E_LIMIT:
    'raise ASTRA_MAX_RATIO / --max-ratio or ASTRA_MAX_UNCOMPRESSED_BYTES / --max-uncompressed-bytes above the observed value, or use a smaller document',
  E_SECURITY: 'remove the blocked content or path and retry',
  E_STATE: 're-inspect live vs saved state, then retry the operation',
  E_DEPENDENCY: 'install the missing backend (LibreOffice for PDF/visual) and retry',
  E_AUTH: 'provide a valid key via --key and retry',
};

export type EngineError = Error & { code: EngineErrorCode; remedy: string; retryable: false };

/** Build a coded engine error (message carries `[CODE]`, plus `.code`/`.remedy`). */
export function toError(
  code: EngineErrorCode,
  message: string,
  remedy?: string
): EngineError {
  // Remedies are values, not labelled sentences: callers occasionally pass a
  // "Remedy: ..." string; strip the label so the message never doubles it.
  const finalRemedy = String(remedy ?? ENGINE_ERROR_REMEDIES[code]).replace(/^\s*Remedy:\s*/i, '');
  const body = String(message ?? '').trim() || code;
  const withCode = body.includes(code) ? body : `[${code}] ${body}`;
  const text = withCode.includes('Remedy:')
    ? withCode
    : `${withCode} Remedy: ${finalRemedy}`;
  return Object.assign(new Error(text), { code, remedy: finalRemedy, retryable: false as const });
}

/** True when `value` is a known engine error code string. */
export function isEngineErrorCode(value: unknown): value is EngineErrorCode {
  return typeof value === 'string' && ENGINE_ERROR_CODES_SET.has(value);
}

/**
 * Extract the engine code from an error: `.code` prop first, then `[E_*]` /
 * `E_*` substring in the message. Returns undefined when no known code rides
 * along (callers apply their own fallback).
 */
export function getEngineErrorCode(err: unknown): EngineErrorCode | undefined {
  const code = (err as { code?: unknown } | null | undefined)?.code;
  if (isEngineErrorCode(code)) return code;
  const msg =
    err instanceof Error
      ? err.message
      : typeof err === 'string'
        ? err
        : String(err ?? '');
  for (const candidate of ENGINE_ERROR_CODES) {
    if (msg.includes(candidate)) return candidate;
  }
  return undefined;
}

/** Extract with fallback (strict: callers always name the expected code). */
export function resolveEngineErrorCode(
  err: unknown,
  fallback: EngineErrorCode
): EngineErrorCode {
  return getEngineErrorCode(err) ?? fallback;
}

/** Default remedy for a code (used when the error carries none of its own). */
export function getEngineRemedy(code: EngineErrorCode): string {
  return ENGINE_ERROR_REMEDIES[code];
}

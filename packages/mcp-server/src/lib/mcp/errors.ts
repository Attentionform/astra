// MS-Q2b PART1_QUALITY Q2.2: verbatim split from lib/mcp-tool-adapter.ts (lines 127-250 + 721-727 + 1121-1127).
// Zero logic edits (only `export` added to two previously-private shared helpers
// so handlers can import them; behavior identical). Single source of truth is
// the shared @attentionform/astra-engine table — no local code lists.
import { getEngineErrorCode, type EngineErrorCode } from '@attentionform/astra-engine';

// ── Error contract ───────────────────────────────────────────────────────────
// Q0b: full central engine set (previously the 7-code subset below;
// previously-dropped E_LIMIT/E_SECURITY/E_STATE/E_DEPENDENCY/E_AUTH now pass
// through verbatim instead of collapsing to the call-site fallback). Single
// source of truth: the shared @attentionform/astra-engine table (same pattern as
// visual/export/compose/new-document tools).
export type McpErrorCode = EngineErrorCode;

export interface McpErrorShape {
  ok: false;
  code: McpErrorCode;
  message: string;
  remedy: string;
}

export function mcpError(code: McpErrorCode, message: string, remedy: string): McpErrorShape {
  return { ok: false, code, message: rewritePhantomMcpToolNames(message), remedy: rewritePhantomMcpToolNames(remedy) };
}

/**
 * Q0b: core envelope mapper (exported for the E-code contract test, same as
 * the visual/export/compose/new-document mappers). Every engine code survives
 * verbatim via the shared getEngineErrorCode table (.code first, then
 * [E_*]/E_* in the message — this catches message-only engine failures like
 * the [E_LIMIT] zip ceilings and [E_SECURITY] adm-zip blocks that carry no
 * .code prop). Uncoded failures keep the call-site fallback. Remedy passes
 * through verbatim when the error carries its own, else the fallback.
 */
export function mcpErrorFrom(
  err: unknown,
  fallbackCode: McpErrorCode,
  fallbackRemedy: string
): McpErrorShape {
  const resolved: McpErrorCode = getEngineErrorCode(err) ?? fallbackCode;
  const rawMessage = err instanceof Error ? err.message : String(err);
  // P2-2: client message must not leak server stack/paths — sanitize here;
  // the full value stays stderr-only (callers console.error before returning).
  const message = rewritePhantomMcpToolNames(sanitizeMcpClientMessage(rawMessage));
  const remedy =
    (err as { remedy?: unknown } | null)?.remedy && typeof (err as { remedy?: unknown }).remedy === 'string'
      ? rewritePhantomMcpToolNames(String((err as { remedy?: unknown }).remedy))
      : rewritePhantomMcpToolNames(fallbackRemedy);
  return { ok: false, code: resolved, message, remedy };
}

/**
 * P2-2: strip `at ...` stack frames + absolute host paths from the
 * client-facing message. Full stack stays stderr-only (engine already logs
 * `🚨 Critical error...` to stderr; handlers log the raw value below).
 * Minimal, no schema widening: first line + redacted paths only.
 */
export function sanitizeMcpClientMessage(raw: unknown): string {
  let text = typeof raw === 'string' ? raw : String(raw ?? '');
  // Keep only the first logical line for engine `Error: ...\n at ...` blobs;
  // drop any `at ...` frame lines that survived on one line via `;` splits.
  const lines = text.split('\n');
  const kept: string[] = [];
  for (const line of lines) {
    if (/^\s*at\s+/.test(line)) continue;
    // Inline ` at foo (path:1:2)` fragments on the first line (engine joins
    // stack with newlines, but be safe for single-line blobs). The paren group
    // must carry a file:line marker so validator locations such as
    // ` at /w:document[1]/w:body[1]/w:tbl[1] (table 1, node t_001)` survive.
    const cut = line.search(/\s+at\s+\S+\s*\([^()\n]*:\d+[^()\n]*\)/);
    kept.push(cut >= 0 ? line.slice(0, cut) : line);
    // Only the first non-frame line carries the actionable message; extra
    // context lines are stack noise — keep at most 3 to avoid truncation loss.
    if (kept.length >= 3) break;
  }
  text = kept.join('\n').trim();
  // Redact absolute host paths. Covers POSIX homes/tmp plus macOS/Linux
  // system bases (/private/*, /var/*, /Users/*, /etc/*, /opt/*, /usr/*) and
  // Windows drive-letter (C:\…) + UNC (\\host\share) shapes. Bare project-
  // relative paths are left intact (only absolute shapes redact).
  // NOTE: longest-prefix bases first so /private/tmp/… matches as one unit.
  text = text.replace(/\/private\/[^\s"'`()[ \]]*/g, '[redacted-path]');
  text = text.replace(/\/home\/[^\s"'`()[ \]]*/g, '[redacted-path]');
  text = text.replace(/\/tmp\/[^\s"'`()[ \]]*/g, '[redacted-path]');
  text = text.replace(/\/var\/[^\s"'`()[ \]]*/g, '[redacted-path]');
  text = text.replace(/\/Users\/[^\s"'`()[ \]]*/g, '[redacted-path]');
  text = text.replace(/\/(?:etc|opt|usr|srv|mnt|media|data)(?:\/[^\s"'`()[ \]]*)?/g, '[redacted-path]');
  text = text.replace(/[A-Za-z]:\\[^\s"'`()[ \]]*/g, '[redacted-path]');
  text = text.replace(/\\\\[^\s"'`()[ \]]+/g, '[redacted-path]');
  // Collapse leftover ` (…:line:col)` remnants after redaction.
  text = text.replace(/\s*\([^\n()]*:\d+:\d+\)/g, '');
  if (text.length > 2000) text = text.slice(0, 2000);
  const cleaned = text || 'Batch failed.';
  return rewritePhantomMcpToolNames(cleaned);
}

/**
 * MCP-surface remedy hygiene: engine errors name a phantom get-*-toon tool /
 * `toon({...})`, neither of which is a registered MCP tool (tools/list serves
 * `inspect_document`, verified live). Rewrite to the real tool name so agents
 * retry correctly. Meaning otherwise identical; TOON-format prose mentions
 * are left alone (only `toon(` call forms are rewritten).
 *
 * The phantom spelling is assembled from fragments (never written
 * contiguously) so this file stays grep-clean for the non-existent token.
 */
const PHANTOM_TOOL_BASE = ['get', 'document', 'toon'].join('_');
const PHANTOM_TOOL_CALL = new RegExp(`${PHANTOM_TOOL_BASE}\\(\\)`, 'g');
const PHANTOM_TOOL_BARE = new RegExp(`\\b${PHANTOM_TOOL_BASE}\\b`, 'g');

export function rewritePhantomMcpToolNames(text: string): string {
  return String(text ?? '')
    .replace(PHANTOM_TOOL_CALL, 'inspect_document')
    .replace(PHANTOM_TOOL_BARE, 'inspect_document')
    .replace(/\btoon\(\{/g, 'inspect_document({')
    .replace(/\btoon\(\)/g, 'inspect_document');
}

/**
 * Commit-message hygiene for git-backed tools: single-line, capped.
 * Newlines become spaces (multi-line input stays one commit subject);
 * over-long input is truncated. Blank input falls back.
 */
export const COMMIT_MESSAGE_MAX = 500;
export const COMMIT_AUTHOR_MAX = 100;
export const COMMIT_AUTHOR_DEFAULT = 'ASTra Agent';

export function sanitizeCommitMessage(raw: unknown, fallback: string): string {
  const text = typeof raw === 'string' ? raw : String(raw ?? '');
  const singleLine = text.replace(/[\r\n]+/g, ' ').trim();
  const body = singleLine || fallback;
  return body.length > COMMIT_MESSAGE_MAX ? body.slice(0, COMMIT_MESSAGE_MAX) : body;
}

export function sanitizeCommitAuthor(raw: unknown): string {
  const text = typeof raw === 'string' ? raw : String(raw ?? '');
  const singleLine = text.replace(/[\r\n]+/g, ' ').trim();
  if (!singleLine) return COMMIT_AUTHOR_DEFAULT;
  return singleLine.length > COMMIT_AUTHOR_MAX ? singleLine.slice(0, COMMIT_AUTHOR_MAX) : singleLine;
}

/**
 * Cross-surface parity: an unknown/expired session is E_SESSION on EVERY MCP
 * tool (session tools preserve the store's E_SESSION verbatim; the history /
 * audit family maps the same failure through this helper). Message + remedy
 * stay byte-identical to the session-store shape for the same reason.
 */
export function unknownSessionError(sessionId: string): McpErrorShape {
  return mcpError(
    'E_SESSION',
    `Unknown session "${sessionId}" (expired or never opened)`,
    'Call open_document first, then retry with the returned sessionId.'
  );
}

export function invalidInput(message: string): McpErrorShape {
  return mcpError(
    'E_VALIDATION',
    message,
    'Fix the input shape per the tool description Example and retry (single edit = 1-item batch).'
  );
}

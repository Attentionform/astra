// MS-Q2b PART1_QUALITY Q2.2: verbatim split from lib/mcp-tool-adapter.ts (lines 3031-3156).
// Zero logic edits.
import {
  checkoutCommit,
  getHistory,
} from '@attentionform/astra-engine';
import type { McpSessionStore } from '../../session-store';
import {
  mcpErrorFrom,
  invalidInput,
  unknownSessionError,
} from '../errors';
import type { McpErrorShape } from '../errors';
import {
  AUDIT_RETENTION_NOTE,
  hashDocumentXml,
  sidecarRepoExists,
} from '../diff';
import { ExportAuditTrailSchema } from '../schemas';

export interface AuditTrailRow {
  oid: string;
  message: string;
  author: string;
  /** Ms-precision epoch millis (audit-bridge meta when tracked, else git seconds). */
  timestampMs: number;
  opCount?: number;
  affectedNodeIds?: string[];
  resultingHeadOid: string;
  /** sha256 hex over the committed document.xml at this revision (reproducible). */
  packageHash: string;
}

/**
 * export_audit_trail handler (handler-only, NOT a registered MCP tool — the
 * 12-tool pin is unaffected). Returns the full saved history plus a sha256
 * over each revision's committed document.xml, recomputed via the read-only
 * checkoutCommit path so verifiers can reproduce every hash independently.
 * Read-only: never mutates revision state, disk bytes, or the redo stack.
 * The HOLD retention note rides along because .astra/ is deletable local
 * state and full server-side retention is OUT of scope.
 */
export async function handleExportAuditTrail(
  store: McpSessionStore,
  input: unknown
): Promise<
  | {
      ok: true;
      sessionId: string;
      documentPath: string;
      totalCommits: number;
      currentCommit: string | null;
      returnedCount: number;
      truncated: boolean;
      trail: AuditTrailRow[];
      retentionNote: string;
    }
  | McpErrorShape
> {
  const parsed = ExportAuditTrailSchema.safeParse(input);
  if (!parsed.success) {
    const detail = parsed.error.issues.map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`).join('; ');
    return invalidInput(`export_audit_trail: invalid input. ${detail}`);
  }
  const limit = parsed.data.limit ?? 1000;
  const offset = parsed.data.offset ?? 0;
  try {
    return await store.withSession(parsed.data.sessionId, async (handle) => {
      const filePath = handle.filePath;
      const sessionId = handle.sessionId;
      if (!sidecarRepoExists(filePath)) {
        throw Object.assign(new Error('export_audit_trail: no saved history for this session yet'), {
          code: 'E_NOT_FOUND',
          remedy: 'Call save_document first to record a history baseline, then retry export_audit_trail.',
        });
      }
      let full: Awaited<ReturnType<typeof getHistory>>;
      try {
        full = await getHistory(filePath);
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        throw Object.assign(new Error(`export_audit_trail: cannot read history: ${msg}`), {
          code: 'E_IO',
          remedy: 'Retry export_audit_trail; re-open the document if the session expired.',
        });
      }
      const slice = full.slice(offset, offset + limit);
      const trail: AuditTrailRow[] = [];
      for (const c of slice) {
        // Read-only per-revision read + hash (reproduces independently).
        let target: Awaited<ReturnType<typeof checkoutCommit>>;
        try {
          target = await checkoutCommit(filePath, c.oid);
        } catch (err: unknown) {
          const msg = err instanceof Error ? err.message : String(err);
          throw Object.assign(new Error(`export_audit_trail: cannot read commit "${c.oid}": ${msg}`), {
            code: 'E_IO',
            remedy: 'Retry export_audit_trail; re-open the document if the session expired.',
          });
        }
        const meta = store.getCommitMeta(c.oid);
        const row: AuditTrailRow = {
          oid: c.oid,
          message: typeof c.message === 'string' ? c.message.replace(/\s+$/, '') : String(c.message ?? ''),
          author: typeof c.author === 'string' ? c.author : String(c.author ?? ''),
          timestampMs: meta ? meta.timestampMs : c.timestamp.getTime(),
          // Always present: 0/[] covers pre-tracking commits and heals-only
          // saves (opCount folds operations, not normalizer rewrites).
          opCount: meta ? meta.opCount : 0,
          affectedNodeIds: meta ? [...meta.affectedNodeIds] : [],
          resultingHeadOid: c.oid,
          packageHash: hashDocumentXml(target.xml),
        };
        trail.push(row);
      }
      return {
        ok: true as const,
        sessionId,
        documentPath: filePath,
        totalCommits: full.length,
        currentCommit: full.length > 0 ? full[0].oid : null,
        returnedCount: trail.length,
        truncated: offset + trail.length < full.length,
        trail,
        retentionNote: AUDIT_RETENTION_NOTE,
      };
    });
  } catch (err: unknown) {
    const code = (err as { code?: unknown } | null)?.code;
    if (code === 'E_CONFLICT') {
      return mcpErrorFrom(err, 'E_CONFLICT', 'The session changed while the trail was reading. Re-open if needed, then retry export_audit_trail.');
    }
    if (code === 'E_SESSION' || /Unknown session|was closed concurrently|expired/.test(err instanceof Error ? err.message : String(err))) {
      return unknownSessionError(parsed.data.sessionId);
    }
    return mcpErrorFrom(err, 'E_IO', 'Retry export_audit_trail; re-open the document if the session expired.');
  }
}

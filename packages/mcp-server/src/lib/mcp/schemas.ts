// MS-Q2b PART1_QUALITY Q2.2: verbatim split from lib/mcp-tool-adapter.ts (lines 923-1076).
// Zero logic edits.
import { z } from 'zod';
import { OperationSchema, SanitizeOptionsSchema } from '@attentionform/astra-engine';

// ── Strict input schemas (no loose records; batch reuses the strict union) ───
export const OpenDocumentSchema = z
  .object({
    path: z.string().min(1).max(1024),
    sessionId: z
      .string()
      .min(1)
      .max(128)
      .optional()
      .describe(
        'DEPRECATED: session ids are server-minted (sess_<uuid>); omit sessionId and use the returned id. ' +
          'Client-supplied values are rejected with E_VALIDATION (deterministic ids are test-only via --allow-client-session-ids).'
      ),
    // Additive flag: true returns the palette-only summary (no stylesXml
    // dump). NO .default() here (same wire-breaker reason as inspect
    // windowing below): defaults are applied manually in the handler so the
    // wire shape stays {path,sessionId?} for existing clients. Omitted/false
    // keeps the full structural summary including the stylesXml dump.
    summaryOnly: z.boolean().optional(),
    // Macro/active-content disposition (refuse-default): a document carrying
    // macro parts opens fail-closed with E_VALIDATION naming the parts.
    // allowMacros:true preserves them byte-identical; stripMacros:true
    // removes them at open. Both together fail with E_VALIDATION.
    allowMacros: z.boolean().optional(),
    stripMacros: z.boolean().optional(),
    // Default author for tracked changes and comments that do not name one (falls back to
    // ASTRA_AUTHOR on the server, then "ASTra"). An operation's own author always wins.
    author: z.string().trim().min(1).max(128).optional(),
  })
  .strict();

export const InspectDocumentSchema = z
  .object({
    sessionId: z.string().min(1).max(128),
    fidelity: z.enum(['compact', 'standard', 'full']).optional(),
    sectionId: z.string().min(1).max(256).optional(),
    maxTextLength: z.number().int().min(1).max(10000).optional(),
    // Phase C-2 smart windowing: offset/limit page the node rows; aroundNode +
    // window zooms to [node-w, node+w]; maxTokens cuts on a clean row boundary.
    // Wire-breaker loop-back: NO .default() here. McpServer.registerTool
    // zod-parses BEFORE the handler, so .default(0/50/5) injected
    // offset/limit/window into rawObj over real wire — every aroundNode-only
    // zoom was rejected as "mutually exclusive" and the small-doc envelope
    // diverged (wire limit:50 vs unit min(total,200)). Bounds stay enforced;
    // defaults (offset 0, limit 50, window 5) are applied manually in
    // handleInspectDocument AFTER the explicit-presence check below.
    offset: z.number().int().min(0).optional(),
    limit: z.number().int().min(1).max(200).optional(),
    aroundNode: z.string().min(1).max(256).optional(),
    window: z.number().int().min(1).max(20).optional(),
    maxTokens: z.number().int().min(200).max(8000).optional(),
  })
  .strict();

export const SearchContentSchema = z
  .object({
    sessionId: z.string().min(1).max(128),
    query: z.string().min(1).max(500),
    // Phase C-3: flag-gated regex. Default false = current literal substring
    // behavior (no ReDoS surface by default). True = RegExp semantics with
    // bounded pattern length + an evaluation time budget.
    regex: z.boolean().optional().default(false),
  })
  .strict();

export const ExecuteBatchSchema = z
  .object({
    sessionId: z.string().min(1).max(128),
    operations: z.array(OperationSchema).min(1).max(500),
    sequential: z.boolean().optional(),
    rollbackOnError: z.boolean().optional(),
    maxSteps: z.number().int().min(1).optional(),
    // Jail for image `filePath` reads. Explicit value wins; otherwise the
    // session anchor (dirname of the open document) applies. Consulted only
    // when an operation carries `filePath` — other batches read nothing.
    allowedDir: z.string().min(1).max(1024).optional(),
    // Per-batch style guide override (bundled name or guide file path).
    // Applies to this batch only; omitted inherits the session guide.
    // Empty (including whitespace-only) is rejected in the handler with a
    // detach hint — no min(1) here on purpose so the hint (not a bare
    // schema error) reaches the caller.
    styleGuide: z.string().max(1024).optional(),
  })
  .strict();

export const SaveDocumentSchema = z
  .object({
    sessionId: z.string().min(1).max(128),
    outPath: z.string().min(1).max(1024).optional(),
    // Caps (500/100) are enforced in executeSaveDocument with a handler-level
    // E_VALIDATION + shorten-and-retry remedy. NO .max() here on purpose:
    // a zod .max() rejects over-long input in McpServer.registerTool BEFORE
    // the handler runs, surfacing as a bare MCP -32602 with no remedy field.
    // Unbounded wire strings reach the handler, which fails closed honestly.
    message: z.string().optional(),
    author: z.string().optional(),
    // Macro/active-content disposition for the save (same refuse-default as
    // open): allowMacros:true preserves macro parts byte-identical;
    // stripMacros:true removes them before the validation gate. Both
    // together fail with E_VALIDATION. Omitted keeps the session disposition.
    allowMacros: z.boolean().optional(),
    stripMacros: z.boolean().optional(),
  })
  .strict();

export const CloseDocumentSchema = z
  .object({
    sessionId: z.string().min(1).max(128),
  })
  .strict();

export const GetDocumentHistorySchema = z
  .object({
    sessionId: z.string().min(1).max(128),
    limit: z.number().int().min(1).max(100).optional(),
    offset: z.number().int().min(0).optional(),
  })
  .strict();

export const RevertDocumentSchema = z
  .object({
    sessionId: z.string().min(1).max(128),
    steps: z.number().int().min(1).optional(),
    commitOid: z.string().min(1).max(128).optional(),
    discardUnsaved: z.boolean().optional(),
    message: z.string().max(500).optional(),
    author: z.string().max(100).optional(),
    reason: z.string().max(500).optional(),
  })
  .strict();

export const RedoDocumentSchema = z
  .object({
    sessionId: z.string().min(1).max(128),
    message: z.string().max(500).optional(),
    author: z.string().max(100).optional(),
    reason: z.string().max(500).optional(),
  })
  .strict();

export const PreviewRevertSchema = z
  .object({
    sessionId: z.string().min(1).max(128),
    steps: z.number().int().min(1).optional(),
    commitOid: z.string().min(1).max(128).optional(),
  })
  .strict();

export const ExportAuditTrailSchema = z
  .object({
    sessionId: z.string().min(1).max(128),
    limit: z.number().int().min(1).max(1000).optional(),
    offset: z.number().int().min(0).optional(),
  })
  .strict();

// ── Document-audit tools (strict schemas, no flat-key drift: nested option
// keys resolve exactly as the engine spells them) ───
export const ValidateDocumentSchema = z
  .object({
    sessionId: z.string().min(1).max(128),
  })
  .strict();

export const GetStatsSchema = z
  .object({
    sessionId: z.string().min(1).max(128),
  })
  .strict();

// Scrub options mirror the engine sanitizer keys 1:1 (spread from the engine
// schema, so the wire can never drift from what the scrub path accepts).
export const SanitizeDocumentSchema = z
  .object({
    sessionId: z.string().min(1).max(128),
    ...SanitizeOptionsSchema.shape,
  })
  .strict();

export const NormalizeDocumentSchema = z
  .object({
    sessionId: z.string().min(1).max(128),
  })
  .strict();

export const ImportDocumentSchema = z
  .object({
    // Open session only: never auto-creates (use new_document first for a
    // fresh templated document, open_document for an existing file).
    sessionId: z.string().min(1).max(128),
    path: z.string().min(1).max(1024),
    // Jail for image refs inside the Markdown (same rule as compose_document).
    allowedDir: z.string().min(1).max(1024).optional(),
    template: z.enum(['report', 'memo', 'blank']).optional(),
    // Style guide attach (bundled name or guide file path, same jail as
    // compose_document). Empty is rejected in the handler with a detach hint.
    styleGuide: z.string().max(1024).optional(),
    // Read $...$ as inline math (off by default: prices and shell text can look like math).
    dollarMath: z.boolean().optional(),
  })
  .strict();

export type OpenDocumentInput = z.infer<typeof OpenDocumentSchema>;
export type InspectDocumentInput = z.infer<typeof InspectDocumentSchema>;
export type SearchContentInput = z.infer<typeof SearchContentSchema>;
export type ExecuteBatchInput = z.infer<typeof ExecuteBatchSchema>;
export type SaveDocumentInput = z.infer<typeof SaveDocumentSchema>;
export type CloseDocumentInput = z.infer<typeof CloseDocumentSchema>;
export type GetDocumentHistoryInput = z.infer<typeof GetDocumentHistorySchema>;
export type RevertDocumentInput = z.infer<typeof RevertDocumentSchema>;
export type RedoDocumentInput = z.infer<typeof RedoDocumentSchema>;
export type PreviewRevertInput = z.infer<typeof PreviewRevertSchema>;
export type ExportAuditTrailInput = z.infer<typeof ExportAuditTrailSchema>;
export type ValidateDocumentInput = z.infer<typeof ValidateDocumentSchema>;
export type GetStatsInput = z.infer<typeof GetStatsSchema>;
export type SanitizeDocumentInput = z.infer<typeof SanitizeDocumentSchema>;
export type NormalizeDocumentInput = z.infer<typeof NormalizeDocumentSchema>;
export type ImportDocumentInput = z.infer<typeof ImportDocumentSchema>;

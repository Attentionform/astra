// MS-Q2b PART1_QUALITY Q2.2: verbatim split from lib/mcp-tool-adapter.ts (lines 76-125).
// Zero logic edits. Lean-6 + git-history tool names + research flag.
// ── MCP Lean-6 core server (Sprint-2 Muse) ────────────────────────────────────
// Lean 6 only: open_document / inspect_document / search_content /
// execute_batch / save_document / close_document. No single-op tool:
// a single edit is a 1-item batch. Strict OperationSchema /
// BatchOperationsSchema gate every mutation (unknown actions, missing fields
// and unknown flat keys are rejected with remedy, never applied).
//
// Research tools are not registered by this core server. No --with-research
// launch flag exists (unknown flags exit 2); this constant is a legacy
// helper and enables nothing at startup.
export const MCP_RESEARCH_FLAG = '--with-research';

/** Legacy helper: reports argv intent only; server startup accepts no such flag. */
export function isResearchEnabled(argv: readonly string[] = process.argv): boolean {
  try {
    return argv.includes(MCP_RESEARCH_FLAG);
  } catch {
    return false;
  }
}

/** Lean-6 tool names (locked, no single_op). */
export const LEAN6_TOOL_NAMES = [
  'open_document',
  'inspect_document',
  'search_content',
  'execute_batch',
  'save_document',
  'close_document',
] as const;

export type Lean6ToolName = (typeof LEAN6_TOOL_NAMES)[number];

export function getLean6ToolNames(): Lean6ToolName[] {
  return [...LEAN6_TOOL_NAMES];
}

/** Git-history tool names (history browsing + preview + revert + redo; registered alongside lean-6). */
export const GIT_HISTORY_TOOL_NAMES = [
  'get_document_history',
  'preview_revert',
  'revert_document',
  'redo_document',
] as const;

export type GitHistoryToolName = (typeof GIT_HISTORY_TOOL_NAMES)[number];

export function getGitHistoryToolNames(): GitHistoryToolName[] {
  return [...GIT_HISTORY_TOOL_NAMES];
}

/** Document-audit tool names (read-only preflight + census + scrub + canonical repair + file ingest). */
export const PARITY_TOOL_NAMES = [
  'validate_document',
  'get_stats',
  'sanitize_document',
  'normalize_document',
  'import_document',
] as const;

export type ParityToolName = (typeof PARITY_TOOL_NAMES)[number];

export function getParityToolNames(): ParityToolName[] {
  return [...PARITY_TOOL_NAMES];
}

// Verbatim split from lib/mcp-tool-adapter.ts.
// Zero logic edits. Lean-6 dispatch table only (no AI SDK dependency:
// handlers are plain functions served over MCP stdio).
import { handleOpenDocument, handleInspectDocument } from './handlers/open-inspect';
import { handleSearchContent, handleExecuteBatch } from './handlers/search-batch';
import { handleSaveDocument, handleCloseDocument } from './handlers/save-close';

/** Lean-6 dispatch table (handlers above; research intentionally absent). */
export const LEAN6_HANDLERS = {
  open_document: handleOpenDocument,
  inspect_document: handleInspectDocument,
  search_content: handleSearchContent,
  execute_batch: handleExecuteBatch,
  save_document: handleSaveDocument,
  close_document: handleCloseDocument,
} as const;

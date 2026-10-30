// lib/mcp-tool-adapter.ts is now a thin barrel.
// All logic moved verbatim to ./mcp/* modules (normalize-read/tool-names/errors/
// diff/descriptions/schemas/handlers/* + registry). This file re-exports the old
// paths so `import { ... } from '../lib/mcp-tool-adapter'` keeps working with
// zero test changes. Hand-rolled MCP handlers preserved (split != swap).
export * from './mcp/normalize-read';
export * from './mcp/tool-names';
export * from './mcp/errors';
export * from './mcp/diff';
export * from './mcp/descriptions';
export * from './mcp/schemas';
export * from './mcp/handlers/open-inspect';
export * from './mcp/handlers/search-batch';
export * from './mcp/handlers/save-close';
export * from './mcp/handlers/history';
export * from './mcp/handlers/export-audit';
export * from './mcp/registry';

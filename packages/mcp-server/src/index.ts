#!/usr/bin/env node
/**
 * @attentionform/astra-mcp-server — Track-B lean-6 MCP server over stdio.
 *
 * Exposes 19 tools (lean-6 + new_document + compose_document + export + visual + git history + document audit):
 *   open_document / inspect_document / search_content /
 *   execute_batch / save_document / close_document / new_document /
 *   compose_document /
 *   export_document / visual_inspect /
 *   get_document_history / preview_revert / revert_document / redo_document /
 *   validate_document / get_stats / sanitize_document / normalize_document / import_document
 *
 * Engine wiring:
 *   the session store + lean-6 handlers live in the local lib
 *   (`./lib/session-store.ts`, `./lib/mcp-tool-adapter.ts`).
 *   Tool names, input shapes and the {ok:false,code,message,remedy} error contract
 *   stay identical across refactors (clients pin them).
 *
 * Transport notes:
 *   - stdio only (stdout is the JSON-RPC wire — NEVER log to stdout).
 *   - every handler result is serialized as a single text content frame;
 *     {ok:false,...} results are additionally flagged isError:true
 *     (fail-closed: the model must fix-and-retry per `remedy`, never assume
 *     the mutation landed).
 */

import * as path from 'path';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';

// ── Engine wiring: local lib imports ─────────
import {
  createMcpSessionStore,
  hasClientSessionIdHatch,
  type McpSessionStore,
} from './lib/session-store';
import {
  CloseDocumentSchema,
  ExecuteBatchSchema,
  GetDocumentHistorySchema,
  GetStatsSchema,
  GIT_HISTORY_DESCRIPTIONS,
  GIT_HISTORY_HANDLERS,
  GIT_HISTORY_TOOL_NAMES,
  ImportDocumentSchema,
  InspectDocumentSchema,
  LEAN6_DESCRIPTIONS,
  LEAN6_HANDLERS,
  LEAN6_TOOL_NAMES,
  NormalizeDocumentSchema,
  OpenDocumentSchema,
  PARITY_DESCRIPTIONS,
  PARITY_TOOL_NAMES,
  PreviewRevertSchema,
  RedoDocumentSchema,
  RevertDocumentSchema,
  SanitizeDocumentSchema,
  SaveDocumentSchema,
  SearchContentSchema,
  ValidateDocumentSchema,
  type GitHistoryToolName,
  type Lean6ToolName,
  type ParityToolName,
} from './lib/mcp-tool-adapter';
import { ExportDocumentSchema, EXPORT_DOCUMENT_DESCRIPTION, handleExportDocument } from './tools/export';
import { VisualInspectSchema, VISUAL_INSPECT_DESCRIPTION, handleVisualInspect } from './tools/visual';
import { NewDocumentSchema, NEW_DOCUMENT_DESCRIPTION, handleNewDocument } from './tools/new-document';
import { ComposeDocumentSchema, COMPOSE_DOCUMENT_DESCRIPTION, handleComposeDocument } from './tools/compose-document';
import {
  handleGetStats,
  handleImportDocument,
  handleNormalizeDocument,
  handleSanitizeDocument,
  handleValidateDocument,
} from './lib/mcp/handlers/parity';
import {
  checkSchemasInDir,
  formatMismatchWarning,
  generateSchemasToDir,
} from './lib/mcp/schema-dump';
import packageJson from '../package.json';

export const SERVER_NAME = 'astra-mcp-server';
export const SERVER_VERSION: string = (packageJson as { version?: string }).version ?? '0.1.0';

/** Re-exported so smoke tests can pin the locked tool set without transport. */
export { LEAN6_TOOL_NAMES, GIT_HISTORY_TOOL_NAMES, PARITY_TOOL_NAMES };
export type { Lean6ToolName, GitHistoryToolName, ParityToolName };

type TextResult = {
  content: [{ type: 'text'; text: string }];
  isError?: boolean;
};

function toTextResult(result: unknown): TextResult {
  const failed =
    typeof result === 'object' &&
    result !== null &&
    (result as { ok?: unknown }).ok === false;
  return {
    content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
    ...(failed ? { isError: true as const } : {}),
  };
}

/**
 * Ordered tool names served on stdio (lean-6 + new_document + compose_document +
 * export + visual + git history + document audit). Single list so --help and the serving line agree.
 */
export function getServedToolNames(): string[] {
  return [
    ...LEAN6_TOOL_NAMES,
    'new_document',
    'compose_document',
    'export_document',
    'visual_inspect',
    ...GIT_HISTORY_TOOL_NAMES,
    ...PARITY_TOOL_NAMES,
  ];
}

/** Usage text for --help and unknown-flag errors. Built from live constants. */
export function buildUsageText(): string {
  const tools = getServedToolNames();
  return [
    `${SERVER_NAME} v${SERVER_VERSION} — document IDE server over MCP stdio`,
    '',
    `Usage: ${SERVER_NAME} [options]`,
    '',
    'Options:',
    '  -h, --help              Print this usage and exit (server never starts)',
    '  -V, --version           Print the server version and exit (server never starts)',
    '  --allowed-dir <dir>     Permit document access under <dir> (repeatable;',
    '                          dash-leading values rejected); also via',
    '                          ASTRA_EXTRA_DOC_ROOTS',
    '  --allow-client-session-ids',
    '                          Permit client-supplied open_document/new_document',
    '                          sessionId (deterministic ids for tests; default:',
    '                          rejected with E_VALIDATION)',
    '  --dump-schemas <dir>    Write one JSON schema per served tool into <dir>',
    '                          (each stamped with the server version); refreshes',
    '                          file-based client copies; server never starts',
    '  --check-schemas <dir>   Warn when stamped copies in <dir> differ from the',
    '                          installed server; exit 1 on mismatch, 0 when fresh',
    '  --                      End of options (no positionals: any use exits 2)',
    '',
    `Tools served — ${tools.length} tools: ${tools.join(', ')}`,
    '',
    'Transport: stdio only. Stdout is the JSON-RPC wire — logs go to stderr.',
    '',
    'Examples:',
    `  ${SERVER_NAME} --allowed-dir /home/user/docs`,
    `  ${SERVER_NAME} --help`,
    `  ${SERVER_NAME} --version`,
  ].join('\n');
}

/** Exit code for unknown flags / missing option values (usage goes to stderr). */
export const CLI_USAGE_ERROR = 2;

/**
 * Parse CLI arguments and environment variables for allowed directories (--allowed-dir / ASTRA_EXTRA_DOC_ROOTS).
 * Appends valid directories to ASTRA_EXTRA_DOC_ROOTS so the sandbox session store and export tools permit access.
 * Dash-leading --allowed-dir values are rejected (never added); parsing stops
 * at the first `--` end-of-options marker so flags after it stay positional.
 */
export function initAllowedDirs(argv: string[] = process.argv): void {
  const currentRoots = new Set(
    (process.env.ASTRA_EXTRA_DOC_ROOTS || '')
      .split(path.delimiter)
      .map((p) => p.trim())
      .filter(Boolean)
  );

  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === '--') break;
    if (argv[i] === '--allowed-dir' && argv[i + 1] && !argv[i + 1].startsWith('-')) {
      currentRoots.add(path.resolve(argv[i + 1]));
      i++;
    } else if (argv[i].startsWith('--allowed-dir=')) {
      const value = argv[i].slice('--allowed-dir='.length).trim();
      if (value && !value.startsWith('-')) currentRoots.add(path.resolve(value));
    }
  }
  if (currentRoots.size > 0) {
    process.env.ASTRA_EXTRA_DOC_ROOTS = Array.from(currentRoots).join(path.delimiter);
  }
}

/**
 * Read a `--flag <dir>` / `--flag=<dir>` value from argv (first win).
 * Returns undefined when the flag is absent. Validation (missing or
 * dash-leading value) happens in main() before this runs.
 */
export function takeDirFlag(argv: readonly string[], flag: string): string | undefined {
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === flag && typeof argv[i + 1] === 'string') return argv[i + 1];
    if (argv[i].startsWith(`${flag}=`)) return argv[i].slice(flag.length + 1);
  }
  return undefined;
}

/**
 * Builds the MCP server bound to `store` (19 tools: lean-6 + new_document + compose_document + export + visual + git history + document audit).
 * A fresh store is created per call when omitted (no module-level singleton —
 * matches the backend session-store ownership rule).
 */
export function createAstraMcpServer(store: McpSessionStore = createMcpSessionStore()): McpServer {
  initAllowedDirs();
  const server = new McpServer({ name: SERVER_NAME, version: SERVER_VERSION });

  // McpServer.registerTool validates inputSchema (zod) at the SDK layer
  // BEFORE handlers run, so schema violations surface as MCP -32602 text
  // (isError:true, no {ok,remedy}) rather than the handler {ok:false,remedy}
  // envelope. Strict gates stay: no schema widening.
  server.registerTool(
    'open_document',
    { description: LEAN6_DESCRIPTIONS.open_document, inputSchema: OpenDocumentSchema },
    async (args: unknown) => toTextResult(await LEAN6_HANDLERS.open_document(store, args))
  );
  server.registerTool(
    'inspect_document',
    { description: LEAN6_DESCRIPTIONS.inspect_document, inputSchema: InspectDocumentSchema },
    async (args: unknown) => toTextResult(await LEAN6_HANDLERS.inspect_document(store, args))
  );
  server.registerTool(
    'search_content',
    { description: LEAN6_DESCRIPTIONS.search_content, inputSchema: SearchContentSchema },
    async (args: unknown) => toTextResult(await LEAN6_HANDLERS.search_content(store, args))
  );
  server.registerTool(
    'execute_batch',
    { description: LEAN6_DESCRIPTIONS.execute_batch, inputSchema: ExecuteBatchSchema },
    async (args: unknown) => toTextResult(await LEAN6_HANDLERS.execute_batch(store, args))
  );
  server.registerTool(
    'save_document',
    { description: LEAN6_DESCRIPTIONS.save_document, inputSchema: SaveDocumentSchema },
    async (args: unknown) => toTextResult(await LEAN6_HANDLERS.save_document(store, args))
  );
  server.registerTool(
    'close_document',
    { description: LEAN6_DESCRIPTIONS.close_document, inputSchema: CloseDocumentSchema },
    async (args: unknown) => toTextResult(await LEAN6_HANDLERS.close_document(store, args))
  );
  server.registerTool(
    'new_document',
    { description: NEW_DOCUMENT_DESCRIPTION, inputSchema: NewDocumentSchema },
    async (args: unknown) => toTextResult(await handleNewDocument(store, args))
  );
  server.registerTool(
    'compose_document',
    { description: COMPOSE_DOCUMENT_DESCRIPTION, inputSchema: ComposeDocumentSchema },
    async (args: unknown) => toTextResult(await handleComposeDocument(store, args))
  );
  server.registerTool(
    'export_document',
    { description: EXPORT_DOCUMENT_DESCRIPTION, inputSchema: ExportDocumentSchema },
    async (args: unknown) => toTextResult(await handleExportDocument(store, args))
  );
  server.registerTool(
    'visual_inspect',
    {
      description: VISUAL_INSPECT_DESCRIPTION,
      inputSchema: VisualInspectSchema,
    },
    async (args: unknown) => toTextResult(await handleVisualInspect(store, args))
  );
  server.registerTool(
    'get_document_history',
    { description: GIT_HISTORY_DESCRIPTIONS.get_document_history, inputSchema: GetDocumentHistorySchema },
    async (args: unknown) => toTextResult(await GIT_HISTORY_HANDLERS.get_document_history(store, args))
  );
  server.registerTool(
    'preview_revert',
    { description: GIT_HISTORY_DESCRIPTIONS.preview_revert, inputSchema: PreviewRevertSchema },
    async (args: unknown) => toTextResult(await GIT_HISTORY_HANDLERS.preview_revert(store, args))
  );
  server.registerTool(
    'revert_document',
    { description: GIT_HISTORY_DESCRIPTIONS.revert_document, inputSchema: RevertDocumentSchema },
    async (args: unknown) => toTextResult(await GIT_HISTORY_HANDLERS.revert_document(store, args))
  );
  server.registerTool(
    'redo_document',
    { description: GIT_HISTORY_DESCRIPTIONS.redo_document, inputSchema: RedoDocumentSchema },
    async (args: unknown) => toTextResult(await GIT_HISTORY_HANDLERS.redo_document(store, args))
  );
  server.registerTool(
    'validate_document',
    { description: PARITY_DESCRIPTIONS.validate_document, inputSchema: ValidateDocumentSchema },
    async (args: unknown) => toTextResult(await handleValidateDocument(store, args))
  );
  server.registerTool(
    'get_stats',
    { description: PARITY_DESCRIPTIONS.get_stats, inputSchema: GetStatsSchema },
    async (args: unknown) => toTextResult(await handleGetStats(store, args))
  );
  server.registerTool(
    'sanitize_document',
    { description: PARITY_DESCRIPTIONS.sanitize_document, inputSchema: SanitizeDocumentSchema },
    async (args: unknown) => toTextResult(await handleSanitizeDocument(store, args))
  );
  server.registerTool(
    'normalize_document',
    { description: PARITY_DESCRIPTIONS.normalize_document, inputSchema: NormalizeDocumentSchema },
    async (args: unknown) => toTextResult(await handleNormalizeDocument(store, args))
  );
  server.registerTool(
    'import_document',
    { description: PARITY_DESCRIPTIONS.import_document, inputSchema: ImportDocumentSchema },
    async (args: unknown) => toTextResult(await handleImportDocument(store, args))
  );

  return server;
}

/** stdio entrypoint: handle --help/--version/unknown flags, else connect and serve until the transport closes. */
export async function main(argv: string[] = process.argv): Promise<number> {
  const args = argv.slice(2);
  if (args.includes('--')) {
    // No positionals defined: `--` alone or with trailing args is a usage
    // error (flags after `--` are positional, never options).
    process.stderr.write(buildUsageText() + '\n');
    return CLI_USAGE_ERROR;
  }
  if (args.includes('--help') || args.includes('-h')) {
    process.stdout.write(buildUsageText() + '\n');
    return 0;
  }
  if (args.includes('--version') || args.includes('-V')) {
    process.stdout.write(`${SERVER_VERSION}\n`);
    return 0;
  }
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (arg === '--allow-client-session-ids') continue;
    if (arg === '--dump-schemas' || arg === '--check-schemas') {
      const value = args[i + 1];
      if (!value || value.startsWith('-')) {
        process.stderr.write(buildUsageText() + '\n');
        return CLI_USAGE_ERROR;
      }
      i++;
      continue;
    }
    if (arg.startsWith('--dump-schemas=') || arg.startsWith('--check-schemas=')) {
      const value = arg.slice(arg.indexOf('=') + 1).trim();
      if (!value || value.startsWith('-')) {
        process.stderr.write(buildUsageText() + '\n');
        return CLI_USAGE_ERROR;
      }
      continue;
    }
    if (arg === '--allowed-dir') {
      const value = args[i + 1];
      if (!value || value.startsWith('-')) {
        process.stderr.write(buildUsageText() + '\n');
        return CLI_USAGE_ERROR;
      }
      i++;
      continue;
    }
    if (arg.startsWith('--allowed-dir=')) {
      const value = arg.slice('--allowed-dir='.length).trim();
      if (!value || value.startsWith('-')) {
        process.stderr.write(buildUsageText() + '\n');
        return CLI_USAGE_ERROR;
      }
      continue;
    }
    if (arg.startsWith('-')) {
      process.stderr.write(buildUsageText() + '\n');
      return CLI_USAGE_ERROR;
    }
  }
  initAllowedDirs(argv);
  const dumpDir = takeDirFlag(argv, '--dump-schemas');
  if (dumpDir !== undefined) {
    const resolved = path.resolve(dumpDir);
    const result = generateSchemasToDir(resolved, SERVER_VERSION);
    process.stdout.write(`[${SERVER_NAME} v${SERVER_VERSION}] wrote ${result.count} schemas to ${resolved}\n`);
    return 0;
  }
  const checkDir = takeDirFlag(argv, '--check-schemas');
  if (checkDir !== undefined) {
    const resolved = path.resolve(checkDir);
    const verdict = checkSchemasInDir(resolved, SERVER_VERSION);
    if (verdict.ok) {
      process.stdout.write(`[${SERVER_NAME} v${SERVER_VERSION}] schemas fresh (${verdict.checked} tools)\n`);
      return 0;
    }
    process.stderr.write(formatMismatchWarning(verdict.mismatches) + '\n');
    return 1;
  }
  const server = createAstraMcpServer(
    createMcpSessionStore({ allowClientSessionIds: hasClientSessionIdHatch(argv) || undefined })
  );
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error(`[${SERVER_NAME} v${SERVER_VERSION}] serving ${getServedToolNames().length} tools (lean-6 + new_document + compose_document + export + visual + git history + document audit) on stdio`);
  return 0;
}

// Auto-run ONLY when executed as the entrypoint (node dist/index.js / bin).
// Under vitest (or any import) process.argv[1] is the test runner, so helpers
// stay import-safe for smoke tests. Stdout is the MCP wire: log via stderr.
function isDirectEntry(): boolean {
  const entry = process.argv?.[1] ?? '';
  return (
    entry.endsWith('/index.js') ||
    entry.endsWith('\\index.js') ||
    entry.endsWith('/index.ts') ||
    entry.endsWith('astra-mcp-server')
  );
}

if (isDirectEntry()) {
  main().then((code: number) => {
    process.exitCode = code;
  }).catch((err: unknown) => {
    const msg = err instanceof Error ? err.message : String(err);
    console.error(`[${SERVER_NAME}] fatal: ${msg}`);
    process.exitCode = 1;
  });
}

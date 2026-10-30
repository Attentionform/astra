import * as fs from 'fs';
import * as path from 'path';
import { z } from 'zod';
import {
  CloseDocumentSchema,
  ExecuteBatchSchema,
  GetDocumentHistorySchema,
  GetStatsSchema,
  GIT_HISTORY_DESCRIPTIONS,
  ImportDocumentSchema,
  InspectDocumentSchema,
  LEAN6_DESCRIPTIONS,
  NormalizeDocumentSchema,
  OpenDocumentSchema,
  PARITY_DESCRIPTIONS,
  PreviewRevertSchema,
  RedoDocumentSchema,
  RevertDocumentSchema,
  SanitizeDocumentSchema,
  SaveDocumentSchema,
  SearchContentSchema,
  ValidateDocumentSchema,
} from '../mcp-tool-adapter';
import { EXPORT_DOCUMENT_DESCRIPTION, ExportDocumentSchema } from '../../tools/export';
import { VISUAL_INSPECT_DESCRIPTION, VisualInspectSchema } from '../../tools/visual';
import { COMPOSE_DOCUMENT_DESCRIPTION, ComposeDocumentSchema } from '../../tools/compose-document';
import { NEW_DOCUMENT_DESCRIPTION, NewDocumentSchema } from '../../tools/new-document';
import packageJson from '../../../package.json';

export const SCHEMA_STAMP_KEY = 'x-generated-from';

export function getServerVersion(): string {
  return (packageJson as { version?: string }).version ?? '0.1.0';
}

export interface ToolDefinition {
  name: string;
  description: string;
  schema: unknown;
}

export function getToolDefinitions(): ToolDefinition[] {
  return [
    { name: 'open_document', description: LEAN6_DESCRIPTIONS.open_document, schema: OpenDocumentSchema },
    { name: 'inspect_document', description: LEAN6_DESCRIPTIONS.inspect_document, schema: InspectDocumentSchema },
    { name: 'search_content', description: LEAN6_DESCRIPTIONS.search_content, schema: SearchContentSchema },
    { name: 'execute_batch', description: LEAN6_DESCRIPTIONS.execute_batch, schema: ExecuteBatchSchema },
    { name: 'save_document', description: LEAN6_DESCRIPTIONS.save_document, schema: SaveDocumentSchema },
    { name: 'close_document', description: LEAN6_DESCRIPTIONS.close_document, schema: CloseDocumentSchema },
    { name: 'new_document', description: NEW_DOCUMENT_DESCRIPTION, schema: NewDocumentSchema },
    { name: 'compose_document', description: COMPOSE_DOCUMENT_DESCRIPTION, schema: ComposeDocumentSchema },
    { name: 'export_document', description: EXPORT_DOCUMENT_DESCRIPTION, schema: ExportDocumentSchema },
    { name: 'visual_inspect', description: VISUAL_INSPECT_DESCRIPTION, schema: VisualInspectSchema },
    { name: 'get_document_history', description: GIT_HISTORY_DESCRIPTIONS.get_document_history, schema: GetDocumentHistorySchema },
    { name: 'preview_revert', description: GIT_HISTORY_DESCRIPTIONS.preview_revert, schema: PreviewRevertSchema },
    { name: 'revert_document', description: GIT_HISTORY_DESCRIPTIONS.revert_document, schema: RevertDocumentSchema },
    { name: 'redo_document', description: GIT_HISTORY_DESCRIPTIONS.redo_document, schema: RedoDocumentSchema },
    { name: 'validate_document', description: PARITY_DESCRIPTIONS.validate_document, schema: ValidateDocumentSchema },
    { name: 'get_stats', description: PARITY_DESCRIPTIONS.get_stats, schema: GetStatsSchema },
    { name: 'sanitize_document', description: PARITY_DESCRIPTIONS.sanitize_document, schema: SanitizeDocumentSchema },
    { name: 'normalize_document', description: PARITY_DESCRIPTIONS.normalize_document, schema: NormalizeDocumentSchema },
    { name: 'import_document', description: PARITY_DESCRIPTIONS.import_document, schema: ImportDocumentSchema },
  ];
}

export function buildStamp(version: string): string {
  return `astra-mcp-server@${version}`;
}

export function generateSchemasToDir(
  targetDir: string,
  version: string = getServerVersion()
): { written: string[]; count: number } {
  if (!fs.existsSync(targetDir)) {
    fs.mkdirSync(targetDir, { recursive: true });
  }
  const stamp = buildStamp(version);
  const written: string[] = [];
  for (const tool of getToolDefinitions()) {
    const raw = z.toJSONSchema(tool.schema as z.ZodType, { target: 'draft-7', io: 'input' }) as Record<string, unknown>;
    if (
      (raw as { definitions?: Record<string, unknown> }).definitions !== undefined &&
      Object.keys((raw as { definitions: Record<string, unknown> }).definitions).length === 0
    ) {
      delete (raw as { definitions?: unknown }).definitions;
    }
    const payload = {
      name: tool.name,
      description: tool.description,
      parameters: raw,
      [SCHEMA_STAMP_KEY]: stamp,
    };
    const filePath = path.join(targetDir, `${tool.name}.json`);
    fs.writeFileSync(filePath, JSON.stringify(payload, null, 2) + '\n', 'utf8');
    written.push(filePath);
  }
  return { written, count: written.length };
}

export interface SchemaMismatch {
  file: string;
  stamped: string;
  expected: string;
}

export function checkSchemasInDir(
  dir: string,
  expectedVersion: string = getServerVersion()
): { ok: boolean; mismatches: SchemaMismatch[]; checked: number } {
  const expected = buildStamp(expectedVersion);
  const mismatches: SchemaMismatch[] = [];
  let checked = 0;
  for (const tool of getToolDefinitions()) {
    const file = path.join(dir, `${tool.name}.json`);
    checked += 1;
    let stamped = '<missing>';
    try {
      const raw = JSON.parse(fs.readFileSync(file, 'utf8')) as Record<string, unknown>;
      const value = raw[SCHEMA_STAMP_KEY];
      stamped = typeof value === 'string' ? value : '<unstamped>';
    } catch {
      stamped = '<missing>';
    }
    if (stamped !== expected) {
      mismatches.push({ file, stamped, expected });
    }
  }
  return { ok: mismatches.length === 0, mismatches, checked };
}

export function formatMismatchWarning(mismatches: SchemaMismatch[]): string {
  const lines = mismatches.map((m) => `  ${path.basename(m.file)}: stamped ${m.stamped}, installed ${m.expected}`);
  return (
    'Schema cache is stale: regenerated wire schemas differ from the installed server. ' +
    'Refresh file-based client copies with `astra-mcp-server --dump-schemas <dir>` ' +
    'and restart memory-cached clients.\n' +
    lines.join('\n')
  );
}

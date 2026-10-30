import { DocumentAST } from '../ast';
import { BaseOperation } from './base';
import { DocxParser } from '../parser';
import { CorePropertiesManipulator } from '../manipulators/core-properties';

export interface MetadataOperation extends BaseOperation {
  action: 'metadata';
  operation: 'setProperties';
  title?: string;
  creator?: string;
  subject?: string;
  description?: string;
  keywords?: string;
  category?: string;
}

const SETTABLE_FIELDS = ['title', 'creator', 'subject', 'description', 'keywords', 'category'] as const;

/**
 * E2 (F-3): partial core-properties update. Only provided fields are
 * touched (AST + docProps/core.xml, synthesized with its Content_Types
 * Override + package rel entry when absent); unprovided fields are
 * preserved. Without a parser only the in-memory AST is updated.
 *
 * Asymmetry (documented): 6 fields are WRITABLE via this op
 * (title|creator|subject|description|keywords|category); 10 fields are
 * READABLE via ast.metadata/inspect (the 6 plus lastModifiedBy|revision|
 * created|modified). revision/lastModifiedBy/created/modified are read-only
 * here — pass them via the op and it rejects loudly (use the strict schema;
 * revision stays Word-managed). `operation` must be 'setProperties'
 * (schema-aligned; no `op` alias fallback).
 */
export async function executeMetadata(
  op: MetadataOperation,
  ast: DocumentAST,
  parser?: DocxParser
): Promise<{ updated: string[] }> {
  const fields = op as unknown as Record<string, unknown>;
  const resolvedOp: unknown = (op as MetadataOperation).operation;
  if (resolvedOp !== 'setProperties') {
    throw new Error(
      `E_VALIDATION: Unknown metadata operation '${String(resolvedOp)}'. Supported: setProperties. ` +
        `Remedy: pass { action: 'metadata', operation: 'setProperties', title: '...' } with at least one of ` +
        `title|creator|subject|description|keywords|category.`
    );
  }
  // Read-only guard: revision/lastModifiedBy/created/modified are inspectable
  // but not settable via this op — reject loudly instead of silently dropping.
  for (const ro of ['lastModifiedBy', 'revision', 'created', 'modified'] as const) {
    if (fields[ro] !== undefined) {
      throw new Error(
        `E_VALIDATION: metadata field '${ro}' is read-only via the metadata op. ` +
          `Remedy: omit '${ro}' (read it via inspect metadata) and set only ` +
          `title|creator|subject|description|keywords|category.`
      );
    }
  }
  const patch: Record<string, string> = {};
  for (const field of SETTABLE_FIELDS) {
    const value: unknown = fields[field];
    if (value === undefined) continue;
    if (typeof value !== 'string') {
      throw new Error(
        `E_VALIDATION: metadata field '${field}' must be a string (got ${typeof value}). ` +
          `Remedy: pass a string value or omit '${field}' to preserve the current value.`
      );
    }
    patch[field] = value;
  }
  if (Object.keys(patch).length === 0) {
    throw new Error(
      `E_VALIDATION: metadata setProperties requires at least one settable field. ` +
        `Remedy: pass at least one of title|creator|subject|description|keywords|category, ` +
        `e.g. { action: 'metadata', operation: 'setProperties', title: 'Q3 Report' }.`
    );
  }
  const adapter = parser?.getZipAdapter();
  const { updated } = new CorePropertiesManipulator(ast, adapter).update(patch);
  return { updated };
}

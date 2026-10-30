import { describe, it, expect } from 'vitest';
import {
  OpenDocumentSchema,
  InspectDocumentSchema,
  SearchContentSchema,
  ExecuteBatchSchema,
  SaveDocumentSchema,
  CloseDocumentSchema,
  GetDocumentHistorySchema,
  RevertDocumentSchema,
  RedoDocumentSchema,
  PreviewRevertSchema,
  ExportAuditTrailSchema,
  ValidateDocumentSchema,
  GetStatsSchema,
  SanitizeDocumentSchema,
  NormalizeDocumentSchema,
  ImportDocumentSchema,
} from '../../src/lib/mcp/schemas';
import { ExportDocumentSchema } from '../../src/tools/export';
import { VisualInspectSchema } from '../../src/tools/visual';
import { NewDocumentSchema } from '../../src/tools/new-document';
import { ComposeDocumentSchema } from '../../src/tools/compose-document';

/**
 * L5: renderer/PDF binary + timeout overrides are CLI-operator-only and must
 * stay unreachable via MCP schemas. Pins the absence of `sofficeBin`,
 * `pdftoppmBin`, `timeoutMs`, and generic `binary` keys on every MCP-wire
 * schema so a future addition fails loudly here instead of silently
 * widening the remote surface.
 */
const FORBIDDEN_KEYS = ['sofficeBin', 'pdftoppmBin', 'timeoutMs', 'binary'] as const;

const SCHEMAS: Array<[string, { shape: Record<string, unknown> }]> = [
  ['OpenDocumentSchema', OpenDocumentSchema as unknown as { shape: Record<string, unknown> }],
  ['InspectDocumentSchema', InspectDocumentSchema as unknown as { shape: Record<string, unknown> }],
  ['SearchContentSchema', SearchContentSchema as unknown as { shape: Record<string, unknown> }],
  ['ExecuteBatchSchema', ExecuteBatchSchema as unknown as { shape: Record<string, unknown> }],
  ['SaveDocumentSchema', SaveDocumentSchema as unknown as { shape: Record<string, unknown> }],
  ['CloseDocumentSchema', CloseDocumentSchema as unknown as { shape: Record<string, unknown> }],
  ['GetDocumentHistorySchema', GetDocumentHistorySchema as unknown as { shape: Record<string, unknown> }],
  ['RevertDocumentSchema', RevertDocumentSchema as unknown as { shape: Record<string, unknown> }],
  ['RedoDocumentSchema', RedoDocumentSchema as unknown as { shape: Record<string, unknown> }],
  ['PreviewRevertSchema', PreviewRevertSchema as unknown as { shape: Record<string, unknown> }],
  ['ExportAuditTrailSchema', ExportAuditTrailSchema as unknown as { shape: Record<string, unknown> }],
  ['ExportDocumentSchema', ExportDocumentSchema as unknown as { shape: Record<string, unknown> }],
  ['VisualInspectSchema', VisualInspectSchema as unknown as { shape: Record<string, unknown> }],
  ['NewDocumentSchema', NewDocumentSchema as unknown as { shape: Record<string, unknown> }],
  ['ComposeDocumentSchema', ComposeDocumentSchema as unknown as { shape: Record<string, unknown> }],
  ['ValidateDocumentSchema', ValidateDocumentSchema as unknown as { shape: Record<string, unknown> }],
  ['GetStatsSchema', GetStatsSchema as unknown as { shape: Record<string, unknown> }],
  ['SanitizeDocumentSchema', SanitizeDocumentSchema as unknown as { shape: Record<string, unknown> }],
  ['NormalizeDocumentSchema', NormalizeDocumentSchema as unknown as { shape: Record<string, unknown> }],
  ['ImportDocumentSchema', ImportDocumentSchema as unknown as { shape: Record<string, unknown> }],
];

describe('L5 MCP schema surface: renderer/PDF overrides stay operator-only', () => {
  for (const [name, schema] of SCHEMAS) {
    it(`${name} exposes no binary/timeout override keys`, () => {
      const keys = Object.keys(schema.shape ?? {});
      for (const forbidden of FORBIDDEN_KEYS) {
        expect(keys, `${name} must not expose ${forbidden}`).not.toContain(forbidden);
      }
    });
  }

  it('VisualInspectSchema keeps only the safe render knobs (no binary paths)', () => {
    const keys = Object.keys(
      (VisualInspectSchema as unknown as { shape: Record<string, unknown> }).shape ?? {}
    );
    // Safe knobs that ARE allowed on the wire (pages/format/jpegQuality/dpi).
    expect(keys).toContain('dpi');
    for (const forbidden of FORBIDDEN_KEYS) {
      expect(keys).not.toContain(forbidden);
    }
  });
});

/**
 * @attentionform/astra-engine
 * Headless, schema-validating, token-optimized OOXML engine for AI agents
 */

// Core Parser & Document Representation
export * from './ast';
export * from './ast-refresh';
export * from './parser';
export * from './document-manager';
export * from './page-map';
export * from './read-cache';
export * from './git';
export * from './recovery';
export * from './grep';
export * from './validator';
export * from './sanitizer';
export * from './ipc-validation';

// Operations & Mutation Engine
export * from './operation-schema';
export * from './operation-engine';
export * from './batch';
export * from './toon';

// OOXML Utilities & Validation Gate
export * from './tmp';
export * from './ooxml/ns';
export * from './ooxml/schema-order';
export * from './ooxml/styles';
export * from './ooxml/typography';
export * from './ooxml/validator-gate';
export * from './ooxml/guideline-bank';
export * from './ooxml/track-changes-emitter';
export * from './ooxml/comments';
export * from './ooxml/omml';

// Manipulators
export * from './manipulators/chart3d';
export * from './manipulators/chart2d';
export * from './manipulators/header-footer';
export * from './manipulators/numbering';
export * from './manipulators/relationships';
export * from './manipulators/toc';
export * from './manipulators/core-properties';

// Exporters
export * from './export/html';
export * from './export/markdown';
export * from './export/pdf';
export { hasExportPdfBackend } from './export/pdf';

// Visual Review & Rendering (Headless)
export {
  hasPdfBackend as hasVisualPdfBackend,
  hasLibreOffice,
  hasPoppler,
  getVisualBackendStatus,
  type VisualBackendStatus,
  type VisualBackendOverride,
} from './visual/backend';
export * from './visual/renderer';
export * from './visual/telemetry';
export * from './visual/prompt';
export * from './visual/provider';

// Hardened in-memory ZIP adapter and hardening utilities
export { default as AdmZip } from './adm-zip-safe';
export * from './zip-hardening';

// Individual Operation Handlers
export * from './operations/base';
export * from './operations/insert';
export * from './operations/delete';
export * from './operations/table';
export * from './operations/chart';
export * from './operations/chart-style';
export * from './operations/style';
export * from './operations/trackedChange';
export * from './operations/replace';
export * from './operations/comment';
export * from './operations/field';
export * from './operations/move';
export * from './operations/coverPage';
export * from './operations/list';
export * from './operations/hyperlink';
export * from './operations/footnote';
export * from './operations/watermark';
export * from './operations/pageBorder';
export * from './operations/section';
export * from './operations/pageSetup';
export * from './operations/namedStyle';
export * from './operations/callout';
export * from './operations/clearFormat';
export * from './operations/image';
export * from './operations/image-file';
export * from './operations/headerFooter';
export * from './operations/equation';
export * from './operations/metadata';

// Compose path (Markdown -> ops -> document) and package normalizer
export { compileMarkdown, composeDocument } from './compose/markdown';
export type { CompileOptions, ComposeOptions } from './compose/markdown';
export { normalizePackage, previewNormalizations } from './ooxml/normalize';
export type { NormalizePackageResult } from './ooxml/normalize';

// Document census counters (words/paragraphs/tables/images + page estimate)
// Single source of truth for CLI `stats` and MCP `get_stats`.
export * from './stats';

// Style guide schema + parser (design system for machine-built documents)
export * from './style-guide/types';
export * from './style-guide/defaults';
export * from './style-guide/parser';
export * from './style-guide/lint';
export * from './style-guide/resolve';
export * from './style-guide/summary';

// Q0a central sandbox jail + Q0b central error table (single source of truth)
export * from './sandbox/paths';
export * from './errors';

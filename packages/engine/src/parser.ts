import fs from 'fs';
import path from 'path';
import AdmZip from './adm-zip-safe';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { DocumentAST, Node as AstNode, calculateNodeChecksum } from './ast';
import { canonicalizeDocumentDom, canonicalizePartXml, generateValidParaId } from './ooxml/schema-order';
import { normalizeDocumentDom, normalizePartXml, previewDocumentDom } from './ooxml/normalize';
import { createWordElement } from './ooxml/ns';
import { ommlToLatex } from './ooxml/omml';
import { validateDocxBuffer, ValidationResult } from './ooxml/validator-gate';
import { NumberingManipulator, NUMBERING_REL_TYPE } from './manipulators/numbering';
import { RelationshipsManipulator } from './manipulators/relationships';
import { CorePropertiesManipulator, parseCorePropertiesXml } from './manipulators/core-properties';
import { StylePalette, extractStylePalette } from './ooxml/styles';
import { clampZipEntryTimes, checkZipBombCeilings, resolveDeterministicZipDate, clampCorePropertiesDates, resolveReproducibleMode } from './zip-hardening';
import { deterministicRevisionIso, resolveDefaultAuthor } from './ooxml/track-changes-emitter';
import { newDocumentParts, newDocumentTitle, NEW_DOCUMENT_CONTENT_TYPE_OVERRIDES, NEW_DOCUMENT_ROOT_RELS, NEW_DOCUMENT_DOCUMENT_RELS } from './ooxml/new-document-parts';
import { assertNoXxe, assertNoXxeAdjacent } from './ooxml/xxe-guard';
import { headerFooterRefIds, pruneOrphanedHeaderFooters, type PruneResult } from './parser/header-footer-prune';
import { isRogueZipEntry, isBlockedExecutableEntry } from './sanitizer';
import {
  MacroEntryStore,
  applyMacroStrip,
  findMacroPartNames,
  macroChoiceError,
  macroRefusalError,
} from './sanitizer';
import { readDocxFileBuffer, getMimeType } from './parser/io';
import { toError } from './errors';
import { upsertSectPrChild } from './parser/sect-pr';
import { addNodeInDocumentOrder, bodySectPr, listFileSections, splitIntoSections } from './parser/sections';
import {
  chartElementDescendants,
  chartPointTextsByIndex,
  chartScatterXValuesBySeries,
  chartTitleElement,
  detectChartKindFromXml,
  findPlotElementLocalNames,
  hasPlotLocalName,
  pairScatterPointTexts,
  scatterSeriesXList,
  scatterXOnlyValues,
} from './parser/factory';
import { readChartFloating, readChartStyle, type ChartStyle } from './operations/chart-style';
import type {
  DocxParserMacroOptions,
  DocxZipAdapter,
  TableMetadata,
  StructuralSummary,
  DocumentSnapshot,
} from './parser/types';
import {
  isValidParaId,
  W14_NS,
  XMLNS_NS,
  parserNowMs,
  parserRandomUnit,
  findDuplicateParaIds,
} from './parser/para-ids';
import type { DuplicateParaId } from './parser/para-ids';
import { buildEmbeddedNodePages } from './parser/node-pages';
import {
  invalidPackageError,
  readPartData,
  invalidXmlError,
} from './parser/package-errors';
import {
  nodeHasTrackedChanges,
  getOrCreateChild,
  unwrapElement,
} from './parser/dom-utils';
import {
  partNameForRelationshipTarget,
  contentTypeForRelationship,
} from './parser/rel-utils';
import { getNodeFromAST } from './parser/ast-utils';
import { detectPersistedCallout } from './parser/callout';
import { hasLiveNumberingForSync } from './parser/numbering';

export { readDocxFileBuffer, getMimeType };
export { SECT_PR_CHILD_ORDER, upsertSectPrChild, linesPerPageFromSectPr } from './parser/sect-pr';
export { detectBar2DChartTypeFromXml, detectChartKindFromXml, findPlotElementLocalNames } from './parser/factory';
export type {
  DocxParserMacroOptions,
  DocxZipAdapter,
  TableMetadata,
  StructuralSummary,
  DocumentSnapshot,
} from './parser/types';
export {
  isValidParaId,
  W14_NS,
  setParserClockForTests,
  setParserRandomForTests,
  resetParserClockForTests,
  findDuplicateParaIds,
  elementXPath,
  paragraphXPath,
} from './parser/para-ids';
export type { DuplicateParaId } from './parser/para-ids';
export { buildEmbeddedNodePages } from './parser/node-pages';
export {
  nodeHasTrackedChanges,
  findDirectChild,
  getOrCreateChild,
  unwrapElement,
} from './parser/dom-utils';
export {
  partNameForRelationshipTarget,
  contentTypeForRelationship,
} from './parser/rel-utils';
export { getNodeFromAST } from './parser/ast-utils';
export { hasLiveNumberingForSync } from './parser/numbering';

/**
 * Document factories (moved from ./parser/factory: they need the
 * DocxParser value, and importing it there created an ESM value cycle
 * parser <-> factory. Defined here they use the local class directly).
 */
/** True when `el` sits inside an element called `tag` (the element itself does not count). */
function hasAncestorTag(el: Element, tag: string): boolean {
  for (let n = el.parentNode; n; n = n.parentNode) {
    if ((n as Element).tagName === tag) return true;
  }
  return false;
}

export function createEmptyDocument(templateType: string = 'blank'): DocxParser {
  return DocxParser.createNewDocument(templateType);
}

export function createBlankDocxBuffer(templateType: string = 'blank'): Buffer {
  const p = DocxParser.createNewDocument(templateType);
  if (templateType === 'blank') {
    const summary = p.getStructuralSummary();
    if (summary.nodes[0]) {
      p.updateNode(summary.nodes[0].id, '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>New Document</w:t></w:r></w:p>');
    }
  }
  return p.rezip();
}

/** Marker prop naming which fail-closed save guard refused (for ok:false mapping). */
export type SaveGuardKind = 'duplicate-paraid';

function guardKindOf(err: unknown): SaveGuardKind | null {
  const g = (err as { astraGuard?: unknown } | null | undefined)?.astraGuard;
  return g === 'duplicate-paraid' ? g : null;
}

/**
 * Maps a fail-closed save-guard refusal to schema-gate error entries so
 * `validateOOXML()` keeps its result contract (`{ok:false}` instead of a
 * throw) for content refusals. Returns null for any other throw (integrity,
 * XXE, macro gates still throw, exactly as before). Both `xpath` and
 * `xPath` casings ride along (CLI reads both).
 */
export function guardFailureToValidationErrors(
  err: unknown
): Array<{ description: string; xpath: string; xPath: string; id?: string; errorType: string }> | null {
  const kind = guardKindOf(err);
  if (kind === 'duplicate-paraid') {
    const dups = (err as { duplicates?: unknown } | null | undefined)?.duplicates;
    const list: DuplicateParaId[] = Array.isArray(dups)
      ? (dups as DuplicateParaId[]).filter(
          (d): d is DuplicateParaId =>
            !!d && typeof (d as DuplicateParaId).paraId === 'string' && Array.isArray((d as DuplicateParaId).xPaths)
        )
      : [];
    return list.slice(0, 10).map((d) => {
      const at = d.xPaths[0] ?? '//w:p';
      return {
        description:
          `duplicate w14:paraId "${d.paraId}" on ${d.count} paragraphs (at ${at}): ` +
          `delete the repeat paragraph (session id ${d.paraId}_dup — see inspect_document) and reinsert its text for a fresh id`,
        xpath: at,
        xPath: at,
        id: d.paraId,
        errorType: 'DuplicateParaId',
      };
    });
  }
  return null;
}

export class DocxParser {
  /**
   * F2a follow-up item 2 — extra template search dirs (packaged hook ONLY).
   * `resolveTemplateCandidatePaths` in electron/main.ts documents the
   * packaged-first order (extraResource → asar → dev); that helper was dead
   * (zero production callers) while this parser only probed cwd/__dirname.
   * Main now pushes those dirs here at startup via `setExtraTemplateDirs`
   * and/or `process.env.STATIC_TEMPLATE_DIRS` (delimiter-separated). Both
   * sources are prepended (packaged-first) ahead of the legacy
   * cwd/__dirname probes in `resolveTemplateSearchPaths` / `createNewDocument`.
   * No other parser logic is touched by this hook.
   */
  private static extraTemplateDirs: string[] = [];

  /** Main-startup hook: prepend packaged template dirs (deduped, resolved). */
  public static setExtraTemplateDirs(dirs: unknown): void {
    if (!Array.isArray(dirs)) {
      DocxParser.extraTemplateDirs = [];
      return;
    }
    const clean: string[] = [];
    for (const d of dirs) {
      if (typeof d !== 'string' || d.trim() === '') continue;
      try {
        const resolved = path.resolve(d.trim());
        if (!clean.includes(resolved)) clean.push(resolved);
      } catch {
        /* ignore unresolvable entries */
      }
    }
    DocxParser.extraTemplateDirs = clean;
  }

  /** Test hook: read back the static extra dirs. */
  public static getExtraTemplateDirs(): string[] {
    return [...DocxParser.extraTemplateDirs];
  }

  /** Test hook: clear the static extra dirs. */
  public static clearExtraTemplateDirs(): void {
    DocxParser.extraTemplateDirs = [];
  }

  /** Env hook: `STATIC_TEMPLATE_DIRS` (path.delimiter-separated dirs). */
  public static getEnvTemplateDirs(): string[] {
    try {
      const raw = (process.env.STATIC_TEMPLATE_DIRS || '').trim();
      if (!raw) return [];
      const parts = raw
        .split(path.delimiter)
        .map((s) => s.trim())
        .filter(Boolean);
      const out: string[] = [];
      for (const p of parts) {
        try {
          const resolved = path.resolve(p);
          if (resolved && !out.includes(resolved)) out.push(resolved);
        } catch {
          /* ignore */
        }
      }
      return out;
    } catch {
      return [];
    }
  }

  /**
   * Ordered template file search paths for `templateType` (packaged-first):
   * static extra dirs → `STATIC_TEMPLATE_DIRS` env dirs → legacy
   * cwd/__dirname probes. Exported for unit tests.
   */
  public static resolveTemplateSearchPaths(templateType: string = 'blank'): string[] {
    const normalizedType = (templateType || 'blank').toLowerCase().trim() || 'blank';
    const fileName = `${normalizedType}.docx`;
    const ordered: string[] = [];
    for (const dir of [...DocxParser.getExtraTemplateDirs(), ...DocxParser.getEnvTemplateDirs()]) {
      try {
        ordered.push(path.resolve(dir, fileName));
      } catch {
        /* ignore */
      }
    }
    ordered.push(
      path.resolve(process.cwd(), 'templates', fileName),
      path.resolve(__dirname, '..', 'templates', fileName),
      path.resolve(__dirname, 'templates', fileName),
    );
    return [...new Set(ordered)];
  }

  private zip: AdmZip;
  private documentDom: Document;
  private stylesDom: Document | null = null;
  private parser: DOMParser;
  private serializer: XMLSerializer;
  /** Public so operation executors can register synthesized nodes (partial
   * test doubles may also carry this map; call sites tolerate absence). */
  nodeMap: Map<string, Element> = new Map();
  /**
   * Identity recursion: stable ids for table-cell content (w:tr/w:tc
   * paragraphs and nested tables). These live in the SAME nodeMap (shared
   * paraId pool + dup-suffix loops, so no collisions with body-level ids)
   * but are tracked here so root-level consumers — getStructuralSummary,
   * getAST, getTrackedNodeIds — keep exact legacy behavior (cell ids stay
   * resolvable via getNode/readNode/updateNode for traceability).
   */
  private cellScopedIds: Set<string> = new Set();
  private mediaCache: Record<string, { base64: string; mimeType: string }> | null = null;

  /**
   * R-10: names of rogue zip entries dropped at parse/save ingress, in
   * arrival order (accessor returns a sorted copy, mirroring the sanitize
   * report). The drop predicate is shared with sanitize by importing
   * isRogueZipEntry/isBlockedExecutableEntry (no duplicated logic): rogue
   * path shapes plus denied executables take the same silent drop+count
   * path here so batch round-trips cannot preserve what sanitize removes.
   * Drops are by entry object without reading entry bytes, so hostile
   * payloads are never decompressed or written anywhere.
   */
  private droppedRogueEntries: string[] = [];

  /**
   * Macro/active-content gate state (refuse-default). Every open (constructor)
   * and every save (`rezip`, incl. the `validateOOXML` gate) enforces the
   * instance disposition: refuse throws `E_VALIDATION` naming the parts,
   * allow preserves them byte-identical, strip removes them with accounting
   * below. Unreferenced `oleObject*.bin` parts stay legit (rel-corroborated
   * detection, see sanitizer).
   */
  private macroMode: 'refuse' | 'allow' | 'strip' = 'refuse';
  /** Macro parts present in the live package (allow mode; [] after a strip). */
  private macroPartNames: string[] = [];
  /** Macro entries removed by strip dispositions, in arrival order. */
  private droppedMacroEntries: string[] = [];

  /** Records a dropped rogue entry name once (deduped, additive only). */
  private noteDroppedRogueEntry(entryName: string): void {
    if (!this.droppedRogueEntries.includes(entryName)) {
      this.droppedRogueEntries.push(entryName);
    }
  }

  /**
   * Deletes every rogue entry currently in the live zip (skips directories,
   * never reads entry data) and records each name. Returns the names newly
   * recorded by this call. Legit parts are untouched: only entries matching
   * the shared sanitizer predicates are removed.
   */
  private purgeRogueZipEntries(): string[] {
    const added: string[] = [];
    for (const entry of [...this.zip.getEntries()]) {
      if (entry.isDirectory) continue;
      const name = entry.entryName;
      if (!isRogueZipEntry(name) && !isBlockedExecutableEntry(name)) continue;
      try {
        this.zip.deleteFile(entry);
      } catch {
        continue;
      }
      if (!this.droppedRogueEntries.includes(name)) {
        this.droppedRogueEntries.push(name);
        added.push(name);
      }
    }
    return added;
  }

  /**
   * Deletes every macro entry currently in the live zip under a strip
   * disposition (bins + companions + macro-only rel files) and sweeps macro
   * Relationship elements / content-type overrides from kept files. Returns
   * the names newly recorded by this call.
   */
  private purgeMacroParts(): string[] {
    const added: string[] = [];
    const store: MacroEntryStore = {
      names: () => this.zip.getEntries().filter((e) => !e.isDirectory).map((e) => e.entryName),
      readText: (name: string) => {
        try {
          const e = this.zip.getEntry(name);
          return e ? e.getData().toString('utf8') : null;
        } catch {
          return null;
        }
      },
      remove: (name: string) => {
        const e = this.zip.getEntry(name);
        if (e) {
          try {
            this.zip.deleteFile(e);
          } catch {
            /* best-effort */
          }
        }
      },
      writeText: (name: string, text: string) => {
        this.zip.updateFile(name, Buffer.from(text, 'utf8'));
      },
    };
    const { removed } = applyMacroStrip(store);
    for (const name of removed) {
      if (!this.droppedMacroEntries.includes(name)) {
        this.droppedMacroEntries.push(name);
        added.push(name);
      }
    }
    return added;
  }

  /** Scans the live zip for macro parts (sorted names, [] when none). */
  private listLiveMacroParts(): string[] {
    const names: string[] = [];
    for (const entry of this.zip.getEntries()) {
      if (!entry.isDirectory) names.push(entry.entryName);
    }
    return findMacroPartNames(names, (n) => {
      try {
        const e = this.zip.getEntry(n);
        return e ? e.getData().toString('utf8') : null;
      } catch {
        return null;
      }
    });
  }

  /**
   * Enforces the instance macro disposition against the live package.
   * Refuse + parts present throws `E_VALIDATION` naming the parts (zero
   * bytes change); allow records them; strip removes them. No-op when the
   * package carries no active content, so clean documents behave exactly as
   * before (byte-identical).
   */
  private enforceMacroGate(verb: 'open' | 'save'): void {
    const parts = this.listLiveMacroParts();
    if (parts.length === 0) {
      this.macroPartNames = [];
      return;
    }
    if (this.macroMode === 'allow') {
      this.macroPartNames = parts;
      return;
    }
    if (this.macroMode === 'strip') {
      this.purgeMacroParts();
      this.macroPartNames = [];
      return;
    }
    throw macroRefusalError(parts, verb);
  }

  /**
   * Applies an explicit macro choice to this parser (open/save option
   * surface). Both flags together throw. Allow records macro parts for
   * byte-identical round-trips; strip removes them now. Returns the entry
   * names removed by this call ([] for allow / clean packages).
   */
  public applyMacroChoice(opts: DocxParserMacroOptions, verb: 'open' | 'save' = 'save'): string[] {
    if (opts.allowMacros === true && opts.stripMacros === true) throw macroChoiceError();
    if (opts.allowMacros === true) this.macroMode = 'allow';
    else if (opts.stripMacros === true) this.macroMode = 'strip';
    const before = new Set(this.droppedMacroEntries);
    this.enforceMacroGate(verb);
    return this.droppedMacroEntries.filter((n) => !before.has(n)).sort();
  }

  /** Macro parts present in the live package (sorted; [] after a strip). */
  public getMacroParts(): string[] {
    return [...this.macroPartNames].sort();
  }

  /**
   * Macro entries removed by strip dispositions so far (sorted). Additive:
   * entries are only appended, so load-time strips stay visible after save.
   */
  public getDroppedMacroEntries(): string[] {
    return [...this.droppedMacroEntries].sort();
  }

  /**
   * R-10 accounting accessor: sorted names dropped at ingress so far
   * (length is the drop count). Additive: entries are only appended, never
   * removed or reset, so pre-batch drops stay visible after save. Clean
   * documents report []. The batch-result count derives from this list;
   * engine-side result wiring lives outside this file.
   */
  public getDroppedRogueEntries(): string[] {
    return [...this.droppedRogueEntries].sort();
  }

  /**
   * BX-03 route-around fix (D1 follow-up): bare rezip() is the production save
   * path (document-manager.ts:91 createNewDocument, :155 saveAsDocument,
   * :236 commitTransaction, :259 updateNode, :354 undoLastAction, :421
   * redoLastAction — all call rezip() with no AST), while syncNumberingToZip
   * previously ran only when an explicit AST was passed. List definitions
   * minted in-memory by NumberingManipulator.getOrCreateList therefore survived
   * rezip(ast) in tests but were silently dropped on every prod save.
   *
   * Fix within parser scope (cannot touch document-manager/operations):
   * getAST() caches the returned AST object reference here. The AST is mutated
   * in place (NumberingManipulator writes into ast.numbering maps), so the
   * cached reference observes later list.create defs without any caller change.
   * Bare rezip() falls back to this cached AST for NUMBERING ONLY (part + its
   * single rel + single content-type override). Media/relationships are NOT
   * synced from the stale cache — those merges are additive-only and would
   * resurrect deleted media/rels if replayed from a stale snapshot.
   * Cleared on reloadDocumentXml/restorePackageSnapshot (zip numbering changed
   * underneath; next getAST() repopulates). Single reference, overwritten on
   * each getAST(), so no unbounded growth.
   */
  private lastAstForNumbering: DocumentAST | null = null;

  /**
   * Allowlist-normalizer log from the most recent rezip() (empty when the
   * package was already canonical — clean saves stay byte-identical).
   * Replaced on every rezip; auxiliary part entries carry a `name: ` prefix.
   * Read-only snapshot for the future envelope/CLI display slice.
   */
  public lastNormalizations: string[] = [];

  /** Test hook: forget the cached AST (forces bare rezip to preserve zip as-is). */
  public clearCachedNumberingAst(): void {
    this.lastAstForNumbering = null;
  }

  /** Test hook: whether a cached AST is available for bare-rezip numbering sync. */
  public hasCachedNumberingAst(): boolean {
    return this.lastAstForNumbering !== null;
  }

  /**
   * D1R2-3 (rollback-divergence hook): re-points the bare-rezip numbering
   * cache at a caller-owned AST object (normally the OperationEngine's
   * this.ast after it rewinds in place).
   *
   * WHY AN ENGINE-SIDE LINE IS REQUIRED (exact spec, engine file untouched):
   * OperationEngine.restoreAfterFailedBatch (operation-engine.ts:182-193)
   * calls parser.restorePackageSnapshot() (which MUST clear this cache — the
   * zip was rewound underneath it) and then parser.getAST(), whose fresh
   * object becomes the new cache. The engine, however, keeps mutating its OWN
   * this.ast object (sections spliced + maps re-keyed IN PLACE, identity
   * preserved). From that point the parser cache (fresh) and the engine AST
   * (this.ast) are DIFFERENT objects: numbering defs minted via
   * NumberingManipulator(engineAst) after any further getAST() are invisible
   * to bare rezip(), which reads only the cache. No parser-side read can
   * distinguish "engine's live object" from any other stale island, and making
   * getAST() always return one stable object would turn the engine's own
   * rewind (splice-into-self + delete-then-assign maps onto themselves) into a
   * map wipe — so the fix must come from the engine side. EXACT SPEC (one
   * line, appended at the end of restoreAfterFailedBatch's parser-attached
   * branch, after `this.ast.nextChangeId = fresh.nextChangeId;`):
   *
   *   this.parser.adoptAstForNumbering(this.ast);
   *
   * Post-adopt, bare rezip() follows the engine's object again. Safe to call
   * redundantly (idempotent pointer swap, no copies). Never call with a stale
   * snapshot: adopting an island whose numbering was rolled back would replay
   * discarded defs on the next bare save.
   */
  public adoptAstForNumbering(ast: DocumentAST): void {
    this.lastAstForNumbering = ast;
  }

  public invalidateMediaCache(): void { this.mediaCache = null; }

  /**
   * G2-6 nodePages perf: AST revision + cached embedded map.
   * getStructuralSummary() is broadcast up to 3x per op (document-manager +
   * electron IPC + batch progress), and the embedded nodePages pass is
   * O(N·words) word-splitting. Cache it by revision/identity and recompute
   * only when nodes actually change. Revision bumps on every mutator below;
   * the cache key additionally carries a cheap ids+lengths fingerprint so a
   * missed bump can never serve a stale map (correctness over speed).
   */
  private _nodePagesRevision = 0;
  private _nodePagesCache: {
    rev: number;
    sectPr: string;
    nodeCount: number;
    fingerprint: string;
    nodePages: Record<string, number>;
    totalPages: number;
    linesPerPage: number;
  } | null = null;
  private _nodePagesComputes = 0;
  private _nodePagesCacheHits = 0;
  private _bumpNodePagesRevision(): void {
    try {
      this._nodePagesRevision += 1;
    } catch { /* ignore */ }
  }
  /** Current AST revision for nodePages cache invalidation (monotonic). */
  public getNodePagesRevision(): number {
    return this._nodePagesRevision;
  }
  /** Cache telemetry for tests: recomputes vs cheap hits. */
  public getNodePagesCacheStats(): { computes: number; hits: number; revision: number; cached: boolean } {
    return {
      computes: this._nodePagesComputes,
      hits: this._nodePagesCacheHits,
      revision: this._nodePagesRevision,
      cached: this._nodePagesCache !== null,
    };
  }
  /** Test hook: clear the embedded-map cache without bumping revision. */
  public clearNodePagesCache(): void {
    this._nodePagesCache = null;
  }

  constructor(input: Buffer | string, opts?: DocxParserMacroOptions) {
    let buffer: Buffer;
    if (typeof input === 'string') {
      buffer = readDocxFileBuffer(input);
    } else {
      buffer = input;
    }
    // We load entirely in memory, which intrinsically prevents Zip Slip path traversal 
    // to the local disk during extraction.
    try {
      this.zip = new AdmZip(buffer);
    } catch (err) {
      throw invalidPackageError(err);
    }
    checkZipBombCeilings(this.zip, { sourceCompressedBytes: buffer.length });
    // R-10: drop rogue entries at parse ingress (after the bomb ceiling, so
    // hostile sizes still count fail-closed) before any part is read.
    this.purgeRogueZipEntries();
    // Macro/active-content gate at open ingress: an explicit choice preserves
    // (allow) or removes (strip) macro parts; default refuses with
    // E_VALIDATION naming the parts before any part is read.
    if (opts?.allowMacros === true && opts?.stripMacros === true) throw macroChoiceError();
    if (opts?.allowMacros === true) this.macroMode = 'allow';
    else if (opts?.stripMacros === true) this.macroMode = 'strip';
    this.enforceMacroGate('open');
    // H2 open-ingress sweep over every stored XML/rels part (classic DTD +
    // XXE-adjacent XInclude/XSLT/stylesheet-PI). document.xml/styles.xml are
    // re-asserted below with part-specific context; the sweep guarantees an
    // auxiliary part (header, theme, chart, ...) cannot smuggle a vector past
    // open into downstream LibreOffice renders. Well-formed OOXML never
    // carries these shapes, so legit packages pass untouched.
    for (const entry of this.zip.getEntries()) {
      const name = entry.entryName;
      if (isRogueZipEntry(name) || isBlockedExecutableEntry(name)) continue;
      if (name.endsWith('.xml') || name.endsWith('.rels')) {
        const text = readPartData(entry, name).toString('utf8');
        assertNoXxe(text, name);
        assertNoXxeAdjacent(text, name);
      }
    }
    this.parser = new DOMParser({
      errorHandler: {
        warning: (w) => { throw invalidXmlError(w) },
        error: (e) => { throw invalidXmlError(e) },
        fatalError: (f) => { throw invalidXmlError(f) }
      }
    });
    this.serializer = new XMLSerializer();
    
    const entry = this.zip.getEntry('word/document.xml');
    if (!entry) throw toError('E_VALIDATION', 'Invalid DOCX: missing word/document.xml', 're-export the file from Word (or pick a valid .docx) and retry');
    
    const xmlStr = readPartData(entry, 'word/document.xml').toString('utf8');

    assertNoXxe(xmlStr, 'word/document.xml');
    assertNoXxeAdjacent(xmlStr, 'word/document.xml');

    // Security: Parse using xmldom safely. It prevents basic string injection because we manipulate AST.
    this.documentDom = this.parser.parseFromString(xmlStr, "text/xml");

    const stylesEntry = this.zip.getEntry('word/styles.xml');
    if (stylesEntry) {
      const stylesXml = readPartData(stylesEntry, 'word/styles.xml').toString('utf8');
      assertNoXxe(stylesXml, 'word/styles.xml');
      assertNoXxeAdjacent(stylesXml, 'word/styles.xml');
      this.stylesDom = this.parser.parseFromString(stylesXml, "text/xml");
    }
    this.assignStableIds();
  }

  public reloadDocumentXml(xmlStr: string): void {
    assertNoXxe(xmlStr, 'word/document.xml');
    assertNoXxeAdjacent(xmlStr, 'word/document.xml');
    const newDoc = this.parser.parseFromString(xmlStr, "text/xml");
    const parserError = newDoc.getElementsByTagName('parsererror');
    if (parserError.length > 0) {
      throw new Error("XML Injection or Invalid XML detected in input.");
    }
    this.documentDom = newDoc;
    this.assignStableIds();
    this._bumpNodePagesRevision();
    // BX-03: zip content changed underneath the cached AST reference — drop it
    // so a later bare rezip() preserves the reloaded zip instead of replaying
    // stale in-memory numbering defs over it.
    this.lastAstForNumbering = null;
  }

  public createAtomicSnapshot(label: string = 'Snapshot', opts?: { nowMs?: number; seed?: string | number }): DocumentSnapshot {
    const xmlContent = this.readDocumentXml();
    const stylesXmlContent = this.stylesDom ? this.serializer.serializeToString(this.stylesDom) : null;
    const now = typeof opts?.nowMs === 'number' && Number.isFinite(opts.nowMs) ? opts.nowMs : parserNowMs();
    let randSuffix: number;
    if (opts?.seed !== undefined) {
      const s = String(opts.seed);
      let h = 2166136261;
      for (let i = 0; i < s.length; i++) {
        h ^= s.charCodeAt(i);
        h = Math.imul(h, 16777619);
      }
      randSuffix = Math.abs(h) % 10000;
    } else {
      randSuffix = Math.floor(parserRandomUnit() * 10000);
    }
    return {
      id: `snap_${now}_${randSuffix}`,
      timestamp: now,
      label,
      xmlContent,
      stylesXmlContent
    };
  }

  public restoreAtomicSnapshot(snapshot: DocumentSnapshot): void {
    this.reloadDocumentXml(snapshot.xmlContent);
    if (snapshot.stylesXmlContent && this.stylesDom) {
      assertNoXxe(snapshot.stylesXmlContent, 'word/styles.xml');
      assertNoXxeAdjacent(snapshot.stylesXmlContent, 'word/styles.xml');
      this.stylesDom = this.parser.parseFromString(snapshot.stylesXmlContent, "text/xml");
    }
  }

  /**
   * Captures a full package snapshot: a clone of every zip entry buffer.
   * Used by OperationEngine for transactional rollback of failed batches.
   */
  public capturePackageSnapshot(): Map<string, Buffer> {
    // R-10: batch ingress drops rogues at snapshot time so a pre-batch
    // snapshot can neither carry them forward nor resurrect them on rewind.
    this.purgeRogueZipEntries();
    const snapshot = new Map<string, Buffer>();
    for (const entry of this.zip.getEntries()) {
      if (isRogueZipEntry(entry.entryName) || isBlockedExecutableEntry(entry.entryName)) {
        this.noteDroppedRogueEntry(entry.entryName);
        continue;
      }
      snapshot.set(entry.entryName, Buffer.from(entry.getData()));
    }
    // A-4: the live document DOM leads the zip bytes (mints + edits land in
    // the DOM; the zip is only flushed at rezip). Snapshot the live
    // serialization so rollback restores the exact pre-batch state instead
    // of reloading stale bytes and re-minting fresh w14:paraIds.
    try {
      snapshot.set('word/document.xml', Buffer.from(this.readDocumentXml(), 'utf8'));
    } catch { /* fall back to the zip bytes on serialization failure */ }
    return snapshot;
  }

  /** Relationship ids of every header and footer the document refers to right now. */
  public headerFooterReferenceIds(): Set<string> {
    return headerFooterRefIds(this.documentDom);
  }

  /**
   * Removes header and footer parts that an edit left unreferenced (a deleted
   * section break), so their text cannot leave inside the saved file. Only
   * parts in `referencedBefore` and no longer referenced are candidates.
   * Never throws: when anything looks unsafe it removes nothing. Also drops
   * the matching entries from `ast`, since rezip writes the AST's
   * relationships and media back into the package.
   */
  public pruneOrphanedHeaderFooters(ast: DocumentAST, referencedBefore: ReadonlySet<string>): PruneResult {
    const zip = this.zip;
    const result = pruneOrphanedHeaderFooters({
      parts: {
        names: () => zip.getEntries().map((entry) => entry.entryName),
        read: (name) => {
          const entry = zip.getEntry(name);
          return entry ? Buffer.from(entry.getData()) : null;
        },
        write: (name, data) => {
          if (zip.getEntry(name)) zip.updateFile(name, data);
          else zip.addFile(name, data);
        },
        remove: (name) => {
          const entry = zip.getEntry(name);
          if (entry) zip.deleteFile(entry);
        },
      },
      documentXml: () => this.readDocumentXml(),
      referencedNow: this.headerFooterReferenceIds(),
      referencedBefore,
      guard: (xml, name) => {
        assertNoXxe(xml, name);
        assertNoXxeAdjacent(xml, name);
      },
    });
    if (result.parts.length === 0) return result;
    // Null-safe on purpose: the package is already pruned, so this must not throw half way.
    for (const id of result.relationshipIds) if (ast.relationships) delete ast.relationships[id];
    for (const part of result.parts) {
      if (part.startsWith('word/media/')) {
        const base = part.slice(part.lastIndexOf('/') + 1);
        if (ast.mediaFiles) for (const key of [part, part.replace(/^word\//, ''), base]) delete ast.mediaFiles[key];
        continue;
      }
      const file = part.slice('word/'.length);
      const map = file.startsWith('header') ? ast.headers : ast.footers;
      if (!map) continue;
      const removed = map[file];
      for (const key of Object.keys(map)) {
        if (removed && map[key] === removed) delete map[key];
      }
      const field = file.startsWith('header') ? 'headerText' : 'footerText';
      if (removed?.text && ast.metadata?.[field] === removed.text) delete ast.metadata[field];
    }
    this.invalidateMediaCache();
    return result;
  }

  /**
   * Restores the package to a previously captured snapshot: every entry is
   * restored (entries added after the snapshot are removed), then documentDom
   * and stylesDom are reloaded and stable IDs re-assigned.
   */
  public restorePackageSnapshot(snapshot: Map<string, Buffer>): void {
    // Remove entries added since the snapshot
    for (const entry of [...this.zip.getEntries()]) {
      if (!snapshot.has(entry.entryName)) {
        this.zip.deleteFile(entry);
      }
    }
    // Restore / add captured entries (R-10: rogue names in an incoming
    // snapshot are dropped+counted, never written back into the package).
    for (const [entryName, data] of snapshot) {
      if (isRogueZipEntry(entryName) || isBlockedExecutableEntry(entryName)) {
        this.noteDroppedRogueEntry(entryName);
        const live = this.zip.getEntry(entryName);
        if (live) {
          try {
            this.zip.deleteFile(live);
          } catch { /* best-effort */ }
        }
        continue;
      }
      const buffer = Buffer.from(data);
      if (this.zip.getEntry(entryName)) {
        this.zip.updateFile(entryName, buffer);
      } else {
        this.zip.addFile(entryName, buffer);
      }
    }

    // Reload document DOM via the existing guarded reload path
    const docEntry = this.zip.getEntry('word/document.xml');
    if (!docEntry) throw toError('E_VALIDATION', 'Invalid DOCX: missing word/document.xml', 're-export the file from Word (or pick a valid .docx) and retry');
    const xmlStr = readPartData(docEntry, 'word/document.xml').toString('utf8');
    assertNoXxe(xmlStr, 'word/document.xml');
    assertNoXxeAdjacent(xmlStr, 'word/document.xml');
    const newDoc = this.parser.parseFromString(xmlStr, 'text/xml');
    const parserError = newDoc.getElementsByTagName('parsererror');
    if (parserError.length > 0) {
      throw new Error('XML Injection or Invalid XML detected in input.');
    }
    this.documentDom = newDoc;

    // Reload styles DOM with the same XXE guards as the constructor
    const stylesEntry = this.zip.getEntry('word/styles.xml');
    if (stylesEntry) {
      const stylesXml = readPartData(stylesEntry, 'word/styles.xml').toString('utf8');
      assertNoXxe(stylesXml, 'word/styles.xml');
      assertNoXxeAdjacent(stylesXml, 'word/styles.xml');
      this.stylesDom = this.parser.parseFromString(stylesXml, 'text/xml');
    } else {
      this.stylesDom = null;
    }

    this.invalidateMediaCache();
    this.assignStableIds();
    this._bumpNodePagesRevision();
    // BX-03: same invalidation as reloadDocumentXml — the zip numbering part
    // may have been rewound; the cached AST reference is now stale.
    this.lastAstForNumbering = null;
  }

  private assignStableIds() {
    this.nodeMap.clear();
    this.cellScopedIds.clear();
    const body = this.documentDom.getElementsByTagName('w:body')[0];
    if (!body) return;

    let tCounter = 1;
    let nestedTCounter = 1;

    // A-4: pre-scan valid native paraIds so minted ids never collide.
    const usedParaIds = new Set<string>();
    const allParas = this.documentDom.getElementsByTagName('w:p');
    for (let i = 0; i < allParas.length; i++) {
      const v = (allParas[i] as unknown as Element).getAttribute?.('w14:paraId');
      if (v && isValidParaId(v)) usedParaIds.add(v.toUpperCase());
    }
    let mintedParaId = false;
    let mintSeq = 0;
    const mintParaId = (node: Element): string => {
      let id: string;
      let attempt = 0;
      const text = node.textContent || '';
      const seed = `load_p_${++mintSeq}_${text}`;
      do {
        id = generateValidParaId(seed, attempt);
        attempt++;
      } while (usedParaIds.has(id));
      usedParaIds.add(id);
      // setAttributeNS (not setAttribute): a namespace-less 'w14:paraId'
      // makes xmldom re-declare xmlns:w14 on every w:p at save time,
      // breaking byte-level checksum stability across save/reload.
      node.setAttributeNS(W14_NS, 'w14:paraId', id);
      mintedParaId = true;
      return id;
    };

    const registerPara = (node: Element, cellScoped = false) => {
      // A-4: reuse MS Word's native paraId when present and valid (Word 2013+);
      // otherwise mint a valid w14:paraId (ST_LongHexNumber, high bit clear) so
      // the id survives save/reopen cycles instead of positional p_NNN churn.
      // NOTE: paraId is an attribute of w:p itself (CT_P), not a w:pPr child.
      let id = node.getAttribute('w14:paraId');
      if (!id || !isValidParaId(id)) {
        id = mintParaId(node);
      } else if (this.nodeMap.has(id)) {
        // Duplicate w14:paraId — generate a fresh id instead of overwriting the map entry
        let candidate = `${id}_dup`;
        let n = 2;
        while (this.nodeMap.has(candidate)) {
          candidate = `${id}_dup${n}`;
          n++;
        }
        id = candidate;
      }

      node.setAttribute('data-astra-id', id);
      this.nodeMap.set(id, node);
      if (cellScoped) this.cellScopedIds.add(id);
    };

    const registerTable = (node: Element, cellScoped = false) => {
      if (cellScoped) {
        // Nested tables: cell-scoped, excluded from root consumers, and kept
        // off the body t_ counter so body t_ numbering matches legacy order.
        let id = `t_cell_${nestedTCounter.toString().padStart(3, '0')}`;
        let n = 1;
        while (this.nodeMap.has(id)) {
          id = `t_cell_${nestedTCounter.toString().padStart(3, '0')}_dup${n > 1 ? n : ''}`;
          n++;
        }
        node.setAttribute('data-astra-id', id);
        this.nodeMap.set(id, node);
        this.cellScopedIds.add(id);
        nestedTCounter++;
        return;
      }
      let id = `t_${tCounter.toString().padStart(3, '0')}`;
      let n = 1;
      while (this.nodeMap.has(id)) {
        id = `t_${tCounter.toString().padStart(3, '0')}_dup${n > 1 ? n : ''}`;
        n++;
      }
      node.setAttribute('data-astra-id', id);
      this.nodeMap.set(id, node);
      tCounter++;
    };

    const directKids = (parent: Element, tag: string): Element[] => {
      const out: Element[] = [];
      for (let i = 0; i < parent.childNodes.length; i++) {
        const child = parent.childNodes[i] as Element;
        if (child.nodeType === 1 && child.tagName === tag) out.push(child);
      }
      return out;
    };

    // Identity recursion: every w:p inside w:tr/w:tc gets the same stable
    // identity as a body paragraph (native w14:paraId reuse or minted id
    // from the shared pool — collisions with the body map are impossible
    // by construction). Nested tables register cell-scoped (t_cell_ ids,
    // off the body t_ counter) and recurse. Cell-scoped ids are flagged
    // so root-level consumers skip them while direct lookups
    // (getNode/readNode/updateNode) resolve them.
    const registerTableCellContent = (tbl: Element): void => {
      for (const tr of directKids(tbl, 'w:tr')) {
        for (const tc of directKids(tr, 'w:tc')) {
          for (let i = 0; i < tc.childNodes.length; i++) {
            const child = tc.childNodes[i] as Element;
            if (child.nodeType !== 1) continue;
            if (child.tagName === 'w:p') {
              registerPara(child, true);
            } else if (child.tagName === 'w:tbl') {
              registerTable(child, true);
              registerTableCellContent(child);
            } else if (child.tagName === 'w:sdt') {
              for (let j = 0; j < child.childNodes.length; j++) {
                const sdtChild = child.childNodes[j] as Element;
                if (sdtChild.nodeType === 1 && sdtChild.tagName === 'w:sdtContent') {
                  for (let k = 0; k < sdtChild.childNodes.length; k++) {
                    const inner = sdtChild.childNodes[k] as Element;
                    if (inner.nodeType !== 1) continue;
                    if (inner.tagName === 'w:p') registerPara(inner, true);
                    else if (inner.tagName === 'w:tbl') {
                      registerTable(inner, true);
                      registerTableCellContent(inner);
                    }
                  }
                }
              }
            }
          }
        }
      }
    };

    const processContainer = (container: Element) => {
      for (let i = 0; i < container.childNodes.length; i++) {
        const node = container.childNodes[i] as Element;
        if (node.nodeType !== 1 /* ELEMENT_NODE */) continue;

        if (node.tagName === 'w:p') {
          registerPara(node);
        } else if (node.tagName === 'w:tbl') {
          registerTable(node);
          registerTableCellContent(node);
        } else if (node.tagName === 'w:sdt') {
          // Recurse into structured-document-tag content so paragraphs/tables
          // wrapped in w:sdt > w:sdtContent are addressable instead of skipped.
          for (let j = 0; j < node.childNodes.length; j++) {
            const sdtChild = node.childNodes[j] as Element;
            if (sdtChild.nodeType === 1 && sdtChild.tagName === 'w:sdtContent') {
              processContainer(sdtChild);
            }
          }
        }
      }
    };

    processContainer(body);

    // A-4: a minted w14:paraId needs its namespace declared on the root,
    // otherwise the saved part carries an undeclared prefix.
    if (mintedParaId) {
      const root = this.documentDom.documentElement;
      if (root && !root.hasAttribute('xmlns:w14')) {
        root.setAttributeNS(XMLNS_NS, 'xmlns:w14', W14_NS);
      }
    }
  }

  /**
   * A-4: mints a w14:paraId not present anywhere in the live document.
   * Used for inserted w:p roots (load-time minting uses assignStableIds'
   * pre-scanned set instead to stay O(n)). Unseeded fallback is
   * deterministic (content-counter seed) when reproducible, random only
   * when opted out.
   */
  private generateUniqueParaId(seed?: string): string {
    const used = new Set<string>();
    const allParas = this.documentDom.getElementsByTagName('w:p');
    for (let i = 0; i < allParas.length; i++) {
      const v = (allParas[i] as unknown as Element).getAttribute?.('w14:paraId');
      if (v && isValidParaId(v)) used.add(v.toUpperCase());
    }
    let id: string;
    let attempt = 0;
    const fallbackSeed = `parser-fallback:${allParas.length}:${this.documentDom.documentElement?.textContent?.length ?? 0}`;
    do {
      if (seed) id = generateValidParaId(seed, attempt);
      else if (resolveReproducibleMode(undefined)) id = generateValidParaId(fallbackSeed, attempt);
      else id = generateValidParaId();
      attempt++;
    } while (used.has(id));
    return id;
  }

  public getDocumentDom(): Document {
    return this.documentDom;
  }
  public getDom(): Document {
    return this.documentDom;
  }

  public getStructuralSummary(): StructuralSummary {
    const summary: StructuralSummary = {
      hasTrackedChanges: false,
      stylesXml: this.stylesDom ? this.serializer.serializeToString(this.stylesDom) : null,
      headers: [],
      footers: [],
      charts: [],
      relationships: {},
      nodes: []
    };

    // Extract relationships
    const relsEntry = this.zip.getEntry('word/_rels/document.xml.rels');
    if (relsEntry) {
      const relsXml = relsEntry.getData().toString('utf8');
      assertNoXxe(relsXml, 'word/_rels/document.xml.rels');
      assertNoXxeAdjacent(relsXml, 'word/_rels/document.xml.rels');
      try {
        const relsDoc = this.parser.parseFromString(relsXml, 'text/xml');
        const relNodes = relsDoc.getElementsByTagName('Relationship');
        for (let i = 0; i < relNodes.length; i++) {
          const rel = relNodes[i];
          const rId = rel.getAttribute('Id');
          if (rId) {
            summary.relationships![rId] = {
              type: rel.getAttribute('Type') || '',
              target: rel.getAttribute('Target') || ''
            };
          }
        }
      } catch (e) {
        if (e instanceof Error && (/XXE|DTD|E_VALIDATION/.test(e.message))) throw e;
        // Ignore rels parse error
      }
    }

    // Extract media files lazily into mediaCache
    if (!this.mediaCache) {
      this.mediaCache = {};
      for (const entry of this.zip.getEntries()) {
        // R-10: rogue bytes never enter the media cache (no read, counted).
        if (isRogueZipEntry(entry.entryName) || isBlockedExecutableEntry(entry.entryName)) {
          this.noteDroppedRogueEntry(entry.entryName);
          continue;
        }
        if (entry.entryName.startsWith('word/media/') && !entry.isDirectory) {
          const data = entry.getData();
          const mimeType = getMimeType(entry.entryName);
          const base64Data = data.toString('base64');
          const base64 = `data:${mimeType};base64,${base64Data}`;
          const item = { base64, mimeType };
          this.mediaCache[entry.entryName] = item;
          this.mediaCache[path.basename(entry.entryName)] = item;
          this.mediaCache[entry.entryName.replace(/^word\//, '')] = item;
        }
      }
    }
    summary.mediaFiles = { ...this.mediaCache };

    // Extract headers, footers, and charts from zip entries
    // ST-09: key headers/footers by sectPr reference type (first/even/default)
    // instead of hardcoding 'default'. Falls back to 'default' when no sectPr
    // reference matches (legacy docs without section refs).
    const headerRefTypes = new Map<string, string>();
    const footerRefTypes = new Map<string, string>();
    try {
      const bodyEl = this.documentDom.getElementsByTagName('w:body')[0];
      const sectPrEls = listFileSections(bodyEl).flatMap((fileSection) => (fileSection.sectPr ? [fileSection.sectPr] : []));
      const relsForType = (kind: 'header' | 'footer'): Map<string, string> => {
        const m = new Map<string, string>();
        const tag = kind === 'header' ? 'w:headerReference' : 'w:footerReference';
        // Every section's references, first section first; a part keeps the type of the first section that uses it.
        for (const sectPrEl of sectPrEls) {
          const refs = sectPrEl.getElementsByTagName(tag);
          for (let i = 0; i < refs.length; i++) {
            const rId = refs[i].getAttribute('r:id') || '';
            const t = (refs[i].getAttribute('w:type') || 'default').toLowerCase();
            if (rId && !m.has(rId)) m.set(rId, t === 'first' || t === 'even' ? t : 'default');
          }
        }
        return m;
      };
      // Build target→type via relationships (target like header1.xml).
      const relsEntryForTypes = this.zip.getEntry('word/_rels/document.xml.rels');
      const relTargetById = new Map<string, string>();
      if (relsEntryForTypes) {
        const relsRawForTypes = relsEntryForTypes.getData().toString('utf8');
        assertNoXxe(relsRawForTypes, 'word/_rels/document.xml.rels');
        assertNoXxeAdjacent(relsRawForTypes, 'word/_rels/document.xml.rels');
        try {
          const relsDoc = this.parser.parseFromString(relsRawForTypes, 'text/xml');
          const relNodes = relsDoc.getElementsByTagName('Relationship');
          for (let i = 0; i < relNodes.length; i++) {
            const id = relNodes[i].getAttribute('Id') || '';
            const tgt = relNodes[i].getAttribute('Target') || '';
            if (id) relTargetById.set(id, tgt);
          }
        } catch (e) {
          if (e instanceof Error && (/XXE|DTD|E_VALIDATION/.test(e.message))) throw e;
          /* ignore */
        }
      }
      const hRefs = relsForType('header');
      const fRefs = relsForType('footer');
      for (const [rId, t] of hRefs) {
        const tgt = relTargetById.get(rId) || '';
        const base = tgt.split('/').pop() || tgt;
        if (base) headerRefTypes.set(base.toLowerCase(), t);
        if (tgt) headerRefTypes.set(tgt.toLowerCase(), t);
      }
      for (const [rId, t] of fRefs) {
        const tgt = relTargetById.get(rId) || '';
        const base = tgt.split('/').pop() || tgt;
        if (base) footerRefTypes.set(base.toLowerCase(), t);
        if (tgt) footerRefTypes.set(tgt.toLowerCase(), t);
      }
    } catch { /* ignore — fallback to default below */ }
    for (const entry of this.zip.getEntries()) {
      if (/^word\/(header\d+\.xml)$/i.test(entry.entryName)) {
        const base = entry.entryName.split('/').pop()!.toLowerCase();
        const headerXml = entry.getData().toString('utf8');
        assertNoXxe(headerXml, entry.entryName);
        assertNoXxeAdjacent(headerXml, entry.entryName);
        summary.headers!.push({
          id: entry.entryName,
          type: headerRefTypes.get(base) || headerRefTypes.get(entry.entryName.toLowerCase()) || 'default',
          xml: headerXml
        });
      } else if (/^word\/(footer\d+\.xml)$/i.test(entry.entryName)) {
        const base = entry.entryName.split('/').pop()!.toLowerCase();
        const footerXml = entry.getData().toString('utf8');
        assertNoXxe(footerXml, entry.entryName);
        assertNoXxeAdjacent(footerXml, entry.entryName);
        summary.footers!.push({
          id: entry.entryName,
          type: footerRefTypes.get(base) || footerRefTypes.get(entry.entryName.toLowerCase()) || 'default',
          xml: footerXml
        });
      } else if (/^word\/charts\/(chart\d+\.xml)$/i.test(entry.entryName)) {
        const chartXml = entry.getData().toString('utf8');
        assertNoXxe(chartXml, entry.entryName);
        assertNoXxeAdjacent(chartXml, entry.entryName);
        summary.charts!.push({
          id: entry.entryName,
          xml: chartXml
        });
      }
    }

    // ST-04/ST-05: minimally plumb theme1.xml + numbering.xml for renderer fidelity.
    // Both hit the central guard like rels do: an XXE throw propagates instead
    // of being swallowed into a silent null (fail-closed, never cleaned-quietly).
    try {
      const themeEntry = this.zip.getEntry('word/theme/theme1.xml') || this.zip.getEntry('word/theme/theme2.xml');
      const themeXml = themeEntry ? themeEntry.getData().toString('utf8') : null;
      if (themeEntry && themeXml) { assertNoXxe(themeXml, themeEntry.entryName); assertNoXxeAdjacent(themeXml, themeEntry.entryName); }
      summary.themeXml = themeXml;
    } catch (err: unknown) {
      if (err instanceof Error && (/XXE|DTD|E_VALIDATION/.test(err.message))) throw err;
      summary.themeXml = null;
    }
    try {
      const numEntry = this.zip.getEntry('word/numbering.xml');
      const numberingXml = numEntry ? numEntry.getData().toString('utf8') : null;
      if (numberingXml) { assertNoXxe(numberingXml, 'word/numbering.xml'); assertNoXxeAdjacent(numberingXml, 'word/numbering.xml'); }
      summary.numberingXml = numberingXml;
    } catch (err: unknown) {
      if (err instanceof Error && (/XXE|DTD|E_VALIDATION/.test(err.message))) throw err;
      summary.numberingXml = null;
    }

    // Extract pageSetup from body sectPr
    const body = this.documentDom.getElementsByTagName('w:body')[0];
    const sectPr = listFileSections(body)[0].sectPr;
    if (sectPr) {
      const sectPrXml = this.serializer.serializeToString(sectPr);
      summary.pageSetup = {
        rawXml: sectPrXml
      };
      // ST-09 G1b: plumb parity flags from the same sectPr (w:titlePg,
      // w:evenAndOddHeaders gate first/even use; never even-on-odd).
      try {
        summary.sectPrFlags = {
          titlePg: /<w:titlePg(?:\s[^>]*)?(?:\/>|>)/i.test(sectPrXml),
          evenAndOddHeaders: /<w:evenAndOddHeaders(?:\s[^>]*)?(?:\/>|>)/i.test(sectPrXml)
        };
      } catch { summary.sectPrFlags = { titlePg: false, evenAndOddHeaders: false }; }
    } else {
      try { summary.sectPrFlags = { titlePg: false, evenAndOddHeaders: false }; } catch { /* ignore */ }
    }

    this.nodeMap.forEach((node, id) => {
      // Identity recursion: cell-scoped ids stay out of the root summary —
      // table text is already surfaced via the table node itself.
      if (this.cellScopedIds.has(id)) return;
      const rawXml = this.serializer.serializeToString(node);
      
      // Extract text
      const textNodes = node.getElementsByTagName('w:t');
      let text = '';
      for (let i = 0; i < textNodes.length; i++) {
        text += textNodes[i].textContent || '';
      }

      // Check if heading
      let isHeading = false;
      let headingLevel: number | undefined = undefined;
      const pStyle = node.getElementsByTagName('w:pStyle')[0];
      if (pStyle) {
        const val = pStyle.getAttribute('w:val') || '';
        const lowerVal = val.toLowerCase();
        if (lowerVal.includes('heading') || lowerVal.includes('title') || lowerVal.includes('subtitle')) {
          isHeading = true;
          if (lowerVal.includes('heading 1') || lowerVal === 'heading1' || lowerVal === 'h1' || lowerVal.includes('title')) {
            headingLevel = 1;
          } else if (lowerVal.includes('heading 2') || lowerVal === 'heading2' || lowerVal === 'h2' || lowerVal.includes('subtitle')) {
            headingLevel = 2;
          } else if (lowerVal.includes('heading 3') || lowerVal === 'heading3' || lowerVal === 'h3') {
            headingLevel = 3;
          } else {
            headingLevel = 1;
          }
        }
      }

      const hasTrackedChanges = nodeHasTrackedChanges(node);
      if (hasTrackedChanges) summary.hasTrackedChanges = true;

      let type: StructuralSummary['nodes'][number]['type'] = node.tagName === 'w:p' ? 'paragraph' : 'table';
      let preview = text.substring(0, 65) + (text.length > 65 ? '...' : '');
      let metadata: TableMetadata | undefined = undefined;

      if (node.tagName === 'w:tbl') {
        const trs = node.getElementsByTagName('w:tr');
        const rows = trs.length;
        const firstTr = trs[0];
        const firstTrTcs = firstTr ? firstTr.getElementsByTagName('w:tc') : null;
        const cols = firstTrTcs?.length || 2;
        preview = `Table [${rows}x${cols}]`;

        const headers: string[] = [];
        if (firstTrTcs) {
          for (let c = 0; c < firstTrTcs.length; c++) {
            const tc = firstTrTcs[c];
            const tcTexts = tc.getElementsByTagName('w:t');
            let cellText = '';
            for (let t = 0; t < tcTexts.length; t++) {
              cellText += tcTexts[t].textContent || '';
            }
            headers.push(cellText.trim());
          }
        }

        // Check internal OOXML captions/titles in tblPr
        let caption: string | undefined = undefined;
        const tblPr = node.getElementsByTagName('w:tblPr')[0];
        if (tblPr) {
          const tblCaptionEl = tblPr.getElementsByTagName('w:tblCaption')[0];
          const tblTitleEl = tblPr.getElementsByTagName('w:tblTitle')[0];
          const tblDescEl = tblPr.getElementsByTagName('w:tblDescription')[0];
          caption = tblCaptionEl?.getAttribute('w:val') ||
                    tblTitleEl?.getAttribute('w:val') ||
                    tblDescEl?.getAttribute('w:val') ||
                    undefined;
        }

        metadata = {
          headers: headers.some(h => h.length > 0) ? headers : undefined,
          tableDimensions: { rows, cols },
          caption
        };
      } else if (node.tagName === 'w:p' && text.length === 0) {
        if (node.getElementsByTagName('w:drawing').length > 0) {
          const isChart = node.getElementsByTagName('c:chart').length > 0 || rawXml.includes('drawingml/2006/chart');
          if (isChart) {
            // P2 loop-back: resolve the referenced chart part and sniff it —
            // 2D parts are 'chart', only true 3D parts stay 'chart3d'.
            let summaryIs3D = false;
            try {
              const cChartEl = node.getElementsByTagName('c:chart')[0] as Element | undefined;
              const cRId = cChartEl?.getAttribute?.('r:id') || '';
              const relTarget = (cRId && summary.relationships?.[cRId]?.target) || '';
              const base = String(relTarget.split('/').pop() || '').toLowerCase();
              const part = (summary.charts || []).find((c: { id: string; xml: string }) =>
                base !== '' && String(c?.id || '').split('/').pop()?.toLowerCase() === base);
              const partXml = part?.xml || '';
              summaryIs3D = partXml ? detectChartKindFromXml(partXml).is3D : /<c:\w*3DChart[\s>]|<c:view3D>/.test(rawXml);
            } catch { summaryIs3D = /<c:\w*3DChart[\s>]|<c:view3D>/.test(rawXml); }
            type = summaryIs3D ? 'chart3d' : 'chart';
            preview = summaryIs3D ? '[3D Chart Visualization]' : '[Chart]';
          } else {
            type = 'image';
            preview = '[Embedded Image]';
          }
        } else if (node.getElementsByTagName('w:br').length > 0) {
          const brs = node.getElementsByTagName('w:br');
          for (let b = 0; b < brs.length; b++) {
            if (brs[b].getAttribute('w:type') === 'page') {
              type = 'pageBreak';
              preview = '[Page Break]';
              break;
            }
          }
        }
      }

      summary.nodes.push({
        id,
        type,
        preview,
        text,
        rawXml,
        charCount: text.length,
        isHeading,
        headingLevel,
        hasTrackedChanges,
        metadata
      });
    });

    // IN-07 + G2-6: embed a renderer-compatible nodePages map (+ geometry
    // params via pageSetup) so visual_review can resolve node_ids without the
    // fixed-45 heuristic diverging from canvas pagination. Explicit page breaks
    // consume exactly like the canvas paginator; density fallback scales with
    // the sectPr geometry. Labeled as an estimate until a live renderer map
    // overrides it downstream.
    // PERF: getStructuralSummary is broadcast up to 3x per op, and the
    // embedded pass is O(N·words). Cache by revision/identity (dirty-flag):
    // recompute only when nodes change; repeat broadcasts are O(1) hits.
    try {
      const sectPrXml: string = summary.pageSetup?.rawXml ?? '';
      // Cheap identity fingerprint: ids + text lengths (no word-splitting, so
      // O(N) integers vs the O(N·words) embedded pass). Guards against a missed
      // revision bump ever serving a stale map.
      let fingerprint = '';
      try {
        const parts: string[] = new Array(summary.nodes.length);
        for (let i = 0; i < summary.nodes.length; i++) {
          const n: StructuralSummary['nodes'][number] = summary.nodes[i];
          parts[i] = `${n?.id ?? ''}:${(n?.text ?? '').length}:${(n?.type ?? '')}`;
        }
        // Simple FNV-ish hash to keep the key small for large docs.
        let h = 2166136261;
        const joined = parts.join('|');
        for (let i = 0; i < joined.length; i++) {
          h ^= joined.charCodeAt(i);
          h = Math.imul(h, 16777619);
        }
        fingerprint = `${summary.nodes.length}:${(h >>> 0).toString(36)}`;
      } catch {
        fingerprint = `${summary.nodes.length}`;
      }
      const cached = this._nodePagesCache;
      if (
        cached &&
        cached.rev === this._nodePagesRevision &&
        cached.sectPr === sectPrXml &&
        cached.nodeCount === summary.nodes.length &&
        cached.fingerprint === fingerprint
      ) {
        this._nodePagesCacheHits += 1;
        summary.nodePages = cached.nodePages;
        summary.totalPages = cached.totalPages;
        summary.pageMapSource = 'estimate';
        summary.pageMapEstimate = true;
        // linesPerPage is derivable but not part of the summary contract; the
        // cached totalPages/nodePages are the broadcast-critical fields.
      } else {
        this._nodePagesComputes += 1;
        const embedded = buildEmbeddedNodePages(summary.nodes, summary.pageSetup?.rawXml);
        summary.nodePages = embedded.nodePages;
        summary.totalPages = embedded.totalPages;
        summary.pageMapSource = 'estimate';
        summary.pageMapEstimate = true;
        this._nodePagesCache = {
          rev: this._nodePagesRevision,
          sectPr: sectPrXml,
          nodeCount: summary.nodes.length,
          fingerprint,
          nodePages: embedded.nodePages,
          totalPages: embedded.totalPages,
          linesPerPage: embedded.linesPerPage,
        };
      }
    } catch { /* best-effort — summary stays valid without a map */ }

    return summary;
  }

  public sanitizePath(filePath: string): string {
    const normalized = filePath.replace(/\\/g, '/');
    if (normalized.includes('..') || normalized.startsWith('/') || /^[a-zA-Z]:/.test(normalized)) {
      throw new Error(`Security Violation: Zip Slip / path traversal attempt detected in ${filePath}`);
    }
    return normalized;
  }

  public getZipAdapter(): DocxZipAdapter {
    return {
      hasFile: (path: string) => !!this.zip.getEntry(this.sanitizePath(path)),
      readText: (path: string) => {
        const cleanPath = this.sanitizePath(path);
        const entry = this.zip.getEntry(cleanPath);
        if (!entry) throw new Error(`File not found: ${cleanPath}`);
        return entry.getData().toString('utf8');
      },
      readTextOrNull: (path: string) => {
        const cleanPath = this.sanitizePath(path);
        const entry = this.zip.getEntry(cleanPath);
        return entry ? entry.getData().toString('utf8') : null;
      },
      writeText: (path: string, content: string) => {
        const cleanPath = this.sanitizePath(path);
        assertNoXxe(content, cleanPath);
        assertNoXxeAdjacent(content, cleanPath);
        if (this.zip.getEntry(cleanPath)) {
          this.zip.updateFile(cleanPath, Buffer.from(content, 'utf8'));
        } else {
          this.zip.addFile(cleanPath, Buffer.from(content, 'utf8'));
        }
      },
      writeBuffer: (path: string, content: Buffer) => {
        const cleanPath = this.sanitizePath(path);
        // Same gate as writeText: decode at the utf8 storage boundary (see
        // the UTF-16 note in ooxml/xxe-guard) and fail closed on DTD markup.
        assertNoXxe(content.toString('utf8'), cleanPath);
        assertNoXxeAdjacent(content.toString('utf8'), cleanPath);
        if (this.zip.getEntry(cleanPath)) {
          this.zip.updateFile(cleanPath, content);
        } else {
          this.zip.addFile(cleanPath, content);
        }
      }
    };
  }

  public getTrackedNodeIds(nodeIds?: string[]): string[] {
    const idsToCheck = nodeIds ? new Set(nodeIds) : null;
    const tracked: string[] = [];
    this.nodeMap.forEach((node, id) => {
      // Identity recursion: cell-scoped ids keep legacy root-only behavior.
      if (this.cellScopedIds.has(id)) return;
      if (idsToCheck && !idsToCheck.has(id)) return;
      if (nodeHasTrackedChanges(node)) tracked.push(id);
    });
    return tracked;
  }

  // --- AST layer: stable id assignment ---

  public async getAST(): Promise<DocumentAST> {
    const ast: DocumentAST = {
      metadata: {},
      sections: [],
      styles: {},
      numbering: { abstractNum: {}, concreteNum: {} },
      relationships: {},
      headers: {},
      footers: {},
      mediaFiles: {},
      settings: {
        trackChanges: false,
        defaultAuthor: resolveDefaultAuthor()
      },
      nextChangeId: 1
    };

    // Blank and parsed documents default to tracking OFF. A real
    // <w:trackRevisions/> flag in word/settings.xml opts back in
    // (an explicit w:val="false" keeps it OFF).
    try {
      const settingsXml = this.getZipAdapter().readTextOrNull('word/settings.xml');
      if (settingsXml) {
        assertNoXxe(settingsXml, 'word/settings.xml');
        assertNoXxeAdjacent(settingsXml, 'word/settings.xml');
        const flag = /<w:trackRevisions(?=[\s/>])[^>]*>/.exec(settingsXml);
        if (flag && !/w:val\s*=\s*["'](false|0|off|no)["']/i.test(flag[0])) {
          ast.settings.trackChanges = true;
        }
      }
    } catch (err: unknown) {
      if (err instanceof Error && (/XXE|DTD|E_VALIDATION/.test(err.message))) throw err;
      /* unreadable settings part: tracking stays OFF */
    }

    // Populate the relationships map from word/_rels/document.xml.rels so that
    // rId allocation (RelationshipsManipulator.calculateMaxRId) sees real rels
    // and never collides with existing entries such as rId1 = styles.xml.
    const docRelsXml = this.getZipAdapter().readTextOrNull('word/_rels/document.xml.rels');
    if (docRelsXml) {
      assertNoXxe(docRelsXml, 'word/_rels/document.xml.rels');
      assertNoXxeAdjacent(docRelsXml, 'word/_rels/document.xml.rels');
      const relsDoc = this.parser.parseFromString(docRelsXml, 'text/xml');
      const rels = relsDoc.getElementsByTagName('Relationship');
      for (let i = 0; i < rels.length; i++) {
        const rel = rels[i];
        const rId = rel.getAttribute('Id');
        if (!rId) continue;
        const targetModeAttr = rel.getAttribute('TargetMode');
        ast.relationships[rId] = {
          type: rel.getAttribute('Type') || '',
          target: rel.getAttribute('Target') || '',
          ...(targetModeAttr ? { targetMode: targetModeAttr as 'External' | 'Internal' } : {})
        };
      }
    }

    // Body-level blocks in document order; grouped into the file's sections
    // once the section properties have been read (see parser/sections.ts).
    const rootNodes: AstNode[] = [];

    // BX-03: hydrate ast.numbering from the stored part so list definitions
    // survive across sessions and rezip can write them back.
    try {
      const numberingXml = this.getZipAdapter().readTextOrNull('word/numbering.xml');
      if (numberingXml) {
        new NumberingManipulator(ast).parseNumbering(numberingXml);
      }
    } catch (err: unknown) {
      if (err instanceof Error && (/XXE|DTD|E_VALIDATION/.test(err.message))) throw err;
      /* numbering stays empty — document simply has no lists */
    }

    // E2 (F-1): hydrate ast.metadata from docProps/core.xml. A missing file
    // or missing fields leaves the corresponding fields undefined — never
    // throws here (many docs lack core.xml entirely). XXE/DTD payloads and
    // malformed core XML (E_VALIDATION) stay loud (fail-closed, matches the
    // document.xml/styles.xml posture).
    try {
      const coreXml = this.getZipAdapter().readTextOrNull('docProps/core.xml');
      if (coreXml) {
        Object.assign(ast.metadata, parseCorePropertiesXml(coreXml));
      }
    } catch (err: unknown) {
      if (err instanceof Error && (/XXE|DTD|E_VALIDATION/.test(err.message))) throw err;
      /* core props stay empty — document simply has none */
    }

    for (const [id, element] of this.nodeMap.entries()) {
      // Identity recursion: cell-scoped paragraphs are addressable via
      // getNode/readNode but are NOT root AST nodes (the table owns them).
      if (this.cellScopedIds.has(id)) continue;
      const xmlStr = this.serializer.serializeToString(element);
      const checksum = await calculateNodeChecksum(xmlStr);
      
      // Extract text content cleanly
      const textNodes = element.getElementsByTagName('w:t');
      let runText = '';
      for (let i = 0; i < textNodes.length; i++) {
        runText += textNodes[i].textContent || '';
      }
      // Inline math (an m:oMath outside an m:oMathPara) reads as \(latex\) in the text of a paragraph that also has
      // run text, in document order with the runs around it. A paragraph that is only math (an equation node) keeps
      // its empty text, like a display equation, because exports read that text as the whole formula.
      let text = runText;
      const inlineMath = Array.from(element.getElementsByTagName('m:oMath')).filter(
        (m) => !hasAncestorTag(m as Element, 'm:oMathPara')
      ) as Element[];
      if (inlineMath.length > 0 && runText.length > 0) {
        text = '';
        const all = element.getElementsByTagName('*');
        for (let i = 0; i < all.length; i++) {
          const el = all[i] as Element;
          if (el.tagName === 'w:t' && !hasAncestorTag(el, 'm:oMath')) text += el.textContent || '';
          else if (el.tagName === 'm:oMath' && !hasAncestorTag(el, 'm:oMathPara')) {
            const latex = ommlToLatex(el) ?? Array.from(el.getElementsByTagName('m:t')).map((t) => t.textContent || '').join('');
            text += `\\(${latex}\\)`;
          }
        }
      }
      
      const pStyleEl = element.getElementsByTagName('w:pStyle')[0];
      const namedStyle = pStyleEl ? pStyleEl.getAttribute('w:val') || undefined : undefined;

      let nodeType: 'paragraph' | 'table' | 'chart' | 'chart3d' | 'image' | 'equation' | 'field' | 'list' = element.tagName === 'w:p' ? 'paragraph' : 'table';
      let chartData: { title?: string; categories?: string[]; series?: { name: string; values: number[]; xValues?: number[]; unpairedX?: number; unpairedY?: number }[]; chartType?: string; holeSize?: number; style?: ChartStyle; synthetic?: boolean } | undefined = undefined;
      // P2 loop-back: provisional 2D default; refined by sniffing the loaded
      // chart part below (only true 3D parts become 'chart3d').
      let detectedChart: { is3D: boolean; chartType: string } | undefined = undefined;

      if (xmlStr.includes('http://schemas.openxmlformats.org/drawingml/2006/chart') || xmlStr.includes('c:chart')) {
        nodeType = 'chart';
        const chartEl = element.getElementsByTagName('c:chart')[0];
        const rId = chartEl?.getAttribute('r:id');
        if (rId && this.zip) {
          const relsEntry = this.zip.getEntry('word/_rels/document.xml.rels');
          if (relsEntry) {
            try {
              const relsRaw = relsEntry.getData().toString('utf8');
              assertNoXxe(relsRaw, 'word/_rels/document.xml.rels');
              assertNoXxeAdjacent(relsRaw, 'word/_rels/document.xml.rels');
              const relsDoc = this.parser.parseFromString(relsRaw, 'text/xml');
              const relNodes = relsDoc.getElementsByTagName('Relationship');
              for (let i = 0; i < relNodes.length; i++) {
                if (relNodes[i].getAttribute('Id') === rId) {
                  const target = relNodes[i].getAttribute('Target') || '';
                  const chartPath = target.startsWith('word/') ? target : `word/${target}`;
                  const chartEntry = this.zip.getEntry(chartPath);
                  if (chartEntry) {
                    const chartXml = chartEntry.getData().toString('utf8');
                    assertNoXxe(chartXml, chartPath);
                    assertNoXxeAdjacent(chartXml, chartPath);
                    // P2 loop-back: element sniffing of the loaded chart part —
                    // never a hardcoded 3D label.
                    detectedChart = detectChartKindFromXml(chartXml);
                    nodeType = detectedChart.is3D ? 'chart3d' : 'chart';
                    const chartDoc = this.parser.parseFromString(chartXml, 'text/xml');
                    const chartRoot = chartDoc?.documentElement as Element | undefined;
                    const titleEl = chartRoot ? chartTitleElement(chartRoot) : undefined;
                    const titleText = titleEl ? (titleEl.getElementsByTagName('a:t')[0]?.textContent ?? undefined) : undefined;
                    // Provenance: never fabricate a title. A missing title
                    // stays undefined; it does not make the part synthetic.
                    const title: string | undefined = titleText || undefined;
                    const holeEl = chartRoot ? chartElementDescendants(chartRoot, 'holeSize')[0] : undefined;
                    const holeSize = holeEl ? parseInt(holeEl.getAttribute('val') || '', 10) : undefined;
                    const pickCache = (parent: Element, names: string[]): Element | undefined => {
                      for (const name of names) {
                        const found = chartElementDescendants(parent, name)[0];
                        if (found) return found;
                      }
                      return undefined;
                    };
                    const pointValues = (container: Element): string[] => chartElementDescendants(container, 'pt')
                      .map((pt) => chartElementDescendants(pt, 'v')[0]?.textContent ?? '');
                    // Support both referenced caches and inline literals:
                    // strCache (inside strRef) first, strLit as fallback.
                    // xVal/yVal fallbacks are scatter-only: bubble parts also
                    // use them but are not a supported shared-x family.
                    const isScatterPart = detectedChart?.chartType === 'scatter';
                    const scatterXText = isScatterPart && chartRoot
                      ? chartScatterXValuesBySeries(chartRoot)
                      : [];
                    const catEl = chartRoot ? chartElementDescendants(chartRoot, 'cat')[0] : undefined;
                    const catContainer = catEl
                      ? pickCache(catEl, ['strCache', 'strLit'])
                      : undefined;
                    let categories: string[] = isScatterPart
                      ? (scatterSeriesXList(scatterXText, 0) ?? (catContainer ? pointValues(catContainer) : []))
                      : catContainer ? pointValues(catContainer) : [];
                    const series: { name: string; values: number[]; xValues?: number[]; unpairedX?: number; unpairedY?: number; chartType?: 'column' | 'line'; axis?: 'primary' | 'secondary' }[] = [];
                    const serEls = chartRoot ? chartElementDescendants(chartRoot, 'ser') : [];
                    const comboPart = detectedChart?.chartType === 'combo';
                    if (comboPart) {
                      serEls.sort((a, b) => {
                        const ai = Number(chartElementDescendants(a, 'idx')[0]?.getAttribute('val') ?? Number.NaN);
                        const bi = Number(chartElementDescendants(b, 'idx')[0]?.getAttribute('val') ?? Number.NaN);
                        return ai - bi;
                      });
                    }
                    for (let sIdx = 0; sIdx < serEls.length; sIdx++) {
                      const ser = serEls[sIdx];
                      const txEl = chartElementDescendants(ser, 'tx')[0];
                      const nameMatch = txEl ? chartElementDescendants(txEl, 'v')[0]?.textContent : undefined;
                      const sName = nameMatch || 'Series';
                      // numCache (inside numRef) first, numLit as fallback.
                      const valEl = chartElementDescendants(ser, 'val')[0];
                      const yValEl = isScatterPart && !valEl ? chartElementDescendants(ser, 'yVal')[0] : undefined;
                      const valContainer = valEl
                        ? pickCache(valEl, ['numCache', 'numLit'])
                        : yValEl
                          ? pickCache(yValEl, ['numCache', 'numLit'])
                          : undefined;
                      // Scatter reads x and y by c:pt idx, then pairs only
                      // points where both sides exist, so a sparse list can
                      // never shift one coordinate onto another point.
                      const xOverride = isScatterPart ? scatterSeriesXList(scatterXText, sIdx) : undefined;
                      if (isScatterPart && xOverride !== undefined && valContainer) {
                        const paired = pairScatterPointTexts(xOverride, chartPointTextsByIndex(valContainer));
                        series.push({
                          name: sName,
                          values: paired.values,
                          ...(paired.xValues.length > 0 ? { xValues: paired.xValues } : {}),
                          ...(paired.unpairedX > 0 ? { unpairedX: paired.unpairedX } : {}),
                          ...(paired.unpairedY > 0 ? { unpairedY: paired.unpairedY } : {}),
                        });
                        continue;
                      }
                      if (isScatterPart && xOverride !== undefined) {
                        // No y container: the whole x list is unpaired, but it
                        // still belongs to the series and must be reported.
                        const xOnly = scatterXOnlyValues(xOverride);
                        series.push({
                          name: sName,
                          values: [],
                          ...(xOnly.xValues.length > 0 ? { xValues: xOnly.xValues } : {}),
                          ...(xOnly.unpairedX > 0 ? { unpairedX: xOnly.unpairedX } : {}),
                        });
                        continue;
                      }
                      const values = valContainer
                        ? pointValues(valContainer).map((v) => parseFloat(v) || 0)
                        : [];
                      let comboType: 'column' | 'line' | undefined;
                      let comboAxis: 'primary' | 'secondary' | undefined;
                      if (comboPart) {
                        let plot: Element | null = ser.parentNode as Element | null;
                        while (plot && plot.localName !== 'barChart' && plot.localName !== 'lineChart') plot = plot.parentNode as Element | null;
                        if (plot) {
                          comboType = plot.localName === 'lineChart' ? 'line' : 'column';
                          const axIds = chartElementDescendants(plot, 'axId').map((el) => el.getAttribute('val'));
                          comboAxis = axIds.includes('2001') || axIds.includes('2002') ? 'secondary' : 'primary';
                        }
                      }
                      series.push({
                        name: sName,
                        values,
                        ...(comboType ? { chartType: comboType } : {}),
                        ...(comboAxis ? { axis: comboAxis } : {}),
                      });
                    }
                    // Scatter categories mirror the first series' paired x
                    // list, so holes a sparse c:xVal list dropped cannot leave
                    // categories.length out of step with xValues.length.
                    if (isScatterPart && series.length > 0 && series[0].xValues !== undefined) {
                      categories = series[0].xValues.map((value) => String(value));
                    }
                    // Provenance: never fabricate. A chart is synthetic only
                    // when data could not be hydrated: no shared categories
                    // and no per-series x list, no series, or no plot type.
                    // A missing title alone does not hide real data.
                    const hasPlot = findPlotElementLocalNames(chartXml).length > 0 || hasPlotLocalName(chartXml, 'view3D');
                    const resolvedChartType: string | undefined = hasPlot ? detectedChart?.chartType : undefined;
                    const hasAnyX = categories.length > 0 || series.some((s) => s.xValues !== undefined && s.xValues.length > 0);
                    const synthetic = !hasAnyX || series.length === 0 || !resolvedChartType;
                    let chartStyle: ChartStyle | undefined;
                    try { chartStyle = readChartStyle(chartXml); } catch { chartStyle = undefined; }
                    let chartFloating: ChartStyle['floating'];
                    try { chartFloating = readChartFloating(element); } catch { chartFloating = undefined; }
                    if (chartFloating !== undefined) chartStyle = { ...(chartStyle ?? {}), floating: chartFloating };
                    chartData = {
                      ...(title !== undefined ? { title } : {}),
                      categories,
                      series,
                      chartType: resolvedChartType,
                      ...(holeSize !== undefined ? { holeSize } : {}),
                      ...(chartStyle !== undefined ? { style: chartStyle } : {}),
                      ...(synthetic ? { synthetic: true } : {}),
                    };
                  }
                  break;
                }
              }
            } catch (e) {
              if (e instanceof Error && (/XXE|DTD|E_VALIDATION/.test(e.message))) throw e;
            }
          }
        }
      }

      // B-7 node-type retention on reload: a w:p carrying Office Math is an
      // equation node, and an otherwise-empty (non-chart) w:p carrying a
      // real embedded picture (w:drawing + a:blip) is an image node. Either
      // condition alone is insufficient: shape-only drawings (w:drawing with
      // no a:blip) and text+image paragraphs keep 'paragraph'. Without this
      // every reload demotes both to 'paragraph'.
      if (nodeType === 'paragraph' && element.tagName === 'w:p') {
        // runText, not text: math text must neither turn a pure equation paragraph into a paragraph nor hide an image.
        const hasDisplayMath = element.getElementsByTagName('m:oMathPara').length > 0;
        const hasInlineMath = element.getElementsByTagName('m:oMath').length > 0;
        if (hasDisplayMath || (hasInlineMath && runText.length === 0)) {
          nodeType = 'equation';
        } else if (
          runText.length === 0 &&
          element.getElementsByTagName('w:drawing').length > 0 &&
          element.getElementsByTagName('a:blip').length > 0
        ) {
          nodeType = 'image';
        }
      }

      // Persisted-callout rebuild: save strips data-astra-id, so a reopened
      // callout table carries no marker — its structure is re-read here into
      // the same metadata.callout shape live inserts carry.
      const persistedCallout = element.tagName === 'w:tbl' && nodeType === 'table'
        ? detectPersistedCallout(element)
        : undefined;

      const node: AstNode = {
        id,
        type: nodeType,
        xml: element,
        metadata: {
          text: persistedCallout ? persistedCallout.text : text,
          ...(persistedCallout
            ? { callout: { theme: persistedCallout.theme, title: persistedCallout.title } }
            : {}),
          namedStyle,
          chartData,
          categories: chartData?.categories,
          series: chartData?.series,
          // Provenance: no 'column' default. Unknown stays undefined so TOON
          // renders explicit-unknown ([CH:?]) instead of a fabricated type.
          chartType: chartData?.chartType ?? detectedChart?.chartType,
          hasCitation: xmlStr.includes('w:fldSimple') && xmlStr.includes('CITE'),
          hasComment: xmlStr.includes('w:commentReference'),
          hasBookmark: xmlStr.includes('w:bookmarkStart'),
          hasTrackedChanges: nodeHasTrackedChanges(element),
          hasEquation: xmlStr.includes('m:oMath'),
          hasField: xmlStr.includes('w:instrText') || xmlStr.includes('w:fldSimple'),
          hasImage: xmlStr.includes('w:drawing'),
          hasHyperlink: xmlStr.includes('w:hyperlink')
        },
        version: 1,
        checksum
      };
      
      rootNodes.push(node);
    }
    
    // Extract media files into AST
    for (const entry of this.zip.getEntries()) {
      if (entry.entryName.startsWith('word/media/') && !entry.isDirectory) {
        const buf = entry.getData();
        ast.mediaFiles[entry.entryName] = buf;
        ast.mediaFiles[path.basename(entry.entryName)] = buf;
        ast.mediaFiles[entry.entryName.replace(/^word\//, '')] = buf;
      }
    }

    // Extract headers and footers
    const body = this.documentDom.getElementsByTagName('w:body')[0];
    const fileSections = listFileSections(body);
    // The first section's properties feed the parity flags below.
    const sectPr = fileSections[0].sectPr;

    // Every section's references count, first section first: a part only a
    // later section references still gets its real type (first/even/default),
    // and a part several sections share keeps the type of the first to use it.
    const headerRefs: Array<{ type: 'default' | 'first' | 'even'; rId: string }> = [];
    const footerRefs: Array<{ type: 'default' | 'first' | 'even'; rId: string }> = [];
    for (const fileSection of fileSections) {
      if (!fileSection.sectPr) continue;
      const hRefs = fileSection.sectPr.getElementsByTagName('w:headerReference');
      for (let i = 0; i < hRefs.length; i++) {
        const hEl = hRefs[i];
        const rId = hEl.getAttribute('r:id') || '';
        const hType = (hEl.getAttribute('w:type') || 'default') as 'default' | 'first' | 'even';
        if (rId) headerRefs.push({ type: hType, rId });
      }
      const fRefs = fileSection.sectPr.getElementsByTagName('w:footerReference');
      for (let i = 0; i < fRefs.length; i++) {
        const fEl = fRefs[i];
        const rId = fEl.getAttribute('r:id') || '';
        const fType = (fEl.getAttribute('w:type') || 'default') as 'default' | 'first' | 'even';
        if (rId) footerRefs.push({ type: fType, rId });
      }
    }

    for (const entry of this.zip.getEntries()) {
      const headerMatch = entry.entryName.match(/^word\/(header\d+\.xml)$/i);
      const footerMatch = entry.entryName.match(/^word\/(footer\d+\.xml)$/i);
      if (headerMatch) {
        try {
          const xmlStr = entry.getData().toString('utf8');
          assertNoXxe(xmlStr, entry.entryName);
          assertNoXxeAdjacent(xmlStr, entry.entryName);
          const doc = this.parser.parseFromString(xmlStr, 'text/xml');
          const textNodes = doc.getElementsByTagName('w:t');
          let text = '';
          for (let i = 0; i < textNodes.length; i++) {
            text += textNodes[i].textContent || '';
          }
          text = text.trim();
          const fileName = headerMatch[1];
          const matchedRelId = Object.keys(ast.relationships).find(id => {
            const target = ast.relationships[id].target;
            return target === fileName || target === `word/${fileName}` || target.endsWith(fileName);
          });
          const ref = headerRefs.find(r => r.rId === matchedRelId);
          const headerType = ref?.type || 'default';
          const headerObj = { type: headerType, xml: doc.documentElement, text };
          if (matchedRelId) {
            ast.headers[matchedRelId] = headerObj;
          }
          ast.headers[fileName] = headerObj;
          if (text && (!ast.metadata.headerText || headerType === 'default')) {
            ast.metadata.headerText = text;
          }
        } catch { /* ignore — best-effort header parse */ }
      } else if (footerMatch) {
        try {
          const xmlStr = entry.getData().toString('utf8');
          assertNoXxe(xmlStr, entry.entryName);
          assertNoXxeAdjacent(xmlStr, entry.entryName);
          const doc = this.parser.parseFromString(xmlStr, 'text/xml');
          const textNodes = doc.getElementsByTagName('w:t');
          let text = '';
          for (let i = 0; i < textNodes.length; i++) {
            text += textNodes[i].textContent || '';
          }
          text = text.trim();
          const fileName = footerMatch[1];
          const matchedRelId = Object.keys(ast.relationships).find(id => {
            const target = ast.relationships[id].target;
            return target === fileName || target === `word/${fileName}` || target.endsWith(fileName);
          });
          const ref = footerRefs.find(r => r.rId === matchedRelId);
          const footerType = ref?.type || 'default';
          const footerObj = { type: footerType, xml: doc.documentElement, text };
          if (matchedRelId) {
            ast.footers[matchedRelId] = footerObj;
          }
          ast.footers[fileName] = footerObj;
          if (text && (!ast.metadata.footerText || footerType === 'default')) {
            ast.metadata.footerText = text;
          }
        } catch { /* ignore — best-effort footer parse */ }
      }
    }

    // ST-09 G1b: expose sectPr parity flags on the AST for renderer parity gating.
    try {
      const sectPrXmlForFlags = sectPr ? this.serializer.serializeToString(sectPr) : '';
      ast.sectPrFlags = {
        titlePg: /<w:titlePg(?:\s[^>]*)?(?:\/>|>)/i.test(sectPrXmlForFlags),
        evenAndOddHeaders: /<w:evenAndOddHeaders(?:\s[^>]*)?(?:\/>|>)/i.test(sectPrXmlForFlags)
      };
    } catch { ast.sectPrFlags = { titlePg: false, evenAndOddHeaders: false }; }

    // Determine the next change ID by finding max w:id in existing ins/del elements
    let maxId = 0;
    const insNodes = Array.from(this.documentDom.getElementsByTagName('w:ins'));
    const delNodes = Array.from(this.documentDom.getElementsByTagName('w:del'));
    [...insNodes, ...delNodes].forEach(el => {
      const id = parseInt(el.getAttribute('w:id') || '0', 10);
      if (id > maxId) maxId = id;
    });
    ast.nextChangeId = maxId + 1;
    
    // One section per section the file defines; a file with no break has just sec_root.
    ast.sections.push(...splitIntoSections(rootNodes, fileSections));
    // BX-03: cache the live AST reference for bare-rezip numbering sync (see
    // field docs). Mutations via NumberingManipulator land in this object.
    //
    // D1R2-2 (second-getAST loss): a previous getAST() reference may hold
    // UNSAVED numbering defs (list.create minted in memory, zip not yet
    // rewritten). Blindly overwriting the cache dropped them: the fresh AST
    // hydrates numbering from the zip (which lacks them), and a later bare
    // rezip() persisted the empty maps. Fix: carry over any abstract/concrete
    // entries the fresh hydration missed. Deletions are NOT resurrected
    // across an explicit prune+rezip (prune deletes the zip part, so a later
    // hydration has nothing to re-add); reload/restore clear the cache so
    // rewound zips still win. Relationships are deliberately NOT merged (rels
    // hydrate fully from the zip; the numbering rel is re-added by
    // syncNumberingToZip on the next save).
    const prevNumberingAst = this.lastAstForNumbering;
    if (prevNumberingAst && prevNumberingAst !== (ast as DocumentAST)) {
      try {
        if (!ast.numbering) ast.numbering = { abstractNum: {}, concreteNum: {} };
        if (!ast.numbering.abstractNum) ast.numbering.abstractNum = {};
        if (!ast.numbering.concreteNum) ast.numbering.concreteNum = {};
        const prevAbs = (prevNumberingAst.numbering?.abstractNum ?? {}) as Record<string, Element>;
        const prevConc = (prevNumberingAst.numbering?.concreteNum ?? {}) as Record<string, Element>;
        for (const [id, el] of Object.entries(prevAbs)) {
          if (!Object.prototype.hasOwnProperty.call(ast.numbering.abstractNum, id)) {
            (ast.numbering.abstractNum as Record<string, Element>)[id] = el as Element;
          }
        }
        for (const [id, el] of Object.entries(prevConc)) {
          if (!Object.prototype.hasOwnProperty.call(ast.numbering.concreteNum, id)) {
            (ast.numbering.concreteNum as Record<string, Element>)[id] = el as Element;
          }
        }
      } catch { /* carry-over stays best-effort — fresh hydration is authoritative */ }
    }
    this.lastAstForNumbering = ast;
    return ast;
  }

  public serializeAST(): Buffer {
    // Because the AST Node's .xml property points to the live DOM elements in this.documentDom,
    // any updates to AST nodes are already reflected in the DOM. 
    // We just need to remove 'data-astra-id' attributes and rezip.
    return this.rezip();
  }

  // --- Query Interface (Task 1.5) ---

  public getNodeFromAST(ast: DocumentAST, id: string): AstNode | undefined {
    return getNodeFromAST(ast, id);
  }

  public searchNodesInAST(ast: DocumentAST, query: string): AstNode[] {
    const results: AstNode[] = [];
    const lowerQuery = query.toLowerCase();
    // Extended projection: legacy body w:t text plus chart title/categories/
    // series, equation m:t runs, footnote/endnote bodies behind their
    // reference ids, and w:delText revision text. Every extension resolves
    // to the host body node, so hit ids stay stable and usable as batch
    // targets. Bad query input throws here exactly as before.
    const decodeEntities = (s: string): string =>
      s
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&amp;/g, '&');
    const noteBodyById = (partPath: string, tag: string): Map<string, string> => {
      const out = new Map<string, string>();
      try {
        const raw = this.getZipAdapter().readTextOrNull(partPath);
        if (!raw) return out;
        assertNoXxe(raw, partPath);
        assertNoXxeAdjacent(raw, partPath);
        const blockRe = new RegExp(`<${tag}\\b[^>]*w:id="([^"]+)"[^>]*>([\\s\\S]*?)<\\/${tag}>`, 'g');
        let block: RegExpExecArray | null;
        while ((block = blockRe.exec(raw)) !== null) {
          const body = block[2];
          const pieces: string[] = [];
          const textRe = /<(?:w:t|w:delText|m:t)\b[^>]*>([^<]*)<\/(?:w:t|w:delText|m:t)>/g;
          let piece: RegExpExecArray | null;
          while ((piece = textRe.exec(body)) !== null) {
            if (piece[1]) pieces.push(decodeEntities(piece[1]));
          }
          if (pieces.length > 0) out.set(block[1], pieces.join(' '));
        }
      } catch (err: unknown) {
        if (err instanceof Error && (/XXE|DTD|E_VALIDATION/.test(err.message))) throw err;
        // Missing or unparsable note part: no note bodies indexed.
      }
      return out;
    };
    const footnoteBodies = noteBodyById('word/footnotes.xml', 'w:footnote');
    const endnoteBodies = noteBodyById('word/endnotes.xml', 'w:endnote');
    for (const sec of ast.sections) {
      for (const node of sec.nodes) {
        const baseText = (node.metadata.text as string).toLowerCase();
        if (baseText.includes(lowerQuery)) {
          results.push(node);
          continue;
        }
        const extra: string[] = [];
        try {
          const md = node.metadata as unknown as Record<string, unknown>;
          const chartData = md['chartData'] as Record<string, unknown> | undefined;
          const pushStr = (v: unknown): void => {
            if (typeof v === 'string' && v.length > 0) extra.push(v);
          };
          const pushStrOrNum = (v: unknown): void => {
            if (typeof v === 'string' && v.length > 0) extra.push(v);
            else if (typeof v === 'number' && Number.isFinite(v)) extra.push(String(v));
          };
          pushStr(chartData?.['title']);
          pushStr(md['title']);
          const rawCategories = (chartData && Array.isArray(chartData['categories']) &&
            (chartData['categories'] as unknown[]).length > 0
            ? chartData['categories']
            : md['categories']) as unknown;
          if (Array.isArray(rawCategories)) {
            for (const c of rawCategories) pushStrOrNum(c);
          }
          const rawSeries = (chartData && Array.isArray(chartData['series']) &&
            (chartData['series'] as unknown[]).length > 0
            ? chartData['series']
            : md['series']) as unknown;
          if (Array.isArray(rawSeries)) {
            for (const s of rawSeries) {
              if (typeof s === 'string') { pushStr(s); continue; }
              if (s && typeof s === 'object') {
                const rec = s as Record<string, unknown>;
                pushStr(rec['name']);
                if (Array.isArray(rec['values'])) {
                  for (const v of rec['values'] as unknown[]) pushStrOrNum(v);
                }
              }
            }
          }
          const el = node.xml as unknown as Element | undefined;
          if (el && typeof (el as Element).getElementsByTagName === 'function') {
            const domEl = el as Element;
            const mathRuns = domEl.getElementsByTagName('m:t');
            for (let i = 0; i < mathRuns.length; i++) {
              const t = (mathRuns[i] as unknown as Element).textContent;
              if (typeof t === 'string' && t.length > 0) extra.push(t);
            }
            const delRuns = domEl.getElementsByTagName('w:delText');
            for (let i = 0; i < delRuns.length; i++) {
              const t = (delRuns[i] as unknown as Element).textContent;
              if (typeof t === 'string' && t.length > 0) extra.push(t);
            }
            const fnRefs = domEl.getElementsByTagName('w:footnoteReference');
            for (let i = 0; i < fnRefs.length; i++) {
              const refId = (fnRefs[i] as unknown as Element).getAttribute?.('w:id');
              if (typeof refId === 'string' && footnoteBodies.has(refId)) {
                extra.push(footnoteBodies.get(refId) as string);
              }
            }
            const enRefs = domEl.getElementsByTagName('w:endnoteReference');
            for (let i = 0; i < enRefs.length; i++) {
              const refId = (enRefs[i] as unknown as Element).getAttribute?.('w:id');
              if (typeof refId === 'string' && endnoteBodies.has(refId)) {
                extra.push(endnoteBodies.get(refId) as string);
              }
            }
          }
        } catch {
          // Per-node extras stay best-effort; the legacy base match stands.
        }
        if (extra.length > 0 && extra.join('\n').toLowerCase().includes(lowerQuery)) {
          results.push(node);
        }
      }
    }
    return results;
  }

  public async updateNodeInAST(ast: DocumentAST, nodeId: string, newXml: string) {
    // Fallback to legacy updateNode to safely parse and inject XML
    this.updateNode(nodeId, newXml);
    
    // Update AST specific fields (Version & Checksum tracking - Task 1.3)
    let node = this.getNodeFromAST(ast, nodeId);
    const updatedElement = this.nodeMap.get(nodeId);

    if (!node && updatedElement) {
      // If node was inserted into DOM/nodeMap after AST snapshot, register into AST
      const textNodes = updatedElement.getElementsByTagName('w:t');
      let text = '';
      for (let i = 0; i < textNodes.length; i++) {
        text += textNodes[i].textContent || '';
      }
      node = {
        id: nodeId,
        type: 'paragraph',
        xml: updatedElement,
        metadata: {
          text,
          hasCitation: false,
          hasComment: false,
          hasBookmark: false,
          hasTrackedChanges: false,
          hasEquation: false,
          hasField: false,
          hasImage: false,
          hasHyperlink: false
        },
        version: 2,
        checksum: await calculateNodeChecksum(newXml)
      };
      addNodeInDocumentOrder(ast.sections, node);
    } else if (node && updatedElement) {
      node.xml = updatedElement;
      node.version += 1;
      node.checksum = await calculateNodeChecksum(newXml);
      
      // Re-extract text
      const textNodes = updatedElement.getElementsByTagName('w:t');
      let text = '';
      for (let i = 0; i < textNodes.length; i++) {
        text += textNodes[i].textContent || '';
      }
      node.metadata.text = text;
    }
  }

  // --- End AST layer ---

  public readNode(nodeId: string): string {
    // Security: Validate nodeId format to prevent traversal-like attempts, even though it's in-memory
    if (!/^[a-zA-Z0-9_:]+$/.test(nodeId)) {
      throw new Error(`Invalid Node ID format: ${nodeId}`);
    }

    const node = this.nodeMap.get(nodeId);
    if (!node) throw new Error(`Node ${nodeId} not found`);
    return this.serializer.serializeToString(node);
  }

  public updateNode(nodeId: string, newXml: string) {
    if (!/^[a-zA-Z0-9_:]+$/.test(nodeId)) {
      throw new Error(`Invalid Node ID format: ${nodeId}`);
    }

    const oldNode = this.nodeMap.get(nodeId);
    if (!oldNode) throw new Error(`Node ${nodeId} not found`);

    assertNoXxe(newXml, `node ${nodeId}`);

    // Parse the incoming XML string to ensure it is valid XML.
    // If it contains unbalanced tags or XML injection, DOMParser will throw or create a parsererror.
    const newDoc = this.parser.parseFromString(newXml, "text/xml");
    const parserError = newDoc.getElementsByTagName('parsererror');
    if (parserError.length > 0) {
      throw new Error("XML Injection or Invalid XML detected in input.");
    }

    const newNode = newDoc.documentElement;
    if (newNode.tagName !== oldNode.tagName) {
      throw new Error(`Tag mismatch: expected ${oldNode.tagName}, got ${newNode.tagName}`);
    }

    // Retain the assigned ID
    newNode.setAttribute('data-astra-id', nodeId);

    // A-4: preserve paragraph identity across wholesale rewrites. If the
    // replacement w:p carries no valid paraId of its own, it inherits the
    // old node's (resume-by-id keeps working after edits, and rollback
    // snapshots round-trip byte-identically).
    if (newNode.tagName === 'w:p' && !isValidParaId(newNode.getAttribute('w14:paraId'))) {
      const inherited = oldNode.getAttribute('w14:paraId');
      if (isValidParaId(inherited)) {
        newNode.setAttributeNS(W14_NS, 'w14:paraId', inherited!);
      }
    }

    // Replace in DOM
    oldNode.parentNode?.replaceChild(newNode, oldNode);

    // Update map
    this.nodeMap.set(nodeId, newNode);
    this._bumpNodePagesRevision();
  }

  public insertRootNode(rawXml: string, anchorId?: string, position: 'before' | 'after' | 'append' = 'append', customId?: string): string {
    assertNoXxe(rawXml, 'inserted node XML');

    const newDoc = this.parser.parseFromString(rawXml, "text/xml");
    const parserError = newDoc.getElementsByTagName('parsererror');
    if (parserError.length > 0) {
      throw new Error("XML Injection or Invalid XML detected in input.");
    }

    const newNode = newDoc.documentElement;
    const body = this.documentDom.getElementsByTagName('w:body')[0];
    if (!body) throw new Error("Invalid DOCX: missing w:body element");

    const newId = customId || `node_${parserNowMs()}_${Math.floor(parserRandomUnit() * 1000)}`;
    newNode.setAttribute('data-astra-id', newId);

    // A-4: inserted w:p roots get a minted w14:paraId now so the paragraph
    // carries a stable reload-surviving identity (session identity stays newId).
    if (newNode.tagName === 'w:p' && !isValidParaId(newNode.getAttribute('w14:paraId'))) {
      const seed = customId || `${anchorId || 'root'}:${position || 'after'}:${newNode.textContent || ''}`;
      newNode.setAttributeNS(W14_NS, 'w14:paraId', this.generateUniqueParaId(seed));
      const root = this.documentDom.documentElement;
      if (root && !root.hasAttribute('xmlns:w14')) {
        root.setAttributeNS(XMLNS_NS, 'xmlns:w14', W14_NS);
      }
    }

    if (anchorId && this.nodeMap.has(anchorId)) {
      const anchorNode = this.nodeMap.get(anchorId)!;
      // BX-11: resolve the TRUE parent (the anchor's own parent) instead of
      // assuming w:body. Anchors wrapped in w:sdt > w:sdtContent previously
      // had the new node ripped out to body level (mis-insert).
      let parent: Element = body;
      try {
        const trueParent = anchorNode.parentNode as Element | null;
        if (trueParent && trueParent.nodeType === 1) {
          parent = trueParent;
        }
      } catch { /* fall back to body */ }
      if (position === 'before') {
        parent.insertBefore(newNode, anchorNode);
      } else {
        parent.insertBefore(newNode, anchorNode.nextSibling);
      }
    } else {
      let sectPr: Element | undefined;
      for (let i = 0; i < body.childNodes.length; i++) {
        const child = body.childNodes[i] as Element;
        if (child.nodeType === 1 && child.tagName === 'w:sectPr') {
          sectPr = child;
          break;
        }
      }
      if (sectPr) {
        body.insertBefore(newNode, sectPr);
      } else {
        body.appendChild(newNode);
      }
    }

    this.nodeMap.set(newId, newNode);
    // BX-11: recursive nodeMap registration mirroring the read path
    // (assignStableIds/processContainer): inner w:p/w:tbl children of an
    // inserted w:sdt subtree become addressable instead of invisible to
    // TOON/ops/summaries (which iterate nodeMap).
    this.registerInsertedSubtree(newNode, newId);
    this._bumpNodePagesRevision();
    return newId;
  }

  /**
   * BX-11: registers inner addressable children of an inserted subtree.
   * Mirrors the read-path recursion (w:sdt > w:sdtContent > w:p/w:tbl).
   * Plain w:p/w:tbl roots need nothing beyond the top-level registration.
   */
  private registerInsertedSubtree(root: Element, rootId: string): void {
    if (!root || root.tagName !== 'w:sdt') return;
    let childSeq = 0;
    const grantId = (prefix: string): string => {
      let candidate: string;
      do {
        childSeq++;
        candidate = `${rootId}__${prefix}${childSeq}`;
      } while (this.nodeMap.has(candidate));
      return candidate;
    };
    const walk = (container: Element): void => {
      for (let i = 0; i < container.childNodes.length; i++) {
        const node = container.childNodes[i] as Element;
        if (!node || node.nodeType !== 1) continue;
        if (node.tagName === 'w:p' || node.tagName === 'w:tbl') {
          const current = typeof node.hasAttribute === 'function' && node.hasAttribute('data-astra-id')
            ? node.getAttribute('data-astra-id')
            : null;
          if (!current || this.nodeMap.has(current)) {
            node.setAttribute('data-astra-id', grantId(node.tagName === 'w:p' ? 'p' : 't'));
          }
          this.nodeMap.set(node.getAttribute('data-astra-id')!, node);
        } else if (node.tagName === 'w:sdtContent' || node.tagName === 'w:sdt') {
          walk(node);
        }
      }
    };
    walk(root);
  }

  public replaceRootNode(nodeId: string, newXml: string): void {
    this.updateNode(nodeId, newXml);
  }

  public getNode(nodeId: string): Element | undefined {
    return this.nodeMap.get(nodeId);
  }

  /**
   * Identity recursion: stable ids minted for table-cell paragraphs
   * (w:tr/w:tc content). Resolvable via getNode/readNode/updateNode;
   * excluded from root-level summaries and AST node lists.
   */
  public getTableCellParagraphIds(): string[] {
    const out: string[] = [];
    for (const id of this.cellScopedIds) {
      if (this.nodeMap.get(id)?.tagName === 'w:p') out.push(id);
    }
    return out;
  }

  /** True when id belongs to table-cell content rather than a root node. */
  public isTableCellScopedId(id: string): boolean {
    return this.cellScopedIds.has(id);
  }

  /** All w:styleId values declared in word/styles.xml (for style-name resolution). */
  public getStyleIds(): string[] {
    if (!this.stylesDom) return [];
    return Array.from(this.stylesDom.getElementsByTagName('w:style'))
      .map(s => s.getAttribute('w:styleId'))
      .filter((v): v is string => !!v);
  }

  public deleteRootNode(nodeId: string): void {
    const node = this.nodeMap.get(nodeId);
    if (!node) throw new Error(`Node ${nodeId} not found`);
    node.parentNode?.removeChild(node);
    this.nodeMap.delete(nodeId);
    this._bumpNodePagesRevision();
  }

  /**
   * Drops a node id from the stable-id map after an external (AST-driven)
   * DOM removal, so stale lookups fail fast instead of returning detached XML.
   */
  public removeFromNodeMap(nodeId: string): void {
    this.nodeMap.delete(nodeId);
    this._bumpNodePagesRevision();
  }

  /**
   * Gives blocks new ids: `renames` maps a block's current id to the id it
   * should have. Used after the package was rewound and re-read, to give the
   * blocks back the ids callers were handed before. All renames apply
   * together, so ids may swap. Changes nothing and returns false when a
   * current id is unknown, two blocks would get the same id, or a new id
   * belongs to a block that is not being renamed.
   */
  public reassignNodeIds(renames: ReadonlyMap<string, string>): boolean {
    const wanted = new Set<string>();
    const elements = new Map<string, Element>();
    for (const [from, to] of renames) {
      const element = this.nodeMap.get(from);
      if (!element || wanted.has(to)) return false;
      if (this.nodeMap.has(to) && !renames.has(to)) return false;
      wanted.add(to);
      elements.set(from, element);
    }
    for (const from of elements.keys()) this.nodeMap.delete(from);
    for (const [from, element] of elements) {
      const to = renames.get(from) as string;
      element.setAttribute('data-astra-id', to);
      this.nodeMap.set(to, element);
    }
    if (elements.size > 0) this._bumpNodePagesRevision();
    return true;
  }

  /**
   * The <w:sectPr> of every section the file defines, in order: one per section
   * break paragraph, then the one at the end of the body. Properties kept inside
   * a tracked section change and sectPr elements in table cells are not sections
   * and are not listed.
   */
  public getSectPrElements(): Element[] {
    const body = this.documentDom.getElementsByTagName('w:body')[0];
    return listFileSections(body).flatMap((fileSection) => (fileSection.sectPr ? [fileSection.sectPr] : []));
  }

  /** The <w:sectPr> that ends the body (the last section's properties), if the file has one. */
  public getBodySectPr(): Element | null {
    return bodySectPr(this.documentDom.getElementsByTagName('w:body')[0]);
  }

  public readDocumentXml(): string {
    const cleanDom = this.documentDom.cloneNode(true) as Document;
    // Recursively strip temporary 'data-astra-id' attributes from all elements
    const stripAstraId = (el: Element) => {
      el.removeAttribute('data-astra-id');
      for (let i = 0; i < el.childNodes.length; i++) {
        const child = el.childNodes[i];
        if (child.nodeType === 1) stripAstraId(child as Element);
      }
    };
    if (cleanDom.documentElement) {
      stripAstraId(cleanDom.documentElement);
    }
    // Canonicalize DOM schema order (rPr, pPr, tblPr, tblBorders, tcPr, tblGrid, w14:paraId)
    canonicalizeDocumentDom(cleanDom);
    return this.serializer.serializeToString(cleanDom);
  }

  /**
   * Live duplicate-paraId scan (load-time AND edit-introduced duplicates —
   * no load-time state to go stale). Powers the save-guard refusal and the
   * `validateOOXML` ok:false mapping.
   */
  public getDuplicateParaIds(): DuplicateParaId[] {
    try {
      return findDuplicateParaIds(this.documentDom);
    } catch {
      return [];
    }
  }

  /**
   * Fail-closed duplicate-paraId refusal for the save path. A repeated
   * native id ships downstream merge/migration collisions no tool would
   * name, so byte-production refuses while open/inspect keep working
   * (Word tolerates duplicates on read). The corpus carries zero
   * duplicates, so only crafted/anomalous inputs newly refuse.
   */
  private assertNoDuplicateParaIds(): void {
    const dups = this.getDuplicateParaIds();
    if (dups.length === 0) return;
    const shown = dups.slice(0, 5).map((d) => {
      const at = d.xPaths.slice(0, 3).join(', ');
      const more = d.xPaths.length > 3 ? ` (+${d.xPaths.length - 3} more)` : '';
      return `duplicate w14:paraId "${d.paraId}" on ${d.count} paragraphs (at ${at}${more})`;
    });
    const tail = dups.length > 5 ? `; +${dups.length - 5} more duplicate ids` : '';
    const first = dups[0];
    const err = toError(
      'E_VALIDATION',
      `[OOXML Duplicate Guard] Export blocked: ${shown.join('; ')}${tail}`,
      `delete the repeat paragraph (session id ${first.paraId}_dup — see inspect_document) and reinsert its text for a fresh id, then retry save_document`
    );
    throw Object.assign(err, { astraGuard: 'duplicate-paraid' as SaveGuardKind, duplicates: dups });
  }

  /**
   * Dry-run preview of pending allowlist normalizations for the live
   * document DOM (clone-based, never mutates). Powers open/inspect
   * surfacing so operators see the healing a save would apply before any
   * bytes change. Empty when the document is already canonical.
   */
  public getPendingNormalizations(): string[] {
    try {
      return previewDocumentDom(this.documentDom);
    } catch {
      return [];
    }
  }

  public validateZipIntegrity(): { isValid: boolean; errors: string[]; warnings: string[] } {
    const errors: string[] = [];
    const warnings: string[] = [];
    const entries = this.zip.getEntries();

    for (const entry of entries) {
      const name = entry.entryName;
      // R-10: rogue entries are purged at save ingress, never validated or
      // read here (their bytes must not fail or pass the integrity gate).
      if (isRogueZipEntry(name) || isBlockedExecutableEntry(name)) continue;
      if (name.endsWith('.xml') || name.endsWith('.rels')) {
        const content = readPartData(entry, name).toString('utf8');
        try {
          let hasError = false;
          let errorMessage = '';
          const strictParser = new DOMParser({
            errorHandler: {
              // Non-blocking XML warnings ride the typed `warnings` channel
              // (same text as the removed console.warn); errors stay in `errors`.
              warning: (msg) => { warnings.push(`[OOXML Integrity Guard] XML warning (non-blocking): ${msg}`); },
              error: (msg) => { hasError = true; errorMessage += msg + '; '; },
              fatalError: (msg) => { hasError = true; errorMessage += msg + '; '; }
            }
          });
          const doc = strictParser.parseFromString(content, 'text/xml');
          const parserError = doc.getElementsByTagName('parsererror');
          if (hasError || parserError.length > 0) {
            errors.push(`XML Well-Formedness Error in ${name}: ${errorMessage || parserError[0]?.textContent || 'Invalid XML'}`);
          }
        } catch (err: unknown) {
          let msg: string;
          if (err instanceof Error) msg = err.message !== '' ? err.message : String(err);
          else if (typeof err === 'object' && err !== null && 'message' in err && typeof (err as { message?: unknown }).message === 'string' && ((err as { message: string }).message !== '')) msg = (err as { message: string }).message;
          else msg = String(err);
          errors.push(`XML Parse Exception in ${name}: ${msg}`);
        }
      }
    }

    return { isValid: errors.length === 0, errors, warnings };
  }

  public rezip(astOrOptions?: DocumentAST | { skipValidation?: boolean; validate?: boolean; reproducible?: boolean; noReproducible?: boolean; ast?: DocumentAST }, options?: { skipValidation?: boolean; validate?: boolean; reproducible?: boolean; noReproducible?: boolean }): Buffer {
    // Backward-compatible arg resolution: supports rezip(), rezip({ skipValidation }),
    // rezip(ast), and rezip(ast, { skipValidation }). Deterministic by default;
    // explicit reproducible:false or noReproducible:true opts out to wall-clock.
    // Fail-closed by default: the duplicate-paraId guard below refuses output
    // with repeated ids. Whether the XML is schema-valid is the async validator's
    // job (`validateDocxBuffer`), not this synchronous method's. Explicit opt-outs
    // (performance-only escape hatches for callers that validate elsewhere):
    // `rezip({ validate: false })` skips the duplicate-paraId guard (integrity,
    // XXE, and macro gates still run); `skipValidation: true` skips all
    // validation (legacy broad opt-out).
    let ast: DocumentAST | undefined;
    let skipValidation = false;
    let explicitRepro: boolean | undefined;
    let explicitValidate: boolean | undefined;
    const readExplicit = (v: unknown): boolean | undefined =>
      v === true ? true : v === false ? false : undefined;
    if (astOrOptions && typeof astOrOptions === 'object' && 'sections' in astOrOptions && 'relationships' in astOrOptions) {
      ast = astOrOptions as DocumentAST;
      skipValidation = !!options?.skipValidation;
      explicitValidate = readExplicit((options as { validate?: unknown } | undefined)?.validate);
      explicitRepro = readExplicit(options?.reproducible);
      if (explicitRepro === undefined && (options as { noReproducible?: unknown } | undefined)?.noReproducible === true) explicitRepro = false;
    } else if (astOrOptions && typeof astOrOptions === 'object' && 'ast' in astOrOptions && (astOrOptions as { ast?: unknown }).ast) {
      ast = (astOrOptions as { ast: DocumentAST }).ast as DocumentAST;
      skipValidation = !!((astOrOptions as { skipValidation?: unknown }).skipValidation || options?.skipValidation);
      explicitValidate = readExplicit((astOrOptions as { validate?: unknown }).validate);
      if (explicitValidate === undefined) explicitValidate = readExplicit((options as { validate?: unknown } | undefined)?.validate);
      explicitRepro = readExplicit((astOrOptions as { reproducible?: unknown }).reproducible);
      if (explicitRepro === undefined) explicitRepro = readExplicit(options?.reproducible);
      if (explicitRepro === undefined && ((astOrOptions as { noReproducible?: unknown }).noReproducible === true || (options as { noReproducible?: unknown } | undefined)?.noReproducible === true)) explicitRepro = false;
    } else {
      skipValidation = !!((astOrOptions as { skipValidation?: boolean } | undefined)?.skipValidation || options?.skipValidation);
      explicitValidate = readExplicit((astOrOptions as { validate?: unknown } | undefined)?.validate);
      if (explicitValidate === undefined) explicitValidate = readExplicit((options as { validate?: unknown } | undefined)?.validate);
      explicitRepro = readExplicit((astOrOptions as { reproducible?: unknown } | undefined)?.reproducible);
      if (explicitRepro === undefined) explicitRepro = readExplicit(options?.reproducible);
      if (explicitRepro === undefined && ((astOrOptions as { noReproducible?: unknown } | undefined)?.noReproducible === true || (options as { noReproducible?: unknown } | undefined)?.noReproducible === true)) explicitRepro = false;
    }
    const reproducible = resolveReproducibleMode(explicitRepro);
    const skipContentGuards = skipValidation || explicitValidate === false;

    // Duplicate-paraId gate first: the live-DOM scan mutates zero state, so
    // refusal leaves the session untouched (verified: open/inspect/retry
    // keep working after a refused save).
    if (!skipContentGuards) this.assertNoDuplicateParaIds();

    // R-10: save ingress drops rogue entries first, so the XXE sweep,
    // integrity guards, and output bytes below never see hostile parts.
    this.purgeRogueZipEntries();

    // Macro/active-content gate at save ingress: refuse-default throws
    // E_VALIDATION naming the parts (zero bytes change); allow preserves
    // them byte-identical; strip removes them before the sweep below.
    this.enforceMacroGate('save');

    // Fail-closed XXE sweep over every stored XML/rels part BEFORE integrity
    // validation or canonicalization: a DTD smuggled into an auxiliary part
    // (theme, chart, header, ...) can neither survive export nor be
    // normalized into something parseable downstream. The adjacent sweep
    // (XInclude/XSLT/stylesheet-PI) rides along: well-formed OOXML never
    // carries these shapes, so legit packages pass untouched.
    for (const entry of this.zip.getEntries()) {
      const name = entry.entryName;
      if (isRogueZipEntry(name) || isBlockedExecutableEntry(name)) continue;
      if (name.endsWith('.xml') || name.endsWith('.rels')) {
        const text = readPartData(entry, name).toString('utf8');
        assertNoXxe(text, name);
        assertNoXxeAdjacent(text, name);
      }
    }

    // 1. Automated Pre-Save XML Integrity Validation Guard on existing entries
    if (!skipValidation) {
      const validation = this.validateZipIntegrity();
      if (!validation.isValid) {
        throw toError('E_VALIDATION', `[OOXML Integrity Guard] Export blocked due to XML corruption:\n${validation.errors.join('\n')}`, 'repair the part XML or re-save the document in Word before retrying');
      }
    }

    // Allowlist normalizer (ooxml/normalize.ts): verified renames, whitelisted
    // default infills, and schema-sequence sorts run on the live DOMs BEFORE
    // canonicalization and the strict gate. Syntax errors already failed
    // closed in the integrity guard above. No-op passes mutate nothing, so
    // clean documents serialize byte-identical with an empty log.
    const normalized: string[] = [];
    normalized.push(...normalizeDocumentDom(this.documentDom));
    if (this.stylesDom) {
      for (const line of normalizeDocumentDom(this.stylesDom)) normalized.push(`word/styles.xml: ${line}`);
    }

    // 2. Remove temporary 'data-astra-id' attributes and canonicalize XML according to ECMA-376
    const finalXmlStr = this.readDocumentXml();
    this.zip.updateFile('word/document.xml', Buffer.from(finalXmlStr, 'utf8'));

    // 3. Canonicalize styles DOM if present
    if (this.stylesDom) {
      canonicalizeDocumentDom(this.stylesDom);
      const stylesXml = this.serializer.serializeToString(this.stylesDom);
      this.zip.updateFile('word/styles.xml', Buffer.from(stylesXml, 'utf8'));
    }

    // 4. Canonicalize any auxiliary headers/footers in the package
    const entries = this.zip.getEntries();
    for (const entry of entries) {
      const name = entry.entryName;
      if (name.startsWith('word/') && (name.includes('header') || name.includes('footer')) && name.endsWith('.xml')) {
        const raw = entry.getData().toString('utf8');
        const part = normalizePartXml(name, raw);
        for (const line of part.normalized) normalized.push(`${name}: ${line}`);
        const canon = canonicalizePartXml(part.xml);
        this.zip.updateFile(name, Buffer.from(canon, 'utf8'));
      }
    }
    this.lastNormalizations = normalized;

    // 4b. PARSER-01: synchronize AST media + relationships into the package so
    // images added via RelationshipsManipulator.addMedia / executeImage survive
    // export even when the caller mutated only the AST (no direct zip write).
    if (ast) {
      this.syncMediaFilesToZip(ast);
      this.syncNumberingToZip(ast);
      this.syncRelationshipsToZip(ast);
      this.syncContentTypesToZip(ast);
      this.syncCorePropertiesToZip(ast);
      // Keep the cache pointing at the freshest explicit AST so a later bare
      // rezip() (prod save) still sees its numbering defs.
      this.lastAstForNumbering = ast;
    } else if (this.lastAstForNumbering) {
      // BX-03 route-around: production saves call bare rezip() with no AST, so
      // in-memory list defs would be dropped. Replay NUMBERING ONLY from the
      // cached getAST() reference (mutated in place by list.create). The aux
      // rel/content-type sync is scoped to the numbering part alone — full
      // media/rel merges are deliberately skipped to avoid resurrecting
      // deleted parts from a stale snapshot.
      try {
        this.syncNumberingToZip(this.lastAstForNumbering);
        this.syncNumberingRelAndContentTypeBare(this.lastAstForNumbering);
      } catch { /* numbering sync stays best-effort — document.xml already written */ }
    }

    // R-10: AST syncs above can stage new parts (e.g. media); purge again
    // so a blocked name staged mid-save cannot reach the output bytes.
    this.purgeRogueZipEntries();

    // 5. Automated Post-Assembly XML Integrity Validation Guard on all package entries
    if (!skipValidation) {
      const validation = this.validateZipIntegrity();
      if (!validation.isValid) {
        throw toError('E_VALIDATION', `[OOXML Integrity Guard] Export blocked due to XML corruption:\n${validation.errors.join('\n')}`, 'repair the part XML or re-save the document in Word before retrying');
      }
    }
    
    // Deterministic-by-default: pin timestamps + docProps to one epoch;
    // opt-out (explicit false / env) restores wall-clock: zip entry times
    // stamped to now, docProps preserved, deterministic revision/comment
    // stamps plus pinned cover dates rewritten to now. Entry order and
    // sequential IDs stay deterministic either way.
    if (reproducible) {
      const stamp = resolveDeterministicZipDate();
      clampCorePropertiesDates(this.zip, stamp);
      clampZipEntryTimes(this.zip, stamp);
    } else {
      try {
        clampZipEntryTimes(this.zip, new Date());
      } catch { /* best-effort — never block save */ }
      this.unpinDeterministicRevisionDates();
    }
    return this.zip.toBuffer();
  }

  /**
   * Save-layer unpin: rewrite already-minted deterministic w:date stamps to
   * wall-clock so an explicit reproducible:false at rezip fully unpins even
   * when the edits were minted under the deterministic default. Only stamps
   * equal to a deterministic value (1980 fallback or current
   * SOURCE_DATE_EPOCH, both revision + comment shapes) are rewritten;
   * explicit user dates and historical stamps are preserved verbatim.
   * Pinned cover dates (1980 fallback or current epoch locale text) are
   * rewritten the same way via the cover deterministic path. Single now for
   * the whole save (revision/cover dates use now, per guide).
   */
  private unpinDeterministicRevisionDates(): void {
    let nowIso: string;
    let nowCover: string;
    try {
      const now = new Date();
      nowIso = now.toISOString();
      nowCover = now.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });
    } catch {
      return;
    }
    const nowCommentIso = nowIso.replace(/\.\d{3}Z$/, 'Z');
    let pinIsos: Set<string>;
    try {
      const current = resolveDeterministicZipDate().toISOString();
      const fallback = new Date(Date.UTC(1980, 0, 1, 0, 0, 0)).toISOString();
      pinIsos = new Set<string>([
        current,
        current.replace(/\.\d{3}Z$/, 'Z'),
        fallback,
        fallback.replace(/\.\d{3}Z$/, 'Z'),
      ]);
    } catch {
      return;
    }
    try {
      for (const tag of ['w:ins', 'w:del']) {
        const els = this.documentDom.getElementsByTagName(tag);
        for (let i = 0; i < els.length; i++) {
          const el = els[i] as Element;
          const v = el.getAttribute?.('w:date');
          if (v && pinIsos.has(v)) el.setAttribute('w:date', nowIso);
        }
      }
      try {
        this.zip.updateFile('word/document.xml', Buffer.from(this.readDocumentXml(), 'utf8'));
      } catch { /* best-effort — never block save */ }
    } catch { /* best-effort — never block save */ }
    try {
      const adapter = this.getZipAdapter();
      for (const part of ['word/comments.xml', 'word/commentsExtended.xml']) {
        let xml: string | null = null;
        try {
          xml = adapter.readTextOrNull(part);
        } catch {
          continue;
        }
        if (!xml) continue;
        let out = xml;
        for (const pin of pinIsos) {
          if (!pin) continue;
          const needle = `w:date="${pin}"`;
          if (out.includes(needle)) {
            const isCommentShape = /\.\d{3}Z$/.test(pin) === false;
            out = out.split(needle).join(`w:date="${isCommentShape ? nowCommentIso : nowIso}"`);
          }
        }
        if (out !== xml) {
          try {
            adapter.writeText(part, out);
          } catch { /* best-effort */ }
        }
      }
    } catch { /* best-effort */ }
    try {
      const fmt = { month: 'long', day: 'numeric', year: 'numeric' } as const;
      const pins = new Set<string>();
      try {
        pins.add(resolveDeterministicZipDate().toLocaleDateString('en-US', fmt));
        pins.add(new Date(Date.UTC(1980, 0, 1, 0, 0, 0)).toLocaleDateString('en-US', fmt));
      } catch { /* best-effort */ }
      if (pins.size > 0 && nowCover && ![...pins].includes(nowCover)) {
        let touched = false;
        const texts = this.documentDom.getElementsByTagName('w:t');
        for (let i = 0; i < texts.length; i++) {
          const t = texts[i] as Element;
          const cur = t.textContent || '';
          if (!cur) continue;
          for (const pin of pins) {
            if (pin && cur.includes(pin)) {
              let coverScoped = false;
              try {
                let anc: Node | null = t.parentNode;
                while (anc) {
                  if (anc.nodeType === 1 && (anc as Element).tagName === 'w:p') {
                    const v = (anc as Element).getAttribute?.('data-astra-id') || '';
                    if (typeof v === 'string' && v.includes('cover_')) coverScoped = true;
                    break;
                  }
                  anc = anc.parentNode;
                }
              } catch { /* best-effort */ }
              if (!coverScoped) continue;
              t.textContent = cur.split(pin).join(nowCover);
              touched = true;
              break;
            }
          }
        }
        if (touched) {
          try {
            this.zip.updateFile('word/document.xml', Buffer.from(this.readDocumentXml(), 'utf8'));
          } catch { /* best-effort — never block save */ }
        }
      }
    } catch { /* best-effort — never block save */ }
  }

  /**
   * E2 (F-2/F-3 companion): reconcile core-properties on the explicit-AST
   * rezip path. Create-if-absent when the part is missing (headless engine
   * use); otherwise merge any in-memory AST core fields that DIFFER from the
   * part (unknown siblings preserved verbatim). Never lets a stale part
   * silently clobber live AST edits. Loud (fail-closed): malformed/XXE parts
   * and adapter write failures throw instead of being swallowed.
   */
  private syncCorePropertiesToZip(ast: DocumentAST): void {
    const adapter = this.getZipAdapter();
    const existing = adapter.readTextOrNull('docProps/core.xml');
    if (!existing) {
      const meta = ast?.metadata as Record<string, unknown> | undefined;
      if (!meta) return;
      const hasAny = ['title', 'creator', 'subject', 'description', 'keywords',
        'category', 'lastModifiedBy', 'revision', 'created', 'modified']
        .some((k) => meta[k] !== undefined);
      if (!hasAny) return;
      new CorePropertiesManipulator(ast, adapter).update({});
      return;
    }
    // Reconcile (a): diff AST vs part; merge differing in-memory fields in.
    const partMeta = parseCorePropertiesXml(existing);
    const astMeta = (ast?.metadata ?? {}) as Record<string, unknown>;
    const sameValue = (a: unknown, b: unknown): boolean => {
      if (a === undefined || b === undefined) return a === b;
      if (a instanceof Date || b instanceof Date) {
        const ta = a instanceof Date ? a.getTime() : new Date(String(a)).getTime();
        const tb = b instanceof Date ? b.getTime() : new Date(String(b)).getTime();
        if (Number.isNaN(ta) || Number.isNaN(tb)) return String(a) === String(b);
        return ta === tb;
      }
      return a === b;
    };
    const patch: Record<string, unknown> = {};
    for (const k of ['title', 'creator', 'subject', 'description', 'keywords',
      'category', 'lastModifiedBy', 'revision', 'created', 'modified'] as const) {
      const v = astMeta[k];
      if (v === undefined) continue;
      if (!sameValue(v, (partMeta as Record<string, unknown>)[k])) {
        (patch as Record<string, unknown>)[k] = v;
      }
    }
    if (Object.keys(patch).length === 0) return;
    new CorePropertiesManipulator(ast, adapter).update(
      patch as unknown as Parameters<CorePropertiesManipulator['update']>[0]
    );
  }

  /**
   * PARSER-01: write every AST media buffer to word/media/*. Uses getEntry to
   * decide between updateFile (overwrite) and addFile (new part). Never deletes
   * existing media, so a stale/empty ast.mediaFiles cannot drop original images.
   */
  private syncMediaFilesToZip(ast: DocumentAST): void {
    const mediaFiles = ast.mediaFiles;
    if (!mediaFiles) return;
    this.invalidateMediaCache();
    for (const [fileName, data] of Object.entries(mediaFiles)) {
      if (!fileName || !data) continue;
      const buf = Buffer.isBuffer(data) ? data : Buffer.from(data as unknown as Uint8Array);
      let zipPath: string;
      if (fileName.startsWith('word/media/')) zipPath = fileName;
      else if (fileName.startsWith('media/')) zipPath = `word/${fileName}`;
      else if (fileName.startsWith('word/')) zipPath = fileName;
      else zipPath = `word/media/${fileName}`;
      // Path-traversal guard (mirrors sanitizePath without throwing mid-export).
      const normalized = zipPath.replace(/\\/g, '/');
      if (normalized.includes('..') || normalized.startsWith('/') || /^[a-zA-Z]:/.test(normalized)) continue;
      if (this.zip.getEntry(zipPath)) {
        this.zip.updateFile(zipPath, buf);
      } else {
        this.zip.addFile(zipPath, buf);
      }
    }
  }

  /**
   * BX-03: writes ast.numbering back to word/numbering.xml so list.create
   * definitions survive save/reload (previously minted in-memory only and
   * silently dropped by rezip). No-op when the AST holds no numbering
   * definitions — documents without lists gain no spurious part. Ensures the
   * numbering relationship + content-type override alongside the part
   * (syncRelationshipsToZip/syncContentTypesToZip pick the rel up from the
   * AST map on the same pass).
   *
   * D1R2-1 (orphan prune, parser-side; numbering.ts untouched — owned by
   * another batch): a payload whose concrete nums reference NO existing
   * abstract (e.g. concreteNum deleted, abstract remains) is an orphan part.
   * Null/orphan now means PRUNE, not "leave the zip alone": a previous save
   * may have written numbering.xml + its rel + its content-type override,
   * which would otherwise linger as dangling references. pruneNumberingFromZip
   * removes all three so the package stays clean.
   */
  private syncNumberingToZip(ast: DocumentAST): void {
    let payload: string | null = null;
    try {
      payload = new NumberingManipulator(ast).serializeNumbering();
    } catch {
      payload = null;
    }
    // D1R2-1 orphan gate: non-empty maps but no live concrete→abstract link
    // must not persist (serializeNumbering only nulls on fully-empty maps).
    if (payload && !this.hasLiveNumberingForSync(ast)) payload = null;
    if (!payload) {
      this.pruneNumberingFromZip();
      return;
    }
    assertNoXxe(payload, 'word/numbering.xml');

    const numberingPath = 'word/numbering.xml';
    if (this.zip.getEntry(numberingPath)) {
      this.zip.updateFile(numberingPath, Buffer.from(payload, 'utf8'));
    } else {
      this.zip.addFile(numberingPath, Buffer.from(payload, 'utf8'));
    }

    const hasNumberingRel = Object.values(ast.relationships ?? {}).some(
      rel => rel?.type === NUMBERING_REL_TYPE
    );
    if (!hasNumberingRel) {
      try {
        new RelationshipsManipulator(ast).addRelationship(NUMBERING_REL_TYPE, 'numbering.xml');
      } catch { /* rel sync stays best-effort — the part itself is written */ }
    }
  }

  /**
   * D1R2-1 orphan prune: removes a stale word/numbering.xml part together
   * with its document rel(s) and its [Content_Types].xml Override. Runs when
   * the AST carries no live numbering (empty maps OR no concrete→abstract
   * link). Best-effort throughout: a prune failure must never block a save
   * whose document.xml was already written. Never throws.
   */
  private pruneNumberingFromZip(): void {
    // 1. Drop the part itself.
    try {
      const part = this.zip.getEntry('word/numbering.xml');
      if (part) this.zip.deleteFile(part);
    } catch { /* best-effort */ }
    // 2. Drop numbering-typed rels from word/_rels/document.xml.rels.
    try {
      const relsPath = 'word/_rels/document.xml.rels';
      const relsEntry = this.zip.getEntry(relsPath);
      if (relsEntry) {
        const raw = relsEntry.getData().toString('utf8');
        assertNoXxe(raw, relsPath);
        assertNoXxeAdjacent(raw, relsPath);
        const relsDoc = new DOMParser().parseFromString(raw, 'text/xml');
        if (relsDoc?.documentElement && relsDoc.getElementsByTagName('parsererror').length === 0) {
          const rels = Array.from(relsDoc.getElementsByTagName('Relationship'));
          let dirty = false;
          for (const el of rels) {
            if ((el as Element).getAttribute('Type') === NUMBERING_REL_TYPE) {
              (el as Element).parentNode?.removeChild(el as Element);
              dirty = true;
            }
          }
          if (dirty) {
            this.zip.updateFile(relsPath, Buffer.from(this.serializer.serializeToString(relsDoc), 'utf8'));
          }
        }
      }
    } catch (e) {
      if (e instanceof Error && (/XXE|DTD|E_VALIDATION/.test(e.message))) throw e;
      /* best-effort */
    }
    // 3. Drop the numbering Override from [Content_Types].xml.
    try {
      const typesPath = '[Content_Types].xml';
      const typesEntry = this.zip.getEntry(typesPath);
      if (typesEntry) {
        const raw = typesEntry.getData().toString('utf8');
        assertNoXxe(raw, typesPath);
        assertNoXxeAdjacent(raw, typesPath);
        const typesDoc = new DOMParser().parseFromString(raw, 'text/xml');
        if (typesDoc?.documentElement && typesDoc.getElementsByTagName('parsererror').length === 0) {
          const overrides = Array.from(typesDoc.getElementsByTagName('Override'));
          let dirty = false;
          for (const ov of overrides) {
            if ((ov as Element).getAttribute('PartName') === '/word/numbering.xml') {
              (ov as Element).parentNode?.removeChild(ov as Element);
              dirty = true;
            }
          }
          if (dirty) {
            this.zip.updateFile(typesPath, Buffer.from(this.serializer.serializeToString(typesDoc), 'utf8'));
          }
        }
      }
    } catch (e) {
      if (e instanceof Error && (/XXE|DTD|E_VALIDATION/.test(e.message))) throw e;
      /* best-effort */
    }
  }

  /**
   * D1R2-1 gate shared by part/rel/CT sync (implemented here so
   * numbering.ts stays untouched): true iff at least one concrete w:num
   * references an existing w:abstractNum via w:abstractNumId/@w:val. Only
   * live numbering earns a part, a rel, or a content-type Override.
   */
  private hasLiveNumberingForSync(ast: DocumentAST): boolean {
    return hasLiveNumberingForSync(ast);
  }

  /**
   * BX-03 bare-rezip companion: ensures the numbering relationship + its
   * content-type override exist in the zip WITHOUT merging the rest of the
   * cached AST's relationships/media (which may be stale vs direct zip writes
   * such as hyperlink/comment ops). Scoped to the single numbering part:
   * - if any numbering-typed rel already exists in document.xml.rels, do
   *   nothing (the part is already reachable; avoids duplicate rels);
   * - otherwise upsert the AST's numbering rel entry (Id/Type/Target) once;
   * - ensure [Content_Types].xml carries the numbering Override.
   * No-op when the cached AST holds no numbering rel or the numbering part was
   * never written (documents without lists stay untouched).
   */
  private syncNumberingRelAndContentTypeBare(ast: DocumentAST): void {
    try {
      if (!this.zip.getEntry('word/numbering.xml')) return;
      let numberingRId: string | null = null;
      let numberingTarget: string | null = null;
      for (const [rId, rel] of Object.entries(ast.relationships ?? {})) {
        if (rel?.type === NUMBERING_REL_TYPE && rel?.target) {
          numberingRId = rId;
          numberingTarget = rel.target;
          break;
        }
      }
      if (!numberingRId || !numberingTarget) return;

      const relsPath = 'word/_rels/document.xml.rels';
      const relsEntry = this.zip.getEntry(relsPath);
      if (relsEntry) {
        const lenientParser = new DOMParser();
        let relsDoc: Document | null = null;
        const rawRels = relsEntry.getData().toString('utf8');
        assertNoXxe(rawRels, relsPath);
        assertNoXxeAdjacent(rawRels, relsPath);
        try {
          const parsed = lenientParser.parseFromString(rawRels, 'text/xml');
          if (parsed?.documentElement && parsed.getElementsByTagName('parsererror').length === 0) {
            relsDoc = parsed;
          }
        } catch { relsDoc = null; }
        if (relsDoc) {
          const root = relsDoc.documentElement;
          const existing = Array.from(root.getElementsByTagName('Relationship'));
          let hasNumberingType = false;
          let hasExactId = false;
          for (const el of existing) {
            const t = (el as Element).getAttribute('Type') || '';
            const id = (el as Element).getAttribute('Id') || '';
            if (t === NUMBERING_REL_TYPE) hasNumberingType = true;
            if (id === numberingRId) hasExactId = true;
          }
          if (!hasNumberingType && !hasExactId) {
            const created = relsDoc.createElement('Relationship');
            created.setAttribute('Id', numberingRId);
            created.setAttribute('Type', NUMBERING_REL_TYPE);
            created.setAttribute('Target', numberingTarget);
            root.appendChild(created);
            this.zip.updateFile(relsPath, Buffer.from(this.serializer.serializeToString(relsDoc), 'utf8'));
          }
        }
      }

      const typesPath = '[Content_Types].xml';
      const typesEntry = this.zip.getEntry(typesPath);
      if (!typesEntry) return;
      const rawTypesBare = typesEntry.getData().toString('utf8');
      assertNoXxe(rawTypesBare, typesPath);
      assertNoXxeAdjacent(rawTypesBare, typesPath);
      try {
        const parsed = new DOMParser().parseFromString(rawTypesBare, 'text/xml');
        if (!parsed?.documentElement || parsed.getElementsByTagName('parsererror').length > 0) return;
        const root = parsed.documentElement;
        const partName = '/word/numbering.xml';
        const expectedCt = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
        const overrides = Array.from(root.getElementsByTagName('Override'));
        let found = false;
        for (const ov of overrides) {
          if ((ov as Element).getAttribute('PartName') === partName) {
            found = true;
            if ((ov as Element).getAttribute('ContentType') !== expectedCt) {
              (ov as Element).setAttribute('ContentType', expectedCt);
              this.zip.updateFile(typesPath, Buffer.from(this.serializer.serializeToString(parsed), 'utf8'));
            }
            break;
          }
        }
        if (!found) {
          const created = parsed.createElement('Override');
          created.setAttribute('PartName', partName);
          created.setAttribute('ContentType', expectedCt);
          root.appendChild(created);
          this.zip.updateFile(typesPath, Buffer.from(this.serializer.serializeToString(parsed), 'utf8'));
        }
      } catch { /* content-type hardening stays best-effort */ }
    } catch { /* bare numbering aux never blocks save */ }
  }

  /**
   * PARSER-01: upsert every AST relationship into word/_rels/document.xml.rels.
   * Merges into the existing rels part (never wipes entries missing from a stale
   * AST). No-op when ast.relationships is empty so bare rezip() calls are safe.
   */
  private syncRelationshipsToZip(ast: DocumentAST): void {
    const rels = ast.relationships;
    if (!rels || Object.keys(rels).length === 0) return;
    const relsPath = 'word/_rels/document.xml.rels';
    const lenientParser = new DOMParser();
    let relsDoc: Document | null = null;
    const existingEntry = this.zip.getEntry(relsPath);
    if (existingEntry) {
      const raw = existingEntry.getData().toString('utf8');
      assertNoXxe(raw, relsPath);
      assertNoXxeAdjacent(raw, relsPath);
      try {
        const parsed = lenientParser.parseFromString(raw, 'text/xml');
        if (parsed?.documentElement && parsed.getElementsByTagName('parsererror').length === 0) {
          relsDoc = parsed;
        }
      } catch {
        relsDoc = null;
      }
    }
    if (!relsDoc) {
      relsDoc = lenientParser.parseFromString(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`,
        'text/xml'
      );
    }
    const root = relsDoc.documentElement;
    const existingById = new Map<string, Element>();
    const existingNodes = root.getElementsByTagName('Relationship');
    for (let i = 0; i < existingNodes.length; i++) {
      const el = existingNodes[i] as Element;
      const id = el.getAttribute('Id');
      if (id) existingById.set(id, el);
    }
    // D1R2-1: never (re)stamp a numbering rel when the AST holds no live
    // numbering — the part was pruned, so the rel would dangle.
    const liveNumbering = this.hasLiveNumberingForSync(ast);
    for (const [rId, rel] of Object.entries(rels)) {
      if (!rId || !rel?.type || !rel?.target) continue;
      if (!liveNumbering && rel.type === NUMBERING_REL_TYPE) continue;
      const current = existingById.get(rId);
      if (current) {
        current.setAttribute('Type', rel.type);
        current.setAttribute('Target', rel.target);
        if (rel.targetMode === 'External') {
          current.setAttribute('TargetMode', 'External');
        } else if (current.hasAttribute('TargetMode')) {
          current.removeAttribute('TargetMode');
        }
      } else {
        const created = relsDoc.createElement('Relationship');
        created.setAttribute('Id', rId);
        created.setAttribute('Type', rel.type);
        created.setAttribute('Target', rel.target);
        if (rel.targetMode === 'External') {
          created.setAttribute('TargetMode', 'External');
        }
        root.appendChild(created);
        existingById.set(rId, created);
      }
    }
    const serialized = this.serializer.serializeToString(relsDoc);
    if (this.zip.getEntry(relsPath)) {
      this.zip.updateFile(relsPath, Buffer.from(serialized, 'utf8'));
    } else {
      this.zip.addFile(relsPath, Buffer.from(serialized, 'utf8'));
    }
  }

  /**
   * PARSER-01: ensure [Content_Types].xml covers parts referenced by the AST.
   * Mirrors backend/operations/image.ts (Default per image extension) and
   * backend/manipulators/chart3d.ts (Override per chart/header part).
   */
  private syncContentTypesToZip(ast: DocumentAST): void {
    const hasMedia = ast.mediaFiles && Object.keys(ast.mediaFiles).length > 0;
    const hasRels = ast.relationships && Object.keys(ast.relationships).length > 0;
    if (!hasMedia && !hasRels) return;
    const typesPath = '[Content_Types].xml';
    const typesEntry = this.zip.getEntry(typesPath);
    if (!typesEntry) return;
    const lenientParser = new DOMParser();
    let typesDoc: Document;
    const rawTypes = typesEntry.getData().toString('utf8');
    assertNoXxe(rawTypes, typesPath);
    assertNoXxeAdjacent(rawTypes, typesPath);
    try {
      const parsed = lenientParser.parseFromString(rawTypes, 'text/xml');
      if (!parsed?.documentElement || parsed.getElementsByTagName('parsererror').length > 0) return;
      typesDoc = parsed;
    } catch {
      return;
    }
    const root = typesDoc.documentElement;
    let dirty = false;

    const ensureDefault = (ext: string, contentType: string) => {
      const defaults = Array.from(root.getElementsByTagName('Default'));
      for (const def of defaults) {
        if ((def as Element).getAttribute('Extension') === ext) return;
      }
      const created = typesDoc.createElement('Default');
      created.setAttribute('Extension', ext);
      created.setAttribute('ContentType', contentType);
      root.appendChild(created);
      dirty = true;
    };

    const ensureOverride = (partName: string, contentType: string) => {
      const overrides = Array.from(root.getElementsByTagName('Override'));
      for (const ov of overrides) {
        const el = ov as Element;
        if (el.getAttribute('PartName') === partName) {
          if (el.getAttribute('ContentType') !== contentType) {
            el.setAttribute('ContentType', contentType);
            dirty = true;
          }
          return;
        }
      }
      const created = typesDoc.createElement('Override');
      created.setAttribute('PartName', partName);
      created.setAttribute('ContentType', contentType);
      root.appendChild(created);
      dirty = true;
    };

    const IMAGE_CONTENT_TYPES: Record<string, string> = {
      png: 'image/png',
      jpeg: 'image/jpeg',
      jpg: 'image/jpeg',
      gif: 'image/gif',
      webp: 'image/webp',
      bmp: 'image/bmp',
      tif: 'image/tiff',
      tiff: 'image/tiff',
      svg: 'image/svg+xml',
      ico: 'image/x-icon',
      emf: 'image/x-emf',
      wmf: 'image/x-wmf'
    };

    if (hasMedia) {
      for (const fileName of Object.keys(ast.mediaFiles)) {
        const match = fileName.toLowerCase().match(/\.([a-z0-9]+)$/);
        if (!match) continue;
        const ext = match[1];
        const contentType = IMAGE_CONTENT_TYPES[ext] || `image/${ext}`;
        ensureDefault(ext, contentType);
      }
    }

    if (hasRels) {
      // D1R2-1: an orphaned numbering rel lingering in the AST map must not
      // earn a content-type Override for a part that was pruned.
      const liveNumbering = this.hasLiveNumberingForSync(ast);
      for (const rel of Object.values(ast.relationships)) {
        if (!rel?.type || !rel?.target) continue;
        if (!liveNumbering && rel.type === NUMBERING_REL_TYPE) continue;
        // Image targets are covered by Default entries above.
        if (rel.type.endsWith('/image') || rel.target.startsWith('media/') || rel.target.includes('/media/')) continue;
        const contentType = contentTypeForRelationship(rel.type, rel.target);
        if (!contentType) continue;
        const partName = partNameForRelationshipTarget(rel.target);
        if (!partName) continue;
        ensureOverride(partName, contentType);
      }
    }

    if (dirty) {
      const serialized = this.serializer.serializeToString(typesDoc);
      this.zip.updateFile(typesPath, Buffer.from(serialized, 'utf8'));
    }
  }

  /**
   * Validates this document against official Microsoft Office ECMA-376 XSD schemas using @attentionform/ooxml-validator.
   * Fail-closed save guards (duplicate paraIds, unknown tags) refuse as
   * `{ok:false}` entries here — every `if (!validation.ok)` caller keeps
   * working unchanged — while direct-`rezip()` writers get the thrown
   * E_VALIDATION. Integrity/XXE/macro throws still propagate, as before.
   */
  public async validateOOXML(options?: { officeVersion?: 'Office2019' | 'Microsoft365' }): Promise<ValidationResult> {
    let buf: Buffer;
    try {
      buf = this.rezip();
    } catch (err) {
      const mapped = guardFailureToValidationErrors(err);
      if (mapped) return { ok: false, errors: mapped } as unknown as ValidationResult;
      throw err;
    }
    return await validateDocxBuffer(buf, options);
  }

  private convertDelTextNodes(run: Element) {
    const delTexts = Array.from(run.getElementsByTagName('w:delText'));
    delTexts.forEach((delText) => {
      const textNode = createWordElement(this.documentDom, 'w:t');
      if (delText.attributes) {
        for (let i = 0; i < delText.attributes.length; i++) {
          const attr = delText.attributes[i];
          textNode.setAttribute(attr.name, attr.value);
        }
      }
      textNode.appendChild(this.documentDom.createTextNode(delText.textContent || ''));
      delText.parentNode?.replaceChild(textNode, delText);
    });
  }

  private replaceDeletionWithText(delElement: Element) {
    const parent = delElement.parentNode;
    if (!parent) return;

    const replacements: Node[] = [];
    for (let i = 0; i < delElement.childNodes.length; i++) {
      const child = delElement.childNodes[i];
      if (child.nodeType !== 1) continue;

      const childEl = child as Element;
      if (childEl.tagName === 'w:r') {
        const clonedRun = childEl.cloneNode(true) as Element;
        this.convertDelTextNodes(clonedRun);
        replacements.push(clonedRun);
      } else if (childEl.tagName === 'w:delText') {
        const run = createWordElement(this.documentDom, 'w:r');
        const textNode = createWordElement(this.documentDom, 'w:t');
        if (childEl.attributes) {
          for (let a = 0; a < childEl.attributes.length; a++) {
            const attr = childEl.attributes[a];
            textNode.setAttribute(attr.name, attr.value);
          }
        }
        textNode.appendChild(this.documentDom.createTextNode(childEl.textContent || ''));
        run.appendChild(textNode);
        replacements.push(run);
      }
    }

    replacements.forEach((node) => parent.insertBefore(node, delElement));
    parent.removeChild(delElement);
  }

  public applyTrackedChanges(mode: 'accept' | 'reject') {
    if (mode === 'accept') {
      const insNodes = Array.from(this.documentDom.getElementsByTagName('w:ins'));
      insNodes.forEach((node) => unwrapElement(node as Element));
      const delNodes = Array.from(this.documentDom.getElementsByTagName('w:del'));
      delNodes.forEach((node) => node.parentNode?.removeChild(node));
      this._bumpNodePagesRevision();
      return;
    }

    const insNodes = Array.from(this.documentDom.getElementsByTagName('w:ins'));
    insNodes.forEach((node) => node.parentNode?.removeChild(node));
    const delNodes = Array.from(this.documentDom.getElementsByTagName('w:del'));
    delNodes.forEach((node) => this.replaceDeletionWithText(node as Element));
    this._bumpNodePagesRevision();
  }

  public getStylesDom(): Document | null {
    return this.stylesDom;
  }

  public getStylePalette(): StylePalette {
    if (!this.stylesDom) {
      return {
        primaryFont: 'Calibri',
        brandColors: [],
        availableStyles: [],
        headingStyles: []
      };
    }
    return extractStylePalette(this.stylesDom);
  }

  public allocateNextTrackedChangeId(): string {
    const insNodes = this.documentDom.getElementsByTagName('w:ins');
    const delNodes = this.documentDom.getElementsByTagName('w:del');
    let maxId = 0;
    for (let i = 0; i < insNodes.length; i++) {
      const id = parseInt(insNodes[i].getAttribute('w:id') || '0', 10);
      if (!isNaN(id) && id > maxId) maxId = id;
    }
    for (let i = 0; i < delNodes.length; i++) {
      const id = parseInt(delNodes[i].getAttribute('w:id') || '0', 10);
      if (!isNaN(id) && id > maxId) maxId = id;
    }
    return String(maxId + 1);
  }

  public insertTrackedChange(targetNodeId: string, text: string, author: string = 'ASTra', date?: string, reproducible?: boolean): void {
    const target = this.nodeMap.get(targetNodeId);
    if (!target) {
      throw new Error(`Node with ID "${targetNodeId}" not found in document.`);
    }

    const stamp = date ?? deterministicRevisionIso(reproducible);
    const changeId = this.allocateNextTrackedChangeId();
    const insEl = createWordElement(this.documentDom, 'w:ins');
    insEl.setAttribute('w:id', changeId);
    insEl.setAttribute('w:author', author);
    insEl.setAttribute('w:date', stamp);

    const rEl = createWordElement(this.documentDom, 'w:r');
    const tEl = createWordElement(this.documentDom, 'w:t');
    tEl.setAttribute('xml:space', 'preserve');
    tEl.appendChild(this.documentDom.createTextNode(text));
    rEl.appendChild(tEl);
    insEl.appendChild(rEl);

    if (target.tagName === 'w:p') {
      target.appendChild(insEl);
    } else if (target.parentNode) {
      target.parentNode.insertBefore(insEl, target.nextSibling);
    } else {
      target.appendChild(insEl);
    }
    this._bumpNodePagesRevision();
  }

  public deleteTrackedChange(targetNodeId: string, author: string = 'ASTra', date?: string, reproducible?: boolean): void {
    const target = this.nodeMap.get(targetNodeId);
    if (!target) {
      throw new Error(`Node with ID "${targetNodeId}" not found in document.`);
    }

    const stamp = date ?? deterministicRevisionIso(reproducible);
    const changeId = this.allocateNextTrackedChangeId();
    const delEl = createWordElement(this.documentDom, 'w:del');
    delEl.setAttribute('w:id', changeId);
    delEl.setAttribute('w:author', author);
    delEl.setAttribute('w:date', stamp);

    if (target.tagName === 'w:p') {
      const runs = Array.from(target.getElementsByTagName('w:r'));
      if (runs.length > 0) {
        for (const run of runs) {
          const tNodes = Array.from(run.getElementsByTagName('w:t'));
          for (const t of tNodes) {
            const delRun = createWordElement(this.documentDom, 'w:r');
            const delText = createWordElement(this.documentDom, 'w:delText');
            delText.setAttribute('xml:space', 'preserve');
            delText.appendChild(this.documentDom.createTextNode(t.textContent || ''));
            delRun.appendChild(delText);
            delEl.appendChild(delRun);
          }
          if (run.parentNode === target) {
            target.removeChild(run);
          }
        }
        target.appendChild(delEl);
      } else {
        const delRun = createWordElement(this.documentDom, 'w:r');
        const delText = createWordElement(this.documentDom, 'w:delText');
        delText.setAttribute('xml:space', 'preserve');
        delText.appendChild(this.documentDom.createTextNode(target.textContent || ''));
        delRun.appendChild(delText);
        delEl.appendChild(delRun);
        target.appendChild(delEl);
      }
    } else if (target.parentNode) {
      const delRun = createWordElement(this.documentDom, 'w:r');
      const delText = createWordElement(this.documentDom, 'w:delText');
      delText.setAttribute('xml:space', 'preserve');
      delText.appendChild(this.documentDom.createTextNode(target.textContent || ''));
      delRun.appendChild(delText);
      delEl.appendChild(delRun);
      target.parentNode.insertBefore(delEl, target);
      target.parentNode.removeChild(target);
    }
    this._bumpNodePagesRevision();
  }

  public updateDocumentProperties(properties: { 
    margins?: { top?: number; bottom?: number; left?: number; right?: number },
    pageSize?: { preset?: 'A4' | 'Letter' | 'Legal'; width?: number; height?: number; orientation?: 'portrait' | 'landscape' },
    defaultFont?: { name?: string; size?: number; ascii?: string; hAnsi?: string; eastAsia?: string; cs?: string }
  }): void {
    const sectPrs = this.documentDom.getElementsByTagName('w:sectPr');
    const pageSizePresets: Record<string, { width: number; height: number }> = {
      A4: { width: 11906, height: 16838 },
      Letter: { width: 12240, height: 15840 },
      Legal: { width: 12240, height: 20160 }
    };

    for (let i = 0; i < sectPrs.length; i++) {
      const sectPr = sectPrs[i];
      // Replace-in-place with schema-correct child order (pgSz must precede pgMar)
      const pgMar = upsertSectPrChild(sectPr, 'w:pgMar');
      
      if (properties.margins) {
        if (properties.margins.top !== undefined) pgMar.setAttribute('w:top', properties.margins.top.toString());
        if (properties.margins.bottom !== undefined) pgMar.setAttribute('w:bottom', properties.margins.bottom.toString());
        if (properties.margins.left !== undefined) pgMar.setAttribute('w:left', properties.margins.left.toString());
        if (properties.margins.right !== undefined) pgMar.setAttribute('w:right', properties.margins.right.toString());
      }

      if (properties.pageSize) {
        const pgSz = upsertSectPrChild(sectPr, 'w:pgSz');

        let width = properties.pageSize.width;
        let height = properties.pageSize.height;
        if (properties.pageSize.preset) {
          const preset = pageSizePresets[properties.pageSize.preset];
          if (preset) {
            width = preset.width;
            height = preset.height;
          }
        }

        if (properties.pageSize.orientation) {
          pgSz.setAttribute('w:orient', properties.pageSize.orientation);
        }

        if (width && height) {
          if (properties.pageSize.orientation === 'landscape' && height > width) {
            [width, height] = [height, width];
          }
          pgSz.setAttribute('w:w', width.toString());
          pgSz.setAttribute('w:h', height.toString());
        }
      }
    }

    if (properties.defaultFont) {
      if (!this.stylesDom) throw new Error('word/styles.xml not found in this document');

      const stylesRoot = this.stylesDom.documentElement;
      const docDefaults = getOrCreateChild(stylesRoot, 'w:docDefaults', this.stylesDom);
      const rPrDefault = getOrCreateChild(docDefaults, 'w:rPrDefault', this.stylesDom);
      const rPr = getOrCreateChild(rPrDefault, 'w:rPr', this.stylesDom);
      const rFonts = getOrCreateChild(rPr, 'w:rFonts', this.stylesDom);

      const fallbackFont = properties.defaultFont.name;
      const ascii = properties.defaultFont.ascii ?? fallbackFont;
      const hAnsi = properties.defaultFont.hAnsi ?? fallbackFont;
      const eastAsia = properties.defaultFont.eastAsia ?? fallbackFont;
      const cs = properties.defaultFont.cs ?? fallbackFont;

      if (ascii) rFonts.setAttribute('w:ascii', ascii);
      if (hAnsi) rFonts.setAttribute('w:hAnsi', hAnsi);
      if (eastAsia) rFonts.setAttribute('w:eastAsia', eastAsia);
      if (cs) rFonts.setAttribute('w:cs', cs);

      if (properties.defaultFont.size !== undefined) {
        const halfPoints = Math.round(properties.defaultFont.size * 2);
        const sz = getOrCreateChild(rPr, 'w:sz', this.stylesDom);
        const szCs = getOrCreateChild(rPr, 'w:szCs', this.stylesDom);
        sz.setAttribute('w:val', halfPoints.toString());
        szCs.setAttribute('w:val', halfPoints.toString());
      }
    }
    // Geometry (margins/page-size) feeds linesPerPageFromSectPr, so the cached
    // embedded map must invalidate too.
    this._bumpNodePagesRevision();
  }

  public static createNewDocument(templateType: 'blank' | 'novel' | 'financial' | 'api-spec' | 'research' | 'corporate' | 'bar-chart-3d' | string = 'blank'): DocxParser {
    // F2a item 2: packaged-aware search (static extra + STATIC_TEMPLATE_DIRS
    // env first, legacy cwd/__dirname after). Hook-only — lookup loop below
    // is unchanged.
    const candidatePaths = DocxParser.resolveTemplateSearchPaths(templateType);
    for (const p of candidatePaths) {
      if (fs.existsSync(p)) {
        try {
          const buf = fs.readFileSync(p);
          return new DocxParser(buf);
        } catch { /* ignore — try next template path */ }
      }
    }

    const zip = new AdmZip();

    const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>${NEW_DOCUMENT_CONTENT_TYPE_OVERRIDES}
</Types>`;

    const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>${NEW_DOCUMENT_ROOT_RELS}
</Relationships>`;

    const documentRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>${NEW_DOCUMENT_DOCUMENT_RELS}
</Relationships>`;

    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri" w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:cstheme="minorBidi"/>
        <w:sz w:val="22"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="9"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:widowControl/>
      <w:spacing w:before="360" w:after="160" w:line="276" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:color w:val="0F172A"/>
      <w:sz w:val="48"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="9"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:widowControl/>
      <w:spacing w:before="280" w:after="120" w:line="260" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:color w:val="1E293B"/>
      <w:sz w:val="32"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="9"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:widowControl/>
      <w:spacing w:before="200" w:after="80" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:color w:val="334155"/>
      <w:sz w:val="26"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:spacing w:before="0" w:after="240" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:color w:val="0F172A"/>
      <w:sz w:val="56"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle">
    <w:name w:val="Subtitle"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:spacing w:before="0" w:after="360" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
      <w:i/>
      <w:color w:val="6366F1"/>
      <w:sz w:val="28"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Caption">
    <w:name w:val="caption"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="9"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepLines/>
      <w:spacing w:before="120" w:after="120" w:line="240" w:lineRule="auto"/>
    </w:pPr>
  </w:style>
  <w:style w:type="character" w:styleId="Code">
    <w:name w:val="Code"/>
    <w:rPr>
      <w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>
      <w:color w:val="4338CA"/>
      <w:sz w:val="20"/>
    </w:rPr>
  </w:style>
</w:styles>`;

    let bodyContentXml = '';

    switch (templateType.toLowerCase()) {
      case 'novel':
        bodyContentXml = `
    <w:p>
      <w:pPr><w:pStyle w:val="Title"/></w:pPr>
      <w:r><w:t>Manuscript Title</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Subtitle"/></w:pPr>
      <w:r><w:t>A Novel by Author Name</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Chapter 1: The Beginning</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>The story starts here. Write your opening narrative...</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>Scene I</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Describe the setting, atmospheric tension, and character motives...</w:t></w:r>
    </w:p>`;
        break;

      case 'report':
        bodyContentXml = `
    <w:p>
      <w:pPr><w:pStyle w:val="Title"/></w:pPr>
      <w:r><w:t>Q3 Executive Financial Performance Report</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Subtitle"/></w:pPr>
      <w:r><w:t>Strategic Financial Analysis &amp; Growth Projections</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>1. Executive Summary</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>During Q3 2026, corporate revenue expanded by 34.2% YoY driven by cloud infrastructure enterprise adoption.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>1.1 Revenue &amp; Operating Margin Analysis</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Operating margins reached record expansion through workflow automation and infrastructure efficiencies.</w:t></w:r>
    </w:p>`;
        break;

      case 'financial':
        bodyContentXml = `
    <w:p>
      <w:pPr><w:pStyle w:val="Title"/></w:pPr>
      <w:r><w:t>Financial Report &amp; Audit</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Subtitle"/></w:pPr>
      <w:r><w:t>Executive Summary &amp; Revenue Projections</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>1. Key Financial Highlights</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Overview of performance metrics for the current reporting period.</w:t></w:r>
    </w:p>
    <w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="TableGrid"/>
        <w:tblW w:w="0" w:type="auto"/>
      </w:tblPr>
      <w:tblGrid>
        <w:gridCol w:w="2340"/>
        <w:gridCol w:w="2340"/>
        <w:gridCol w:w="2340"/>
        <w:gridCol w:w="2340"/>
      </w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:pPr><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>Metric</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:pPr><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>Previous Q</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:pPr><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>Current Q</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:pPr><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>Growth (%)</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Net Revenue</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>$1,250,000</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>$1,520,000</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>+21.6%</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Operating Profit</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>$340,000</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>$415,000</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>+22.0%</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>`;
        break;

      case 'academic':
        bodyContentXml = `
    <w:p>
      <w:pPr><w:pStyle w:val="Title"/></w:pPr>
      <w:r><w:t>Neural Document Architectures: A Comparative Benchmark</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Subtitle"/></w:pPr>
      <w:r><w:t>Department of Computer Science &amp; Systems Engineering</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Abstract</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Recent breakthroughs in autonomous document synthesis demonstrate marked improvements in structural coherence and typographical fidelity.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>1. Introduction</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Document processing pipelines must maintain rigorous schema compliance across complex OpenXML structures while supporting high-throughput transformations.</w:t></w:r>
    </w:p>`;
        break;

      case 'newsletter':
        bodyContentXml = `
    <w:p>
      <w:pPr><w:pStyle w:val="Title"/></w:pPr>
      <w:r><w:t>The ASTra Dispatch — Monthly Engineering Gazette</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Subtitle"/></w:pPr>
      <w:r><w:t>Volume 12 • Fall Edition 2026</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Highlights: Next-Gen Autonomous Publishing</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Welcome to this month's digest of advancements in agentic tools, typographic guards, and developer infrastructure.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>Community &amp; Ecosystem Updates</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Our contributor base expanded significantly following recent performance benchmarks and open schema validators.</w:t></w:r>
    </w:p>`;
        break;

      case 'api':
      case 'api-spec':
        bodyContentXml = `
    <w:p>
      <w:pPr><w:pStyle w:val="Title"/></w:pPr>
      <w:r><w:t>ASTra Document Engine API Specification</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Subtitle"/></w:pPr>
      <w:r><w:t>RESTful Endpoints &amp; Developer Reference v1.0</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>1. Overview &amp; Authentication</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>All API requests must include a valid Bearer token in the HTTP Authorization header.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>GET /api/v1/documents</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Retrieves a list of accessible documents with pagination parameters.</w:t></w:r>
    </w:p>`;
        break;

      case 'research':
        bodyContentXml = `
    <w:p>
      <w:pPr><w:pStyle w:val="Title"/></w:pPr>
      <w:r><w:t>Research Benchmark Report</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Subtitle"/></w:pPr>
      <w:r><w:t>Comparative Analysis &amp; Systems Evaluation</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>1. Methodology &amp; Scope</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Comprehensive benchmark evaluation across cloud workloads.</w:t></w:r>
    </w:p>`;
        break;

      case 'corporate':
        bodyContentXml = `
    <w:p>
      <w:pPr><w:pStyle w:val="Title"/></w:pPr>
      <w:r><w:t>Corporate Strategy Roadmap</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Subtitle"/></w:pPr>
      <w:r><w:t>Enterprise Governance &amp; Multi-Year Objectives</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>1. Strategic Vision</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Roadmap outlining enterprise transformation, governance, and operational agility.</w:t></w:r>
    </w:p>`;
        break;

      case 'bar-chart-3d':
        bodyContentXml = `
    <w:p>
      <w:pPr><w:pStyle w:val="Title"/></w:pPr>
      <w:r><w:t>Velocity Audit &amp; Performance</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Subtitle"/></w:pPr>
      <w:r><w:t>SDLC Metrics &amp; Visualization</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>1. Delivery Metrics</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Distribution of engineering cycle times and sprint delivery velocity.</w:t></w:r>
    </w:p>`;
        break;

      // FINDING-01 / P1-ENG-2: memo/letter/proposal/invoice previously fell
      // through to the blank boilerplate (valid OOXML but indistinguishable).
      // Each now gets minimal bespoke content (Title + one body line naming
      // the template) so all 6 templates validate ok:true AND are distinct.
      // No tables here (keeps tblGrid/validation trivially green).
      case 'memo':
        bodyContentXml = `
    <w:p>
      <w:pPr><w:pStyle w:val="Title"/></w:pPr>
      <w:r><w:t>MEMORANDUM</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Memo template — To / From / Date / Subject, followed by the body.</w:t></w:r>
    </w:p>`;
        break;

      case 'letter':
        bodyContentXml = `
    <w:p>
      <w:pPr><w:pStyle w:val="Title"/></w:pPr>
      <w:r><w:t>Letter</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Letter template — sender address, date, recipient, salutation, and body.</w:t></w:r>
    </w:p>`;
        break;

      case 'proposal':
        bodyContentXml = `
    <w:p>
      <w:pPr><w:pStyle w:val="Title"/></w:pPr>
      <w:r><w:t>Proposal</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Proposal template — executive summary, scope, timeline, and pricing.</w:t></w:r>
    </w:p>`;
        break;

      case 'invoice':
        bodyContentXml = `
    <w:p>
      <w:pPr><w:pStyle w:val="Title"/></w:pPr>
      <w:r><w:t>INVOICE</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Invoice template — bill-to, invoice number, line items, and total due.</w:t></w:r>
    </w:p>`;
        break;

      case 'blank':
      default:
        // FINDING-02: root legacy docs (demo.docx etc) carry non-hex paraId /
        // missing tblGrid and are flagged by the validator; they are legacy
        // binaries (DO NOT EDIT). Gallery T1-T9 are the canonical valid set.
        bodyContentXml = `
    <w:p>
      <w:pPr><w:pStyle w:val="Title"/></w:pPr>
      <w:r><w:t>Blank Document</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Subtitle"/></w:pPr>
      <w:r><w:t>Start typing your content here...</w:t></w:r>
    </w:p>`;
        break;
    }

    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:w10="urn:schemas-microsoft-com:office:word"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>${bodyContentXml}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>`;

    zip.addFile('[Content_Types].xml', Buffer.from(contentTypesXml, 'utf8'));
    zip.addFile('_rels/.rels', Buffer.from(relsXml, 'utf8'));
    zip.addFile('word/_rels/document.xml.rels', Buffer.from(documentRelsXml, 'utf8'));
    zip.addFile('word/styles.xml', Buffer.from(stylesXml, 'utf8'));
    zip.addFile('word/document.xml', Buffer.from(documentXml, 'utf8'));
    const pinned = resolveReproducibleMode(undefined);
    const parts = newDocumentParts({
      title: newDocumentTitle(templateType),
      author: resolveDefaultAuthor(),
      at: pinned ? resolveDeterministicZipDate() : new Date(),
    });
    for (const [partName, xml] of Object.entries(parts)) zip.addFile(partName, Buffer.from(xml, 'utf8'));

    if (pinned) clampZipEntryTimes(zip, resolveDeterministicZipDate());
    return new DocxParser(zip.toBuffer());
  }

  public static createEmptyDocument(templateType: 'blank' | 'novel' | 'financial' | 'api-spec' | 'research' | 'corporate' | 'bar-chart-3d' | string = 'blank'): DocxParser {
    return DocxParser.createNewDocument(templateType);
  }
}

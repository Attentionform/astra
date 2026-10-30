import pLimit from 'p-limit';
import { DocxParser } from './parser';
import { FileReadCache } from './read-cache';

export interface XMLWorkerRequest {
  nodeId: string;
  originalXml: string;
  instruction: string;
  styleSnapshot?: string;
  modelId: string;
}

export interface XMLWorkerResult {
  nodeId: string;
  success: boolean;
  newXml?: string;
  error?: string;
  rolledBack?: boolean;
}

export type XMLWorkerFn = (request: XMLWorkerRequest) => Promise<XMLWorkerResult>;

let defaultXmlWorker: XMLWorkerFn | undefined;

export function setDefaultXmlWorker(worker: XMLWorkerFn): void {
  defaultXmlWorker = worker;
}

export function getDefaultXmlWorker(): XMLWorkerFn | undefined {
  return defaultXmlWorker;
}

function extractTextFromXml(xml: string): string {
  const matches = xml.match(/<w:t[^>]*>(.*?)<\/w:t>/g);
  if (!matches) return '';
  return matches.map(t => t.replace(/<[^>]+>/g, '')).join('');
}

function estimateXmlTokens(xml: string): number {
  return Math.max(1, Math.ceil(xml.length / 3.5));
}

function chunkByTokens(nodeIds: string[], parser: DocxParser, maxTokensPerChunk: number = 3000): string[][] {
  const chunks: string[][] = [];
  let currentChunk: string[] = [];
  let currentTokens = 0;

  for (const id of nodeIds) {
    const xml = parser.readNode(id);
    const estimated = estimateXmlTokens(xml);
    if (currentTokens + estimated > maxTokensPerChunk && currentChunk.length > 0) {
      chunks.push(currentChunk);
      currentChunk = [];
      currentTokens = 0;
    }
    currentChunk.push(id);
    currentTokens += estimated;
  }

  if (currentChunk.length > 0) chunks.push(currentChunk);
  return chunks;
}

function buildSectionsByHeading(nodeIdsInOrder: { id: string; isHeading: boolean }[]): string[][] {
  const sections: string[][] = [];
  let current: string[] = [];

  for (const node of nodeIdsInOrder) {
    if (node.isHeading && current.length > 0) {
      sections.push(current);
      current = [];
    }
    current.push(node.id);
  }

  if (current.length > 0) sections.push(current);
  return sections;
}

export interface BatchProcessOptions {
  nodeIds: string[];
  instruction: string;
  strategy: 'mechanical' | 'semantic' | 'hierarchical';
  modelId: string;
  parser: DocxParser;
  readCache?: FileReadCache;
  onProgress?: (processed: number, total: number, lastResult?: XMLWorkerResult) => void;
  worker?: XMLWorkerFn;
}

/** Guard-narrowed view over parser snapshot duck-typing (no `as any`). */
interface BatchSnapshotCapable {
  capturePackageSnapshot?: () => Map<string, Buffer>;
  restorePackageSnapshot?: (snap: Map<string, Buffer>) => void;
  readDocumentXml?: () => string;
  reloadDocumentXml?: (xml: string) => void;
}

function asBatchSnapshotCapable(parser: unknown): BatchSnapshotCapable {
  if (typeof parser !== 'object' || parser === null) return {};
  const rec = parser as Record<string, unknown>;
  const out: BatchSnapshotCapable = {};
  if (typeof rec['capturePackageSnapshot'] === 'function') {
    out.capturePackageSnapshot = (rec['capturePackageSnapshot'] as () => Map<string, Buffer>).bind(parser);
  }
  if (typeof rec['restorePackageSnapshot'] === 'function') {
    out.restorePackageSnapshot = (rec['restorePackageSnapshot'] as (snap: Map<string, Buffer>) => void).bind(parser);
  }
  if (typeof rec['readDocumentXml'] === 'function') {
    out.readDocumentXml = (rec['readDocumentXml'] as () => string).bind(parser);
  }
  if (typeof rec['reloadDocumentXml'] === 'function') {
    out.reloadDocumentXml = (rec['reloadDocumentXml'] as (xml: string) => void).bind(parser);
  }
  return out;
}

function errorMessageOf(err: unknown): string {
  if (err instanceof Error) return err.message !== '' ? err.message : String(err);
  if (typeof err === 'object' && err !== null && 'message' in err) {
    const m = (err as { message?: unknown }).message;
    if (typeof m === 'string' && m !== '') return m;
  }
  return String(err);
}

export async function processBatch(options: BatchProcessOptions): Promise<XMLWorkerResult[]> {
  const { nodeIds, instruction, strategy, modelId, parser, readCache, onProgress, worker } = options;
  const executeWorker = worker || defaultXmlWorker;
  if (!executeWorker) {
    throw new Error("No XML worker provided for batch processing. Pass options.worker or call setDefaultXmlWorker(worker).");
  }
  const results: XMLWorkerResult[] = [];

  /**
   * OX-05: batch owns its snapshot (parser-level, NOT the engine's AST
   * clone). Rationale: this module mutates ONLY via parser.updateNode
   * (live document DOM + nodeMap) — it never touches an AST object, so an
   * AST clone would rewind nothing. Snapshot preference order:
   *   1. capturePackageSnapshot/restorePackageSnapshot (full zip, byte-exact
   *      — covers document.xml plus any side parts updateNode may touch);
   *   2. readDocumentXml/reloadDocumentXml (document DOM only);
   *   3. per-node readNode map + updateNode restore (minimal parser doubles
   *      in unit tests that expose only readNode/updateNode).
   * Apply failures aggregate into their result entries (success=false +
   * error) and rewind ALL already-applied updates via the snapshot — no
   * console.error swallow, no partial doc. LLM-stage failures
   * (processXmlNode success=false) skip apply individually and do NOT
   * trigger a rollback on their own: at apply time nothing has been written
   * yet (mechanical/semantic), and the hierarchical consistency pass is
   * best-effort refinement that must not discard main-stage work.
   * If the restore itself fails, throw loudly (rollback-or-FAIL).
   */
  const caps = asBatchSnapshotCapable(parser);
  const pkgCap = caps.capturePackageSnapshot;
  const pkgRes = caps.restorePackageSnapshot;
  const hasPkgSnap = typeof pkgCap === 'function' && typeof pkgRes === 'function';
  const hasDocSnap =
    typeof caps.readDocumentXml === 'function' &&
    typeof caps.reloadDocumentXml === 'function';
  let pkgSnap: Map<string, Buffer> | null = null;
  let docSnap: string | null = null;
  const nodeSnap = new Map<string, string>();
  try {
    if (hasPkgSnap) {
      pkgSnap = (pkgCap as () => Map<string, Buffer>).call(parser) as Map<string, Buffer>;
    } else if (hasDocSnap) {
      docSnap = (caps.readDocumentXml as () => string)() as string;
    } else {
      for (const id of nodeIds) {
        try {
          nodeSnap.set(id, parser.readNode(id));
        } catch { /* unreadable anchor — apply will report it */ }
      }
    }
  } catch { /* snapshot best-effort; apply errors still aggregate below */ }

  const restoreSnapshot = (): void => {
    if (pkgSnap) {
      (pkgRes as (snap: Map<string, Buffer>) => void).call(parser, pkgSnap);
      return;
    }
    if (docSnap !== null) {
      (caps.reloadDocumentXml as (xml: string) => void)(docSnap);
      return;
    }
  };

  const applyErrors: { nodeId: string; error: string }[] = [];
  const appliedIds: string[] = [];
  const applyOne = (result: XMLWorkerResult): void => {
    if (!(result.success && result.newXml)) return;
    try {
      parser.updateNode(result.nodeId, result.newXml);
      appliedIds.push(result.nodeId);
    } catch (e: unknown) {
      const msg = errorMessageOf(e);
      applyErrors.push({ nodeId: result.nodeId, error: msg });
      result.success = false;
      result.error = `Failed to apply updated XML to node ${result.nodeId}: ${msg}`;
      delete result.newXml;
    }
  };

  const rollbackOrFail = (): void => {
    if (applyErrors.length === 0) return;
    // Rewind everything applied so far in this batch.
    if (pkgSnap || docSnap !== null) {
      try {
        restoreSnapshot();
      } catch (e: unknown) {
        throw new Error(
          `Batch apply failed for ${applyErrors.map(a => a.nodeId).join(', ')} ` +
          `AND rollback failed: ${errorMessageOf(e)}. ` +
          `First apply error: ${applyErrors[0].nodeId}: ${applyErrors[0].error}`
        );
      }
    } else {
      // Minimal-double fallback: re-apply the per-node pre-images.
      const restoreErrors: string[] = [];
      for (const id of appliedIds) {
        const oldXml = nodeSnap.get(id);
        if (oldXml === undefined) continue;
        try {
          parser.updateNode(id, oldXml);
        } catch (e: unknown) {
          restoreErrors.push(`${id}: ${errorMessageOf(e)}`);
        }
      }
      if (restoreErrors.length > 0) {
        throw new Error(
          `Batch apply failed for ${applyErrors.map(a => a.nodeId).join(', ')} ` +
          `AND per-node rollback failed: ${restoreErrors.join('; ')}. ` +
          `First apply error: ${applyErrors[0].nodeId}: ${applyErrors[0].error}`
        );
      }
    }
    // Mark reverted successes so callers cannot mistake rolled-back LLM
    // output for live document state.
    for (const r of results) {
      if (r.success && appliedIds.includes(r.nodeId)) {
        r.rolledBack = true;
      }
    }
  };
  
  // Mark all nodes as read since batch operations read them internally
  if (readCache) {
    for (const nodeId of nodeIds) {
      const xml = parser.readNode(nodeId);
      readCache.markAsRead(nodeId, xml, true);
    }
  }
  
  if (strategy === 'mechanical') {
    // Mechanical: Concurrent execution, no cross-node dependency
    const limit = pLimit(8);
    let completed = 0;
    
    const tasks = nodeIds.map(nodeId => limit(async () => {
      const originalXml = parser.readNode(nodeId);
      const result = await executeWorker({
        nodeId,
        originalXml,
        instruction,
        modelId
      });
      
      completed++;
      if (onProgress) onProgress(completed, nodeIds.length, result);
      return result;
    }));
    
    const batchResults = await Promise.all(tasks);
    results.push(...batchResults);
    
  } else if (strategy === 'semantic') {
    // Semantic: Sequential execution, passing the previous node's text as a style snapshot
    let previousText = '';
    
    for (let i = 0; i < nodeIds.length; i++) {
      const nodeId = nodeIds[i];
      const originalXml = parser.readNode(nodeId);
      
      const result = await executeWorker({
        nodeId,
        originalXml,
        instruction,
        modelId,
        styleSnapshot: previousText ? previousText : undefined
      });
      
      if (result.success && result.newXml) {
        previousText = extractTextFromXml(result.newXml);
      }
      
      results.push(result);
      if (onProgress) onProgress(i + 1, nodeIds.length, result);
    }
  } else if (strategy === 'hierarchical') {
    const summary = parser.getStructuralSummary();
    const nodeIdSet = new Set(nodeIds);
    const orderedNodes = summary.nodes
      .filter(n => nodeIdSet.has(n.id))
      .map(n => ({ id: n.id, isHeading: n.isHeading }));

    const sections = buildSectionsByHeading(orderedNodes);
    const sectionLimit = pLimit(3);
    let processed = 0;
    const totalNodes = nodeIds.length;
    const sectionSummaries: { nodeIds: string[]; summaryText: string }[] = [];

    const sectionResults = await Promise.all(sections.map(section => sectionLimit(async () => {
      let styleSnapshot = '';
      const localResults: XMLWorkerResult[] = [];
      const chunks = chunkByTokens(section, parser, 3000);

      for (const chunk of chunks) {
        for (const nodeId of chunk) {
          const originalXml = parser.readNode(nodeId);
          const result = await executeWorker({
            nodeId,
            originalXml,
            instruction,
            modelId,
            styleSnapshot: styleSnapshot || undefined
          });
          if (result.success && result.newXml) {
            styleSnapshot = extractTextFromXml(result.newXml);
          }
          localResults.push(result);
          processed++;
          if (onProgress) onProgress(processed, totalNodes, result);
        }
      }

      const summaryText = localResults
        .filter(r => r.success && r.newXml)
        .map(r => extractTextFromXml(r.newXml!))
        .join(' ')
        .slice(0, 500);

      return { localResults, summaryText, sectionNodeIds: section };
    })));

    for (const section of sectionResults) {
      results.push(...section.localResults);
      sectionSummaries.push({ nodeIds: section.sectionNodeIds, summaryText: section.summaryText });
    }

    // Apply updates before the consistency pass
    for (const result of results) {
      applyOne(result);
    }
    // OX-05: a failed main-stage apply rewinds before the consistency pass
    // runs (it would otherwise refine a half-written document) — and the
    // consistency pass is skipped entirely so a failed batch cannot gain
    // new writes after the rewind. results.length may therefore exceed
    // nodeIds.length only when the consistency pass actually ran.
    rollbackOrFail();

    const consistencyNodeIds = applyErrors.length === 0 ? sectionSummaries.map(section => {
      const ordered = orderedNodes.filter(n => section.nodeIds.includes(n.id));
      const firstNonHeading = ordered.find(n => !n.isHeading);
      return (firstNonHeading || ordered[0])?.id;
    }).filter((id): id is string => Boolean(id)) : [];

    if (consistencyNodeIds.length > 0) {
      const globalSummary = sectionSummaries
        .map((s, idx) => `Section ${idx + 1}: ${s.summaryText}`)
        .join('\n')
        .slice(0, 500);

      for (const nodeId of consistencyNodeIds) {
        const originalXml = parser.readNode(nodeId);
        const result = await executeWorker({
          nodeId,
          originalXml,
          instruction: `${instruction}\n\nConsistency pass: adjust tone/voice to match the overall style described below while preserving meaning.`,
          modelId,
          styleSnapshot: globalSummary
        });
        results.push(result);
        if (result.success && result.newXml) {
          applyOne(result);
        }
        processed++;
        if (onProgress) onProgress(processed, totalNodes + consistencyNodeIds.length, result);
      }
      // OX-05: aggregate consistency-apply failures and rewind the whole
      // batch (main + consistency) via the pre-batch snapshot.
      rollbackOrFail();
    }
  }

  if (strategy !== 'hierarchical') {
    // Apply successful updates to the parser
    for (const result of results) {
      applyOne(result);
    }
    rollbackOrFail();
  }

  return results;
}

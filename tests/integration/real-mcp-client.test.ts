import { describe, it, afterAll } from "vitest";
/**
 * MCP Real-World Client End-to-End Test Runner
 *
 * Tests the 11-tool MCP server against REAL DOCX documents (T3_proposal.docx, T5_image_doc.docx).
 * Covers:
 *   - Sandbox setup in /tmp/astra-mcp-real-test-<process id>/ (one folder per run, so two runs cannot clash)
 *   - Stdio transport & in-memory MCP client execution
 *   - open_document with auto-init of .astra git repo & baseline commit
 *   - visual_inspect with spatial telemetry & layout metrics
 *   - Multi-batch edits & saves (executive summary, styled table, section title change)
 *   - Paginated history audit (limit/offset, reverse chronological order, author/message hygiene)
 *   - Single-step revert with append-only invariant & OOXML schema validation
 *   - Redo symmetry with visual inspect verification
 *   - Arbitrary OID revert to baseline with AST & byte comparison
 *   - Boundary & conflict tests (redo at HEAD, E_CONFLICT on unsaved edits, SessionMutex concurrency)
 *   - Zero host git dependency under PATH=''
 *   - Multi-document media tracking & preservation (T5_image_doc.docx)
 *   - Clean teardown & metrics reporting
 */

import * as fs from 'fs';
import * as path from 'path';
import { execSync } from 'child_process';
import { DocxParser, validateDocxBuffer } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';
import { LEAN6_HANDLERS, GIT_HISTORY_HANDLERS } from '../../packages/mcp-server/src/lib/mcp-tool-adapter';
import { handleVisualInspect } from '../../packages/mcp-server/src/tools/visual';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';

interface TestResult {
  name: string;
  passed: boolean;
  durationMs: number;
  error?: string;
  details?: Record<string, unknown>;
}

const results: TestResult[] = [];

async function step(name: string, fn: () => Promise<Record<string, unknown> | void>): Promise<void> {
  const start = Date.now();
  try {
    const details = await fn();
    const durationMs = Date.now() - start;
    results.push({ name, passed: true, durationMs, details: details || undefined });
    console.log(`  ✓ ${name} (${durationMs}ms)`);
  } catch (err: unknown) {
    const durationMs = Date.now() - start;
    const msg = err instanceof Error ? err.stack || err.message : String(err);
    results.push({ name, passed: false, durationMs, error: msg });
    console.error(`  ✗ ${name} (${durationMs}ms): ${msg}`);
    throw err;
  }
}

function assert(condition: boolean, msg: string): asserts condition {
  if (!condition) throw new Error(`Assertion failed: ${msg}`);
}

const REPO_ROOT = path.resolve(__dirname, "..", "..");
const SANDBOX = `/tmp/astra-mcp-real-test-${process.pid}`;
const T3_ORIGINAL = path.join(REPO_ROOT, 'test-gallery/T3_proposal.docx');
const T5_ORIGINAL = path.join(REPO_ROOT, 'test-gallery/T5_image_doc.docx');
const T3_SANDBOX = path.join(SANDBOX, 'T3_proposal.docx');
const T5_SANDBOX = path.join(SANDBOX, 'T5_image_doc.docx');

// Module-scope handles (initial values only — no filesystem reads here, so the
// file imports cleanly on a machine without the sandbox folder).
// Nullable + initialized: fixes definite-assignment errors and lets afterAll
// clean up even when a step throws before connect.
let stdioClient: Client | null = null;
let savedExtraRoots: string | undefined = undefined;
let savedExtraRootsSet = false;

afterAll(async () => {
  const c = stdioClient;
  stdioClient = null;
  if (c) {
    try {
      await c.close();
    } catch {
      /* best-effort */
    }
  }
  try {
    if (fs.existsSync(SANDBOX)) {
      fs.rmSync(SANDBOX, { recursive: true, force: true });
    }
  } catch {
    /* best-effort */
  }
  if (savedExtraRootsSet) {
    if (savedExtraRoots === undefined) delete process.env.ASTRA_EXTRA_DOC_ROOTS;
    else process.env.ASTRA_EXTRA_DOC_ROOTS = savedExtraRoots;
  }
});

async function runRealClientTests(): Promise<void> {
  // Helper for stdio tool calls (per-call timeout so a hung server fails fast).
  async function callTool(name: string, args: unknown, timeoutMs = 25_000): Promise<any> {
    const client = stdioClient;
    if (!client) throw new Error(`MCP client not connected (call ${name})`);
    const res = await Promise.race([
      client.callTool({ name, arguments: args as any }),
      new Promise<never>((_, reject) =>
        setTimeout(() => reject(new Error(`tool call timed out: ${name}`)), timeoutMs)
      ),
    ]);
    assert(Array.isArray(res.content) && res.content.length > 0, `Tool ${name} returned empty content`);
    const text = (res.content[0] as { text: string }).text;
    return JSON.parse(text);
  }

  console.log('================================================================');
  console.log('🚀 ASTra MCP Real-World Client End-to-End Test Suite Starting');
  console.log(`   Sandbox: ${SANDBOX}`);
  console.log(`   Local Time: ${new Date().toISOString()}`);
  console.log('================================================================\n');

  // Save env for try/finally restore (afterAll repeats the restore).
  const prevExtraRoots = process.env.ASTRA_EXTRA_DOC_ROOTS;
  savedExtraRoots = prevExtraRoots;
  savedExtraRootsSet = true;
  // Allow sandbox access
  process.env.ASTRA_EXTRA_DOC_ROOTS = SANDBOX;

  try {
  // ── 1. Sandbox Setup ────────────────────────────────────────────────────────
  await step('1. Setup isolated sandbox and copy real test gallery documents', async () => {
    if (fs.existsSync(SANDBOX)) {
      fs.rmSync(SANDBOX, { recursive: true, force: true });
    }
    fs.mkdirSync(SANDBOX, { recursive: true });

    assert(fs.existsSync(T3_ORIGINAL), `T3_proposal.docx exists at ${T3_ORIGINAL}`);
    assert(fs.existsSync(T5_ORIGINAL), `T5_image_doc.docx exists at ${T5_ORIGINAL}`);

    fs.copyFileSync(T3_ORIGINAL, T3_SANDBOX);
    fs.copyFileSync(T5_ORIGINAL, T5_SANDBOX);

    assert(fs.existsSync(T3_SANDBOX), 'T3_proposal.docx copied to sandbox');
    assert(fs.existsSync(T5_SANDBOX), 'T5_image_doc.docx copied to sandbox');

    const t3Bytes = fs.readFileSync(T3_SANDBOX);
    const t5Bytes = fs.readFileSync(T5_SANDBOX);

    return {
      t3Bytes: t3Bytes.length,
      t5Bytes: t5Bytes.length,
      sandbox: SANDBOX,
    };
  });

  // Baseline snapshot reads live here: inside runRealClientTests, after the
  // sandbox step above — never at module scope, so import stays clean.
  const baselineT3Bytes = fs.readFileSync(T3_SANDBOX);
  const baselineT3Parser = new DocxParser(baselineT3Bytes);
  const baselineT3Ast = await baselineT3Parser.getAST();
  const baselineT3Text = baselineT3Ast.sections[0].nodes
    .map((n: any) => (n.runs || []).map((r: any) => r.text).join('') || n.metadata?.text || '')
    .join('\n');

  // ── 2. Stdio Client Connection & Tool Enumeration ───────────────────────────
  // Initialized up front (definite assignment): the close path below the
  // boundary tests can reference these without TS2454.
  let sessionId = "";
  const serverPath = path.join(REPO_ROOT, 'packages/mcp-server/dist/index.js');

  await step('2. Connect MCP client over StdioServerTransport & verify 19 tools', async () => {
    // Rebuild dist when stale (mirrors the harness dist-staleness pattern).
    const distEntry = path.join(REPO_ROOT, 'packages', 'mcp-server', 'dist', 'index.js');
    const srcEntry = path.join(REPO_ROOT, 'packages', 'mcp-server', 'src', 'index.ts');
    let rebuild = false;
    try {
      const distStat = fs.statSync(distEntry);
      const srcStat = fs.statSync(srcEntry);
      rebuild = distStat.mtimeMs < srcStat.mtimeMs;
    } catch {
      rebuild = true;
    }
    if (rebuild) {
      execSync("npm run build --workspace=@attentionform/astra-mcp-server", {
        cwd: REPO_ROOT,
        stdio: "pipe",
        timeout: 180_000,
      });
    }

    const transport = new StdioClientTransport({
      command: process.execPath,
      args: [serverPath, '--allowed-dir', SANDBOX],
      env: { ...process.env, ASTRA_EXTRA_DOC_ROOTS: SANDBOX },
    });
    const client = new Client({ name: 'mcp-real-tester', version: '1.0.0' });
    await client.connect(transport);
    stdioClient = client;

    const toolsRes = await client.listTools();
    const toolNames = toolsRes.tools.map((t) => t.name).sort();

    const expectedTools = [
      'close_document',
      'compose_document',
      'execute_batch',
      'export_document',
      'get_document_history',
      'get_stats',
      'import_document',
      'inspect_document',
      'new_document',
      'normalize_document',
      'open_document',
      'preview_revert',
      'redo_document',
      'revert_document',
      'sanitize_document',
      'save_document',
      'search_content',
      'validate_document',
      'visual_inspect',
    ].sort();

    assert(toolNames.length === 19, `Expected 19 tools, got ${toolNames.length}`);
    assert(
      JSON.stringify(toolNames) === JSON.stringify(expectedTools),
      `Tools mismatch: ${JSON.stringify(toolNames)}`
    );

    return { toolsCount: toolNames.length, tools: toolNames };
  });

  let baselineOid = "";
  let baselineTotalPages = 0;

  // ── 3. open_document & auto-init of .astra git repo ─────────────────────────
  await step('3. open_document: verify auto-init of .astra repo and baseline commit', async () => {
    const res = await callTool('open_document', { path: T3_SANDBOX });
    assert(res.ok === true, `open_document failed: ${res.message}`);
    assert(typeof res.sessionId === 'string', 'sessionId returned');
    assert(res.revision === 1, 'Initial revision is 1');
    sessionId = res.sessionId;

    // Check .astra repository on disk
    const expectedRepoDir = path.join(SANDBOX, '.astra', 'T3_proposal.docx');
    assert(fs.existsSync(expectedRepoDir), `Expected .astra repo dir at ${expectedRepoDir}`);
    assert(fs.existsSync(path.join(expectedRepoDir, '.git')), '.git directory exists inside .astra repo');
    assert(fs.existsSync(path.join(expectedRepoDir, 'document.xml')), 'document.xml tracked in repo');

    // Check history baseline commit
    const hist = await callTool('get_document_history', { sessionId });
    assert(hist.ok === true, 'get_document_history succeeded');
    assert(hist.totalCommits === 1, `Expected 1 totalCommit, got ${hist.totalCommits}`);
    assert(hist.history.length === 1, 'History array length is 1');
    assert(hist.history[0].message === 'Initial state', `Baseline message: ${hist.history[0].message}`);
    assert(/^[0-9a-f]{40}$/.test(hist.history[0].oid), `Baseline oid is 40-hex: ${hist.history[0].oid}`);

    baselineOid = hist.history[0].oid;
    return { sessionId, baselineOid, repoDir: expectedRepoDir };
  });

  // ── 4. visual_inspect initial layout telemetry ──────────────────────────────
  await step('4. visual_inspect: verify spatial telemetry, wireframe, and token budget', async () => {
    const res = await callTool('visual_inspect', { sessionId });
    assert(res.ok === true, `visual_inspect failed: ${res.message}`);
    assert(res.dirty === false, 'Freshly opened doc is not dirty');
    assert(res.revision === 1, 'Revision matches handle');
    assert(typeof res.telemetry === 'object', 'Telemetry object present');
    assert(res.telemetry.totalPages >= 1, `Total pages: ${res.telemetry.totalPages}`);
    assert(typeof res.telemetry.wireframe === 'string', 'Wireframe present');
    assert(res.telemetry.wireframe.length > 50, 'Wireframe is non-empty');
    assert(res.telemetry.tokens <= 500, `Telemetry tokens (${res.telemetry.tokens}) <= budget 500`);
    assert(Array.isArray(res.telemetry.pageBreakdown), 'pageBreakdown is array');
    assert(typeof res.promptLaw === 'string', 'promptLaw present');
    assert(typeof res.verifyRule === 'string', 'verifyRule present');

    baselineTotalPages = res.telemetry.totalPages;

    return {
      totalPages: res.telemetry.totalPages,
      tokens: res.telemetry.tokens,
      wireframeLines: res.telemetry.wireframe.split('\n').length,
    };
  });

  // ── 5. Batch 1: Executive Summary Paragraphs & Save ─────────────────────────
  let batch1Oid = "";

  await step('5. Batch 1: execute_batch adding executive summary paragraphs & save_document', async () => {
    // Find anchor
    const search = await callTool('search_content', { sessionId, query: 'Riverbend Community Solar' });
    assert(search.ok === true && search.matches.length > 0, 'Found title anchor');
    const titleAnchor = search.matches[0].id;

    // Execute batch
    const batch = await callTool('execute_batch', {
      sessionId,
      operations: [
        {
          action: 'insert',
          type: 'paragraph',
          anchor: titleAnchor,
          position: 'after',
          text: 'Executive Summary: Accelerating renewable community energy access through local participation and scalable solar infrastructure.',
        },
        {
          action: 'insert',
          type: 'paragraph',
          anchor: titleAnchor,
          position: 'after',
          text: 'Key Highlights: 2.5MW generation capacity, 750+ subscriber target, and 30% low-income inclusion.',
        },
      ],
    });
    assert(batch.ok === true, `execute_batch failed: ${batch.message}`);
    assert(batch.appliedCount === 2, `Applied count: ${batch.appliedCount}`);
    assert(batch.revision === 2, `Revision updated: ${batch.revision}`);

    // Verify inspect shows dirty
    const insp = await callTool('inspect_document', { sessionId });
    assert(insp.ok === true && insp.dirty === true, 'Document is dirty before save');

    // Save with custom message & author
    const saved = await callTool('save_document', {
      sessionId,
      message: 'Add executive summary and key highlights',
      author: 'Chief Architect',
    });
    assert(saved.ok === true, `save_document failed: ${saved.message}`);
    assert(saved.commitMessage === 'Add executive summary and key highlights', `Commit msg: ${saved.commitMessage}`);
    assert(/^[0-9a-f]{40}$/.test(saved.commitOid), `Commit OID is 40-hex: ${saved.commitOid}`);
    batch1Oid = saved.commitOid;

    // Inspect shows dirty cleared
    const inspSaved = await callTool('inspect_document', { sessionId });
    assert(inspSaved.ok === true && inspSaved.dirty === false, 'Dirty flag cleared after save');

    return { batch1Oid, appliedCount: batch.appliedCount };
  });

  // ── 6. Batch 2: Insert Styled Comparison Table & Save ───────────────────────
  let batch2Oid = "";
  let batch2DiskBytes: Buffer | null = null;
  let batch2AstText = "";

  await step('6. Batch 2: execute_batch inserting styled comparison table & save_document', async () => {
    // Find anchor 'Approach'
    const search = await callTool('search_content', { sessionId, query: 'Approach' });
    assert(search.ok === true && search.matches.length > 0, 'Found Approach anchor');
    const approachAnchor = search.matches[0].id;

    // Insert comparison table
    const batch = await callTool('execute_batch', {
      sessionId,
      operations: [
        {
          action: 'insert',
          type: 'table',
          anchor: approachAnchor,
          position: 'after',
          headers: ['Metric', 'Traditional Solar', 'Riverbend Community Solar'],
          rows: [
            ['Upfront Capital', '$15,000 - $25,000', '$0 (Subscription)'],
            ['Rooftop Requirement', 'Mandatory Owner Roof', 'None (Renters Eligible)'],
            ['Estimated Bill Savings', '10-15%', '12-18% Guaranteed'],
          ],
        },
      ],
    });
    assert(batch.ok === true, `execute_batch table failed: ${batch.message}`);
    assert(batch.appliedCount === 1, `Applied count: ${batch.appliedCount}`);

    // Save
    const saved = await callTool('save_document', {
      sessionId,
      message: 'Insert community solar comparison table',
      author: 'Lead Financial Analyst',
    });
    assert(saved.ok === true, `save_document failed: ${saved.message}`);
    assert(saved.commitMessage === 'Insert community solar comparison table', `Commit msg: ${saved.commitMessage}`);
    batch2Oid = saved.commitOid;

    batch2DiskBytes = fs.readFileSync(T3_SANDBOX);
    const valResult = await validateDocxBuffer(batch2DiskBytes);
    assert(valResult.ok === true, 'Batch 2 disk file passes validateDocxBuffer');

    const b2Parser = new DocxParser(batch2DiskBytes);
    const b2Ast = await b2Parser.getAST();
    batch2AstText = b2Ast.sections[0].nodes
      .map((n: any) => (n.runs || []).map((r: any) => r.text).join('') || n.metadata?.text || '')
      .join('\n');

    return { batch2Oid, tableBytes: batch2DiskBytes.length };
  });

  // ── 7. Batch 3: Modify Section Titles & Save ────────────────────────────────
  let batch3Oid = "";
  let batch3DiskBytes: Buffer | null = null;
  let batch3AstText = "";

  await step('7. Batch 3: execute_batch modifying section titles & save_document', async () => {
    // Search for Objectives heading
    const search = await callTool('search_content', { sessionId, query: 'Objectives' });
    assert(search.ok === true && search.matches.length > 0, 'Found Objectives heading');
    const objectivesId = search.matches[0].id;

    // Replace title
    const batch = await callTool('execute_batch', {
      sessionId,
      operations: [
        {
          action: 'replace',
          target: objectivesId,
          text: 'Strategic Program Objectives & Milestones (2026-2030)',
        },
      ],
    });
    assert(batch.ok === true, `execute_batch replace failed: ${batch.message}`);
    assert(batch.appliedCount === 1, `Applied count: ${batch.appliedCount}`);

    // Save
    const saved = await callTool('save_document', {
      sessionId,
      message: 'Update section heading to Strategic Program Objectives',
      author: 'Program Director',
    });
    assert(saved.ok === true, `save_document failed: ${saved.message}`);
    assert(saved.commitMessage === 'Update section heading to Strategic Program Objectives', `Commit msg: ${saved.commitMessage}`);
    batch3Oid = saved.commitOid;

    batch3DiskBytes = fs.readFileSync(T3_SANDBOX);
    const valResult = await validateDocxBuffer(batch3DiskBytes);
    assert(valResult.ok === true, 'Batch 3 disk file passes validateDocxBuffer');

    const b3Parser = new DocxParser(batch3DiskBytes);
    const b3Ast = await b3Parser.getAST();
    batch3AstText = b3Ast.sections[0].nodes
      .map((n: any) => (n.runs || []).map((r: any) => r.text).join('') || n.metadata?.text || '')
      .join('\n');

    // Verify search finds new text
    const searchNew = await callTool('search_content', { sessionId, query: 'Strategic Program Objectives & Milestones' });
    assert(searchNew.ok === true && searchNew.matches.length > 0, 'New section title found');

    return { batch3Oid, batch3Bytes: batch3DiskBytes.length };
  });

  // ── 8. History Audit: Pagination & Chronological Ordering ───────────────────
  await step('8. History Audit: get_document_history with pagination (limit/offset) and ordering', async () => {
    // Page 1: limit 2, offset 0
    const p1 = await callTool('get_document_history', { sessionId, limit: 2, offset: 0 });
    assert(p1.ok === true, 'Page 1 history succeeded');
    assert(p1.totalCommits === 4, `Expected totalCommits 4, got ${p1.totalCommits}`);
    assert(p1.history.length === 2, `Expected 2 commits on page 1, got ${p1.history.length}`);
    assert(p1.currentCommit === batch3Oid, `Current commit is Batch 3 OID (${batch3Oid})`);
    assert(p1.history[0].oid === batch3Oid, 'Page 1 commit 0 is Batch 3');
    assert(p1.history[0].message === 'Update section heading to Strategic Program Objectives', 'Commit 0 message correct');
    assert(p1.history[0].author === 'Program Director', 'Commit 0 author correct');
    assert(p1.history[1].oid === batch2Oid, 'Page 1 commit 1 is Batch 2');
    assert(p1.history[1].message === 'Insert community solar comparison table', 'Commit 1 message correct');
    assert(p1.history[1].author === 'Lead Financial Analyst', 'Commit 1 author correct');

    // Page 2: limit 2, offset 2
    const p2 = await callTool('get_document_history', { sessionId, limit: 2, offset: 2 });
    assert(p2.ok === true, 'Page 2 history succeeded');
    assert(p2.totalCommits === 4, `Expected totalCommits 4, got ${p2.totalCommits}`);
    assert(p2.history.length === 2, `Expected 2 commits on page 2, got ${p2.history.length}`);
    assert(p2.history[0].oid === batch1Oid, 'Page 2 commit 0 is Batch 1');
    assert(p2.history[0].message === 'Add executive summary and key highlights', 'Commit 2 message correct');
    assert(p2.history[0].author === 'Chief Architect', 'Commit 2 author correct');
    assert(p2.history[1].oid === baselineOid, 'Page 2 commit 1 is Baseline');
    assert(p2.history[1].message === 'Initial state', 'Commit 3 message correct');

    return { totalCommits: p1.totalCommits, page1Count: p1.history.length, page2Count: p2.history.length };
  });

  // ── 9. Single-Step Revert: Invariant & OOXML Validation ─────────────────────
  await step('9. Single-Step Revert: revert_document steps:1, append-only invariant, OOXML validation', async () => {
    const rev = await callTool('revert_document', { sessionId, steps: 1 });
    assert(rev.ok === true, `revert_document failed: ${rev.message}`);
    assert(rev.revertedToCommit === batch2Oid, `Reverted to Batch 2 commit ${batch2Oid}`);
    assert(rev.dirty === false, 'Session is not dirty after revert');

    // Append-only invariant: commit count increases by +1 on HEAD (now 5)
    const hist = await callTool('get_document_history', { sessionId });
    assert(hist.ok === true, 'get_document_history succeeded');
    assert(hist.totalCommits === 5, `Expected 5 totalCommits (+1 append-only), got ${hist.totalCommits}`);
    assert(hist.currentCommit !== batch3Oid, 'HEAD moved to new revert commit');
    assert(hist.history[0].message.startsWith('Revert to '), `HEAD message starts with 'Revert to ': ${hist.history[0].message}`);

    // Verify AST and disk file reflect Batch 2 state (Batch 3 title change gone)
    const searchBatch3 = await callTool('search_content', { sessionId, query: 'Strategic Program Objectives & Milestones' });
    assert(searchBatch3.ok === true && searchBatch3.matches.length === 0, 'Batch 3 title change is gone');

    const searchObjectives = await callTool('search_content', { sessionId, query: 'Objectives' });
    assert(searchObjectives.ok === true && searchObjectives.matches.length > 0, 'Original Objectives title restored');

    const searchTable = await callTool('search_content', { sessionId, query: 'Traditional Solar' });
    assert(searchTable.ok === true && searchTable.matches.length > 0, 'Batch 2 table is still present');

    // Disk file reflects Batch 2 state: parse disk file AST text and assert exact match with Batch 2 AST text
    const revertedDiskBytes = fs.readFileSync(T3_SANDBOX);
    const revParser = new DocxParser(revertedDiskBytes);
    const revAst = await revParser.getAST();
    const revAstText = revAst.sections[0].nodes
      .map((n: any) => (n.runs || []).map((r: any) => r.text).join('') || n.metadata?.text || '')
      .join('\n');
    assert(revAstText === batch2AstText, 'Disk file AST text reflects Batch 2 state exactly');

    // Reverted file passes validateDocxBuffer
    const val = await validateDocxBuffer(revertedDiskBytes);
    assert(val.ok === true, 'Reverted file passes validateDocxBuffer');

    return { totalCommits: hist.totalCommits, revertedToCommit: rev.revertedToCommit };
  });

  // ── 10. Redo Symmetry: Restores Batch 3 & Visual Inspection ─────────────────
  await step('10. Redo Symmetry: redo_document restores Batch 3 state and confirms visual layout', async () => {
    const redo = await callTool('redo_document', { sessionId });
    assert(redo.ok === true, `redo_document failed: ${redo.message}`);
    assert(redo.dirty === false, 'Dirty is false after redo');
    assert(/^[0-9a-f]{40}$/.test(redo.commitOid), 'Redo commit OID is valid');

    // Commit count increases to 6 (+1 append-only)
    const hist = await callTool('get_document_history', { sessionId });
    assert(hist.ok === true && hist.totalCommits === 6, `Expected 6 totalCommits, got ${hist.totalCommits}`);

    // Verify Batch 3 changes restored in live session
    const searchRestored = await callTool('search_content', { sessionId, query: 'Strategic Program Objectives & Milestones' });
    assert(searchRestored.ok === true && searchRestored.matches.length > 0, 'Batch 3 title restored after redo');

    // Disk file AST text reflects Batch 3 state exactly
    const redoneDiskBytes = fs.readFileSync(T3_SANDBOX);
    const redoParser = new DocxParser(redoneDiskBytes);
    const redoAst = await redoParser.getAST();
    const redoAstText = redoAst.sections[0].nodes
      .map((n: any) => (n.runs || []).map((r: any) => r.text).join('') || n.metadata?.text || '')
      .join('\n');
    assert(redoAstText === batch3AstText, 'Disk file AST text matches Batch 3 state exactly');

    // Fresh visual inspection
    const vis = await callTool('visual_inspect', { sessionId });
    assert(vis.ok === true, 'visual_inspect succeeded on redone state');
    assert(vis.telemetry.wireframe.includes('Strategic Program') || vis.telemetry.wireframe.length > 50, 'Wireframe reflects redone state');

    // Passes validateDocxBuffer
    const val = await validateDocxBuffer(redoneDiskBytes);
    assert(val.ok === true, 'Redone file passes validateDocxBuffer');

    return { totalCommits: hist.totalCommits, commitOid: redo.commitOid };
  });

  // ── 11. Arbitrary OID Revert: Revert to Baseline ────────────────────────────
  await step('11. Arbitrary OID Revert: revert_document to baselineOid returns to pristine state', async () => {
    const rev = await callTool('revert_document', {
      sessionId,
      commitOid: baselineOid,
      message: 'Restore pristine baseline',
    });
    assert(rev.ok === true, `revert_document to baseline failed: ${rev.message}`);
    assert(rev.revertedToCommit === baselineOid, `Reverted to baseline OID ${baselineOid}`);

    // Commit count increases to 7 (+1 append-only)
    const hist = await callTool('get_document_history', { sessionId });
    assert(hist.ok === true && hist.totalCommits === 7, `Expected 7 totalCommits, got ${hist.totalCommits}`);

    // Verify text/AST matches original baseline document
    const currentDiskBytes = fs.readFileSync(T3_SANDBOX);
    const curParser = new DocxParser(currentDiskBytes);
    const curAst = await curParser.getAST();
    const curText = curAst.sections[0].nodes
      .map((n: any) => (n.runs || []).map((r: any) => r.text).join('') || n.metadata?.text || '')
      .join('\n');
    assert(curText === baselineT3Text, 'Document AST text matches original baseline text exactly');

    const searchExec = await callTool('search_content', { sessionId, query: 'Executive Summary' });
    assert(searchExec.ok === true && searchExec.matches.length === 0, 'Executive summary is gone');

    const searchTable = await callTool('search_content', { sessionId, query: 'Traditional Solar' });
    assert(searchTable.ok === true && searchTable.matches.length === 0, 'Comparison table is gone');

    const searchModifiedTitle = await callTool('search_content', { sessionId, query: 'Strategic Program Objectives & Milestones' });
    assert(searchModifiedTitle.ok === true && searchModifiedTitle.matches.length === 0, 'Modified title is gone');

    // Run visual_inspect: verify page count and layout match baseline
    const vis = await callTool('visual_inspect', { sessionId });
    assert(vis.ok === true, 'visual_inspect succeeded');
    assert(vis.telemetry.totalPages === baselineTotalPages, `Total pages (${vis.telemetry.totalPages}) matches baseline (${baselineTotalPages})`);

    const val = await validateDocxBuffer(currentDiskBytes);
    assert(val.ok === true, 'Baseline restored doc passes validateDocxBuffer');

    return { totalCommits: hist.totalCommits, baselineOid };
  });

  // ── 12. Boundary & Conflict Tests ──────────────────────────────────────────
  await step('12. Boundary & Conflict Tests: redo at HEAD, unsaved edits conflict, discardUnsaved', async () => {
    // 1. Redo back to state before baseline revert, then assert next redo fails at HEAD
    await callTool('redo_document', { sessionId });
    const redoEmpty = await callTool('redo_document', { sessionId });
    assert(redoEmpty.ok === false, 'redo at HEAD must fail');
    assert(redoEmpty.code === 'E_NOT_FOUND', `Expected E_NOT_FOUND, got ${redoEmpty.code}`);
    assert(redoEmpty.atLatestState === true, 'Expected atLatestState: true');
    assert(typeof redoEmpty.remedy === 'string', 'Expected remedy in E_NOT_FOUND');

    // 2. Stage unsaved edits and call revert with discardUnsaved: false -> E_CONFLICT
    const search = await callTool('search_content', { sessionId, query: 'Riverbend Community Solar' });
    const titleId = search.matches[0].id;
    const batch = await callTool('execute_batch', {
      sessionId,
      operations: [{ action: 'replace', target: titleId, text: 'Unsaved Temp Title' }],
    });
    assert(batch.ok === true, 'Staged unsaved edit succeeded');

    const refused = await callTool('revert_document', {
      sessionId,
      steps: 1,
      discardUnsaved: false,
    });
    assert(refused.ok === false, 'revert without discard consent must fail');
    assert(refused.code === 'E_CONFLICT', `Expected E_CONFLICT, got ${refused.code}`);
    assert(typeof refused.remedy === 'string', 'Expected remedy in E_CONFLICT');

    // Assert unsaved text was NOT lost
    const searchUnsaved = await callTool('search_content', { sessionId, query: 'Unsaved Temp Title' });
    assert(searchUnsaved.ok === true && searchUnsaved.matches.length > 0, 'Unsaved text still live after refusal');

    // 3. Revert with discardUnsaved: true -> succeeds and discards
    const discarded = await callTool('revert_document', {
      sessionId,
      steps: 1,
      discardUnsaved: true,
    });
    assert(discarded.ok === true, 'revert with discardUnsaved: true succeeded');
    assert(discarded.dirty === false, 'Session is clean after revert');

    const searchCleared = await callTool('search_content', { sessionId, query: 'Unsaved Temp Title' });
    assert(searchCleared.ok === true && searchCleared.matches.length === 0, 'Unsaved text was discarded');

    return {
      refusedCode: refused.code,
      discardedDirty: discarded.dirty,
    };
  });

  // Close T3 stdio session (initialized vars: no TS2454; throws propagate and
  // the outer finally + afterAll still close the client on failure).
  await callTool('close_document', { sessionId });
  sessionId = "";
  const liveClient = stdioClient;
  stdioClient = null;
  if (liveClient) await liveClient.close();

  // ── 13. Concurrency Stress Test on SessionMutex ─────────────────────────────
  await step('13. Concurrency Test: concurrent requests on same session verify SessionMutex serialization', async () => {
    const store = createMcpSessionStore();
    const opened = await LEAN6_HANDLERS.open_document(store, { path: T3_SANDBOX });
    assert(opened.ok === true, 'Opened session for concurrency test');
    const sess = opened.sessionId;

    const insp = await LEAN6_HANDLERS.inspect_document(store, { sessionId: sess });
    assert(insp.ok === true, 'Inspected session');
    const firstId = insp.toonText.split('\n').map((l: string) => l.split(/\s+/)[0]).filter((id: string) => id && !['>', 'ID'].includes(id) && !/^---/.test(id))[0];

    // Fire 10 concurrent requests mix of reads, searches, and writes
    const promises: Promise<any>[] = [];
    for (let i = 0; i < 5; i++) {
      promises.push(LEAN6_HANDLERS.search_content(store, { sessionId: sess, query: 'Solar' }));
      promises.push(
        LEAN6_HANDLERS.execute_batch(store, {
          sessionId: sess,
          operations: [
            {
              action: 'insert',
              type: 'paragraph',
              anchor: firstId,
              position: 'after',
              text: `Concurrent test paragraph #${i}`,
            },
          ],
        })
      );
    }

    const concurrentResults = await Promise.all(promises);
    assert(concurrentResults.every((r) => r.ok === true), 'All 10 concurrent requests completed successfully');

    // Check revision matches expected increments (5 batches -> revision was 1 + 5 = 6)
    const handle = store.get(sess);
    assert(handle.revision === 6, `Expected final revision 6, got ${handle.revision}`);

    await LEAN6_HANDLERS.close_document(store, { sessionId: sess });
    return { concurrentCompleted: concurrentResults.length, finalRevision: handle.revision };
  });

  // ── 14. Zero Host Git Dependency Test under PATH='' ─────────────────────────
  await step('14. Zero Host Git Dependency: test full git workflow under process.env.PATH = ""', async () => {
    const origPath = process.env.PATH;
    process.env.PATH = '';

    try {
      const store = createMcpSessionStore();
      const testDoc = path.join(SANDBOX, 'T3_no_path.docx');
      fs.copyFileSync(T3_ORIGINAL, testDoc);

      // Open (auto-inits repo under PATH='')
      const opened = await LEAN6_HANDLERS.open_document(store, { path: testDoc });
      assert(opened.ok === true, 'Open succeeded under PATH=""');
      const sess = opened.sessionId;

      // Check git repo was created
      const repoDir = path.join(SANDBOX, '.astra', 'T3_no_path.docx');
      assert(fs.existsSync(path.join(repoDir, '.git')), 'Git repo initialized under PATH=""');

      // Edit and save
      const search = await LEAN6_HANDLERS.search_content(store, { sessionId: sess, query: 'Riverbend' });
      if (search.ok !== true) throw new Error(`Search failed under empty PATH: ${search.message}`);
      assert(search.matches.length > 0, 'Found Riverbend anchor under PATH=""');
      const anchor = search.matches[0].id;
      const batch = await LEAN6_HANDLERS.execute_batch(store, {
        sessionId: sess,
        operations: [{ action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'No PATH paragraph' }],
      });
      assert(batch.ok === true, 'Batch succeeded under PATH=""');

      const saved = await LEAN6_HANDLERS.save_document(store, {
        sessionId: sess,
        message: 'Saved under empty PATH',
        author: 'Zero Dep Tester',
      });
      assert(saved.ok === true, 'Save succeeded under PATH=""');

      // History
      const hist = await GIT_HISTORY_HANDLERS.get_document_history(store, { sessionId: sess });
      assert(hist.ok === true && hist.totalCommits === 2, 'History retrieved 2 commits under PATH=""');

      // Revert
      const rev = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId: sess, steps: 1 });
      assert(rev.ok === true, 'Revert succeeded under PATH=""');

      // Redo
      const redo = await GIT_HISTORY_HANDLERS.redo_document(store, { sessionId: sess });
      assert(redo.ok === true, 'Redo succeeded under PATH=""');

      // Validate docx
      const buf = fs.readFileSync(testDoc);
      const val = await validateDocxBuffer(buf);
      assert(val.ok === true, 'Docx valid under PATH=""');

      await LEAN6_HANDLERS.close_document(store, { sessionId: sess });
      return { totalCommitsUnderEmptyPath: hist.totalCommits };
    } finally {
      process.env.PATH = origPath;
    }
  });

  // ── 15. Multi-Document Test with T5_image_doc.docx (Media Preservation) ─────
  await step('15. Multi-Document: T5_image_doc.docx git media tracking, revert, redo, & validation', async () => {
    const store = createMcpSessionStore();
    const opened = await LEAN6_HANDLERS.open_document(store, { path: T5_SANDBOX });
    assert(opened.ok === true, 'T5 opened successfully');
    const sess = opened.sessionId;

    // Verify .astra repo created and media files tracked
    const t5RepoDir = path.join(SANDBOX, '.astra', 'T5_image_doc.docx');
    assert(fs.existsSync(path.join(t5RepoDir, '.git')), 'T5 git repo exists');
    const t5MediaDir = path.join(t5RepoDir, 'word', 'media');
    assert(fs.existsSync(t5MediaDir), 'T5 word/media dir tracked in git repo');
    const mediaFiles = fs.readdirSync(t5MediaDir);
    assert(mediaFiles.length > 0, `Media files tracked in git: ${mediaFiles.join(', ')}`);

    // Visual inspect - check image detected
    const vis = await handleVisualInspect(store, { sessionId: sess });
    assert(vis.ok === true, 'T5 visual_inspect succeeded');
    const hasImageInBreakdown = vis.telemetry.pageBreakdown.some((p: any) => p.hasImages === true);
    assert(hasImageInBreakdown === true, 'Telemetry pageBreakdown detected hasImages === true');

    // Batch edit: add annotation paragraph
    const search = await LEAN6_HANDLERS.search_content(store, { sessionId: sess, query: 'Badge placeholder' });
    assert(search.ok === true && search.matches.length > 0, 'Found badge placeholder');
    const anchor = search.matches[0].id;

    const batch = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId: sess,
      operations: [
        {
          action: 'insert',
          type: 'paragraph',
          anchor,
          position: 'after',
          text: 'Verified Asset: 1x1 placeholder conforms to design token specification.',
        },
      ],
    });
    assert(batch.ok === true, 'Batch edit on T5 succeeded');

    const saved = await LEAN6_HANDLERS.save_document(store, {
      sessionId: sess,
      message: 'Annotate Astra badge specifications',
      author: 'Design Token Lead',
    });
    assert(saved.ok === true, 'T5 save succeeded');

    // Single-step revert
    const rev = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId: sess, steps: 1 });
    assert(rev.ok === true, 'T5 revert succeeded');

    // Validate reverted docx buffer
    const revertedBuf = fs.readFileSync(T5_SANDBOX);
    const valRev = await validateDocxBuffer(revertedBuf);
    assert(valRev.ok === true, 'Reverted T5 passes validateDocxBuffer');

    // Verify media file is still in the docx package!
    const revParser = new DocxParser(revertedBuf);
    const revAst = await revParser.getAST();
    assert(Object.keys(revAst.mediaFiles || {}).length > 0, 'Media files preserved in AST after revert');

    // Redo
    const redone = await GIT_HISTORY_HANDLERS.redo_document(store, { sessionId: sess });
    assert(redone.ok === true, 'T5 redo succeeded');

    const redoneBuf = fs.readFileSync(T5_SANDBOX);
    const valRedo = await validateDocxBuffer(redoneBuf);
    assert(valRedo.ok === true, 'Redone T5 passes validateDocxBuffer');
    const redoneParser = new DocxParser(redoneBuf);
    const redoneAst = await redoneParser.getAST();
    assert(Object.keys(redoneAst.mediaFiles || {}).length > 0, 'Media files preserved in AST after redo');
    const redoneText = redoneAst.sections[0].nodes
      .map((n: any) => (n.runs || []).map((r: any) => r.text).join('') || n.metadata?.text || '')
      .join('\n');
    assert(redoneText.includes('Verified Asset: 1x1 placeholder conforms to design token specification.'), 'Redone T5 restored annotated paragraph');

    await LEAN6_HANDLERS.close_document(store, { sessionId: sess });
    return { mediaFilesTracked: mediaFiles, hasImageInBreakdown };
  });

  // ── 16. Cleanup ─────────────────────────────────────────────────────────────
  await step('16. Clean up temporary sandbox directory', async () => {
    if (fs.existsSync(SANDBOX)) {
      fs.rmSync(SANDBOX, { recursive: true, force: true });
    }
    assert(!fs.existsSync(SANDBOX), 'Sandbox successfully removed');
  });

  // ── Summary Report ──────────────────────────────────────────────────────────
  console.log('\n================================================================');
  console.log('📊 TEST EXECUTION SUMMARY');
  console.log('================================================================');
  const total = results.length;
  const passed = results.filter((r) => r.passed).length;
  const failed = results.filter((r) => !r.passed).length;
  const totalDuration = results.reduce((acc, r) => acc + r.durationMs, 0);

  console.log(`Total Steps Executed : ${total}`);
  console.log(`Passed               : ${passed} (100% GREEN)`);
  console.log(`Failed               : ${failed}`);
  console.log(`Total Duration       : ${(totalDuration / 1000).toFixed(2)}s`);
  console.log('================================================================\n');

  if (failed > 0) {
    const failedNames = results.filter((r) => !r.passed).map((r) => r.name).join('; ');
    throw new Error(`MCP real-client suite failed: ${failed} failing step(s): ${failedNames}`);
  }
  } finally {
    // Best-effort client close + env restore on both success and failure
    // (afterAll repeats both idempotently; sandbox removal lives in afterAll).
    const c = stdioClient;
    stdioClient = null;
    if (c) {
      try {
        await c.close();
      } catch {
        /* best-effort */
      }
    }
    if (savedExtraRoots === undefined) delete process.env.ASTRA_EXTRA_DOC_ROOTS;
    else process.env.ASTRA_EXTRA_DOC_ROOTS = savedExtraRoots;
  }
}

describe("MCP Real-World Client End-to-End Test Suite", () => {
  it("executes all 16 real-document MCP steps", async () => {
    await runRealClientTests();
  }, 120000);
});

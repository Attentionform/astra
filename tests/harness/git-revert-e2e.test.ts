import { describe, it, expect, beforeAll, afterAll, afterEach } from "vitest";
import {
  DocxParser,
  validateDocxBuffer,
  OperationEngine,
  initAstraRepo,
  commitChange,
  getHistory,
  checkoutCommit,
  undoPackage,
  getRepoDir,
  isTrackedPackageFile,
} from '@attentionform/astra-engine';
import AdmZip from "adm-zip";
import fs from "fs";
import os from "os";
import path from "path";
import crypto from "crypto";
import { spawn, execSync, type ChildProcess } from "child_process";
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

const REPO_ROOT = path.resolve(__dirname, "..", "..");
const GALLERY = path.join(REPO_ROOT, "test-gallery");

function sha256Hex(input: string | Buffer): string {
  return crypto.createHash("sha256").update(input).digest("hex");
}

function trackedExtrasOf(parser: DocxParser): Record<string, Buffer> {
  const snap = parser.capturePackageSnapshot();
  const extras: Record<string, Buffer> = {};
  for (const [name, buf] of snap) {
    if (isTrackedPackageFile(name)) extras[name] = buf;
  }
  return extras;
}

function mediaEntriesOf(parser: DocxParser): Map<string, Buffer> {
  const out = new Map<string, Buffer>();
  for (const [name, buf] of parser.capturePackageSnapshot()) {
    if (name.startsWith("word/media/")) out.set(name, Buffer.from(buf));
  }
  return out;
}

const tmpDirs: string[] = [];
function trackDir(dir: string): string {
  tmpDirs.push(dir);
  return dir;
}

afterEach(() => {
  while (tmpDirs.length > 0) {
    const d = tmpDirs.pop() as string;
    try {
      fs.rmSync(d, { recursive: true, force: true });
    } catch {
      /* best-effort */
    }
  }
});

describe("engine git revert end-to-end (real gallery docs)", () => {
  const fixtures = [
    { name: "T5_image_doc.docx", hasTables: false },
    { name: "T4_budget.docx", hasTables: true },
    { name: "T3_proposal.docx", hasTables: true },
  ];

  for (const fixture of fixtures) {
    it(`hash-oracle revert roundtrip for ${fixture.name}`, async () => {
      const dir = trackDir(fs.mkdtempSync(path.join(os.tmpdir(), "astra-git-e2e-")));
      const docPath = path.join(dir, fixture.name);
      fs.copyFileSync(path.join(GALLERY, fixture.name), docPath);

      const rawBuf = fs.readFileSync(docPath);
      const parser = new DocxParser(rawBuf);

      const baselineXml = parser.readDocumentXml();
      const baselineHash = sha256Hex(baselineXml);
      const baselineFiles = [...parser.capturePackageSnapshot().keys()].sort();
      expect(baselineFiles).toContain("word/document.xml");
      const baselineMedia = mediaEntriesOf(parser);
      const baselineExtras = trackedExtrasOf(parser);

      await initAstraRepo(docPath, baselineXml, baselineExtras);

      const ast = await parser.getAST();
      const engine = new OperationEngine(ast, parser);
      const summary = parser.getStructuralSummary();
      const paraTargets = summary.nodes.filter(
        (n) => n.type === "paragraph" && n.text && n.text.length > 0
      );
      expect(paraTargets.length).toBeGreaterThan(0);
      const anchor = paraTargets[0];
      const tableNode = summary.nodes.find((n) => n.type === "table");

      // Commit 1: replace text on the first text-bearing paragraph.
      const r1 = await engine.executeBatch({
        operations: [
          { action: "replace", target: anchor.id, text: "E2E ORACLE EDIT ONE" } as never,
        ],
      });
      expect(r1.success).toBe(true);
      const xmlAfter1 = parser.readDocumentXml();
      const hashAfter1 = sha256Hex(xmlAfter1);
      expect(hashAfter1).not.toBe(baselineHash);
      const oid1 = await commitChange(docPath, xmlAfter1, "harness edit 1", trackedExtrasOf(parser));

      // Commit 2: insert a fresh paragraph after the anchor.
      const r2 = await engine.executeBatch({
        operations: [
          {
            action: "insert",
            type: "paragraph",
            anchor: anchor.id,
            position: "after",
            text: "E2E ORACLE INSERTED PARA",
          } as never,
        ],
      });
      expect(r2.success).toBe(true);
      const xmlAfter2 = parser.readDocumentXml();
      const hashAfter2 = sha256Hex(xmlAfter2);
      expect(hashAfter2).not.toBe(hashAfter1);
      const oid2 = await commitChange(docPath, xmlAfter2, "harness edit 2", trackedExtrasOf(parser));

      // Commit 3: table shading when a table exists, otherwise a second replace.
      if (fixture.hasTables) {
        expect(tableNode).toBeDefined();
        const r3 = await engine.executeBatch({
          operations: [
            {
              action: "table",
              target: tableNode!.id,
              operation: "setCellShading",
              rowIndex: 1,
              colIndex: 1,
              fill: "D9E8F5",
            } as never,
          ],
        });
        expect(r3.success).toBe(true);
      } else {
        const r3 = await engine.executeBatch({
          operations: [
            { action: "replace", target: anchor.id, text: "E2E ORACLE EDIT THREE" } as never,
          ],
        });
        expect(r3.success).toBe(true);
      }
      const xmlAfter3 = parser.readDocumentXml();
      const hashAfter3 = sha256Hex(xmlAfter3);
      expect(hashAfter3).not.toBe(hashAfter2);
      await commitChange(docPath, xmlAfter3, "harness edit 3", trackedExtrasOf(parser));

      // History is newest-first with the three edit messages above the baseline.
      const history = await getHistory(docPath);
      expect(history).toHaveLength(4);
      expect(history.map((c) => c.message.trim())).toEqual([
        "harness edit 3",
        "harness edit 2",
        "harness edit 1",
        "Initial state",
      ]);
      expect(history[1].oid).toBe(oid2);
      expect(history[2].oid).toBe(oid1);

      // Read-only checkout of the middle commit reproduces its bytes exactly.
      const middle = await checkoutCommit(docPath, oid2);
      expect(middle.xml).toBe(xmlAfter2);
      expect(sha256Hex(middle.xml)).toBe(hashAfter2);

      // Checkout is read-only: history depth and working-tree bytes are untouched.
      const historyAfterCheckout = await getHistory(docPath);
      expect(historyAfterCheckout.map((c) => c.oid)).toEqual(history.map((c) => c.oid));
      const worktreeHead = fs.readFileSync(path.join(getRepoDir(docPath), "document.xml"), "utf8");
      expect(worktreeHead).toBe(xmlAfter3);
      expect(sha256Hex(worktreeHead)).toBe(hashAfter3);

      // Media-heavy assertion for the image doc: tracked media round-trips byte-identical.
      if (fixture.name === "T5_image_doc.docx") {
        expect(baselineMedia.size).toBeGreaterThan(0);
        const mediaKeys = [...middle.extra.keys()].filter((k) => k.startsWith("word/media/"));
        expect(mediaKeys.length).toBe(baselineMedia.size);
        for (const [name, buf] of baselineMedia) {
          const restored = middle.extra.get(name);
          expect(restored).toBeDefined();
          expect(Buffer.isBuffer(restored)).toBe(true);
          expect((restored as Buffer).equals(buf)).toBe(true);
        }
      }

      // The restored candidate rezips into a valid package.
      const headBuf = parser.rezip();
      const candidateZip = new AdmZip(headBuf);
      candidateZip.updateFile("word/document.xml", Buffer.from(middle.xml, "utf8"));
      const candidateValidation = await validateDocxBuffer(candidateZip.toBuffer());
      expect(candidateValidation.ok).toBe(true);

      // A single undo from head lands back on the middle bytes.
      const undone = await undoPackage(docPath);
      expect(undone.xml).toBe(xmlAfter2);
      expect(sha256Hex(undone.xml)).toBe(hashAfter2);
      const historyAfterUndo = await getHistory(docPath);
      expect(historyAfterUndo).toHaveLength(3);
      expect(historyAfterUndo[0].message.trim()).toBe("harness edit 2");
      const worktreeAfterUndo = fs.readFileSync(path.join(getRepoDir(docPath), "document.xml"), "utf8");
      expect(worktreeAfterUndo).toBe(xmlAfter2);
    }, 30_000);
  }
});

type RpcMessage = {
  jsonrpc: string;
  id?: number;
  method?: string;
  result?: any;
  error?: any;
};

class McpStdioClient {
  private proc: ChildProcess;
  private nextId = 1;
  private pending = new Map<number, (msg: RpcMessage) => void>();
  private buffer = "";
  public stderrText = "";

  constructor(proc: ChildProcess) {
    this.proc = proc;
    proc.stdout!.on("data", (d: Buffer) => this.onData(d.toString()));
    proc.stderr!.on("data", (d: Buffer) => {
      this.stderrText += d.toString();
    });
  }

  private onData(chunk: string): void {
    this.buffer += chunk;
    let idx: number;
    while ((idx = this.buffer.indexOf("\n")) >= 0) {
      const line = this.buffer.slice(0, idx).trim();
      this.buffer = this.buffer.slice(idx + 1);
      if (!line) continue;
      try {
        const msg = JSON.parse(line) as RpcMessage;
        if (msg.id !== undefined && this.pending.has(msg.id)) {
          const resolve = this.pending.get(msg.id)!;
          this.pending.delete(msg.id);
          resolve(msg);
        }
      } catch {
        /* ignore non-JSON stdio noise */
      }
    }
  }

  request(method: string, params: unknown, timeoutMs = 25_000): Promise<RpcMessage> {
    const id = this.nextId++;
    return new Promise<RpcMessage>((resolve, reject) => {
      this.pending.set(id, resolve);
      try {
        this.proc.stdin!.write(`${JSON.stringify({ jsonrpc: "2.0", id, method, params })}\n`);
      } catch (err) {
        this.pending.delete(id);
        reject(err);
        return;
      }
      setTimeout(() => {
        if (this.pending.has(id)) {
          this.pending.delete(id);
          reject(new Error(`mcp request timed out: ${method}`));
        }
      }, timeoutMs);
    });
  }

  notify(method: string, params?: unknown): void {
    this.proc.stdin!.write(`${JSON.stringify({ jsonrpc: "2.0", method, params: params ?? {} })}\n`);
  }

  async callTool(name: string, args: Record<string, unknown>): Promise<{ envelope: any; isError: boolean }> {
    const res = await this.request("tools/call", { name, arguments: args });
    if (res.error) throw new Error(`mcp transport error: ${JSON.stringify(res.error).slice(0, 300)}`);
    const content = res.result?.content?.[0]?.text as string;
    expect(typeof content).toBe("string");
    return { envelope: JSON.parse(content), isError: res.result?.isError === true };
  }
}

describe("mcp git revert end-to-end (live stdio server)", () => {
  let serverProc: ChildProcess | null = null;
  let client: McpStdioClient | null = null;
  let allowedDir = "";

  beforeAll(async () => {
    allowedDir = trackDir(fs.mkdtempSync(path.join(os.tmpdir(), "astra-mcp-e2e-")));
    fs.copyFileSync(path.join(GALLERY, "T5_image_doc.docx"), path.join(allowedDir, "image.docx"));
    fs.copyFileSync(path.join(GALLERY, "T4_budget.docx"), path.join(allowedDir, "budget.docx"));

    const distEntry = path.join(REPO_ROOT, "packages", "mcp-server", "dist", "index.js");
    const srcEntry = path.join(REPO_ROOT, "packages", "mcp-server", "src", "index.ts");
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

    serverProc = spawn("node", [distEntry, "--allowed-dir", allowedDir], {
      stdio: ["pipe", "pipe", "pipe"],
    });
    client = new McpStdioClient(serverProc);
    const init = await client.request("initialize", {
      protocolVersion: "2024-11-05",
      capabilities: {},
      clientInfo: { name: "git-revert-harness", version: "0.0.0" },
    });
    expect(init.result?.serverInfo?.name).toBe("astra-mcp-server");
    client.notify("notifications/initialized");
    const listed = await client.request("tools/list", {});
    const toolNames = (listed.result?.tools ?? []).map((t: any) => t.name);
    for (const expected of ["open_document", "execute_batch", "save_document", "get_document_history", "revert_document", "redo_document"]) {
      expect(toolNames).toContain(expected);
    }
  }, 30_000);

  afterAll(async () => {
    if (serverProc) {
      try {
        serverProc.kill("SIGTERM");
        await new Promise<void>((resolve) => {
          const timer = setTimeout(resolve, 5000);
          serverProc!.once("exit", () => {
            clearTimeout(timer);
            resolve();
          });
        });
      } catch {
        /* best-effort */
      }
      serverProc = null;
      client = null;
    }
  });

  function firstRowId(toonText: string): string {
    const ids = toonText
      .split("\n")
      .map((l) => l.split(/\s+/)[0])
      .filter((id) => !!id && !["|", ">", "ID"].includes(id) && !/^---/.test(id));
    expect(ids.length).toBeGreaterThan(0);
    return ids[0];
  }

  function documentXmlOfDocx(docxPath: string): string {
    const zip = new AdmZip(fs.readFileSync(docxPath));
    const entry = zip.getEntry("word/document.xml");
    expect(entry).not.toBeNull();
    return entry!.getData().toString("utf8");
  }

  it("revert and redo restore exact saved bytes with hash-oracle checks", async () => {
    expect(client).not.toBeNull();
    const c = client!;
    const docPath = path.join(allowedDir, "image.docx");

    const opened = await c.callTool("open_document", { path: docPath });
    expect(opened.envelope.ok).toBe(true);
    const sessionId = opened.envelope.sessionId as string;

    const insp1 = await c.callTool("inspect_document", { sessionId, limit: 10 });
    expect(insp1.envelope.ok).toBe(true);
    const edit1 = await c.callTool("execute_batch", {
      sessionId,
      operations: [{ action: "replace", target: firstRowId(insp1.envelope.toonText), text: "MCP E2E EDIT ONE" }],
    });
    expect(edit1.envelope.ok).toBe(true);
    const save1 = await c.callTool("save_document", { sessionId, message: "mcp e2e first save" });
    expect(save1.envelope.ok).toBe(true);

    const insp2 = await c.callTool("inspect_document", { sessionId, limit: 10 });
    expect(insp2.envelope.ok).toBe(true);
    const edit2 = await c.callTool("execute_batch", {
      sessionId,
      operations: [{ action: "replace", target: firstRowId(insp2.envelope.toonText), text: "MCP E2E EDIT TWO" }],
    });
    expect(edit2.envelope.ok).toBe(true);
    const save2 = await c.callTool("save_document", { sessionId, message: "mcp e2e second save" });
    expect(save2.envelope.ok).toBe(true);

    const hist = await c.callTool("get_document_history", { sessionId });
    expect(hist.envelope.ok).toBe(true);
    expect(hist.envelope.totalCommits).toBe(3);
    expect(hist.envelope.history[0].message).toBe("mcp e2e second save");
    expect(hist.envelope.history[1].message).toBe("mcp e2e first save");
    expect(hist.envelope.history[2].message).toBe("Initial state");
    const oidFirst = hist.envelope.history[1].oid as string;
    const oidSecond = hist.envelope.history[0].oid as string;

    // Expected bytes come from the on-disk sidecar, independent of the server session.
    const expectedFirst = (await checkoutCommit(docPath, oidFirst)).xml;
    const expectedSecond = (await checkoutCommit(docPath, oidSecond)).xml;
    expect(sha256Hex(expectedFirst)).not.toBe(sha256Hex(expectedSecond));

    const reverted = await c.callTool("revert_document", { sessionId, steps: 1 });
    expect(reverted.envelope.ok).toBe(true);
    expect(reverted.envelope.revertedToCommit).toBe(oidFirst);
    expect(sha256Hex(documentXmlOfDocx(docPath))).toBe(sha256Hex(expectedFirst));
    const inspAfterRevert = await c.callTool("inspect_document", { sessionId, limit: 10 });
    expect(inspAfterRevert.envelope.ok).toBe(true);
    expect(inspAfterRevert.envelope.toonText).toContain("MCP E2E EDIT ONE");

    const redone = await c.callTool("redo_document", { sessionId });
    expect(redone.envelope.ok).toBe(true);
    expect(sha256Hex(documentXmlOfDocx(docPath))).toBe(sha256Hex(expectedSecond));

    // Malformed OID fails closed with an E_VALIDATION envelope.
    const malformed = await c.callTool("revert_document", { sessionId, commitOid: "not-an-oid" });
    expect(malformed.envelope.ok).toBe(false);
    expect(malformed.envelope.code).toBe("E_VALIDATION");

    // Unsaved edits plus discardUnsaved:false fails closed with E_CONFLICT.
    const insp3 = await c.callTool("inspect_document", { sessionId, limit: 10 });
    expect(insp3.envelope.ok).toBe(true);
    const dirty = await c.callTool("execute_batch", {
      sessionId,
      operations: [{ action: "replace", target: firstRowId(insp3.envelope.toonText), text: "MCP E2E DIRTY PROBE" }],
    });
    expect(dirty.envelope.ok).toBe(true);
    const conflict = await c.callTool("revert_document", { sessionId, steps: 1, discardUnsaved: false });
    expect(conflict.envelope.ok).toBe(false);
    expect(conflict.envelope.code).toBe("E_CONFLICT");

    const closed = await c.callTool("close_document", { sessionId });
    expect(closed.envelope.ok).toBe(true);
  }, 30_000);
});

import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { OperationSchema, validateDocxBuffer } from '@attentionform/astra-engine';

const REPO_ROOT = path.resolve(__dirname, '../..');
const CONFIG_PATH = path.join(REPO_ROOT, 'integrations/opencode/opencode.json');
const SKILL_PATH = path.join(REPO_ROOT, 'integrations/opencode/agent/doc-editor.md');
const COFFEE_DOC_PATH = path.join(REPO_ROOT, 'test-gallery/T1_coffee_onepager.docx');
const MCP_SERVER_DIST = path.join(REPO_ROOT, 'packages/mcp-server/dist/index.js');

const tmpFiles: string[] = [];

afterAll(() => {
  for (const f of tmpFiles) {
    try {
      if (fs.existsSync(f)) {
        fs.unlinkSync(f);
      }
    } catch {
      /* best-effort */
    }
  }
});

describe('OpenCode Integration Suite', () => {
  describe('1. Schema Test: opencode.json', () => {
    it('verifies integrations/opencode/opencode.json is valid JSON with $schema and mcp server dictionary', () => {
      expect(fs.existsSync(CONFIG_PATH)).toBe(true);
      const raw = fs.readFileSync(CONFIG_PATH, 'utf-8');
      const config = JSON.parse(raw);

      // Schema definition
      expect(config.$schema).toBe('https://opencode.ai/config.json');
      expect(config.mcp).toBeDefined();
      expect(typeof config.mcp).toBe('object');

      // Registered servers: exactly 2 local workspace runtimes (astra-production removed)
      expect(config.mcp.astra).toBeDefined();
      expect(config.mcp['astra-dev']).toBeDefined();
      expect(config.mcp['astra-production']).toBeUndefined();
      expect(Object.keys(config.mcp).sort()).toEqual(['astra', 'astra-dev']);

      // Main active server configuration (local build)
      const astra = config.mcp.astra;
      expect(astra.type).toBe('local');
      expect(Array.isArray(astra.command)).toBe(true);
      expect(astra.enabled).toBe(true);
      expect(astra.command).toContain('node');
      expect(astra.command.some((c: string) => c.includes('dist/index.js'))).toBe(true);
      expect(astra.command).toContain('--allowed-dir');
      expect(astra.environment).toBeDefined();
      expect(astra.environment.ASTRA_ALLOWED_DIR).toBeUndefined();
      const allowedDirIndex = astra.command.indexOf('--allowed-dir');
      expect(allowedDirIndex).toBeGreaterThan(-1);
      const allowedDirValue = astra.command[allowedDirIndex + 1];
      expect(typeof allowedDirValue).toBe('string');
      expect(path.isAbsolute(allowedDirValue)).toBe(true);

      // Development configuration (direct tsx)
      const astraDev = config.mcp['astra-dev'];
      expect(astraDev.type).toBe('local');
      expect(Array.isArray(astraDev.command)).toBe(true);
      expect(astraDev.command).toContain('npx');
      expect(astraDev.command).toContain('tsx');
      expect(astraDev.command.some((c: string) => c.includes('src/index.ts'))).toBe(true);
      expect(astraDev.command).toContain('--allowed-dir');
      expect(astraDev.environment).toBeDefined();
      expect(astraDev.environment.ASTRA_ALLOWED_DIR).toBeUndefined();
      const allowedDirDevIndex = astraDev.command.indexOf('--allowed-dir');
      expect(allowedDirDevIndex).toBeGreaterThan(-1);
      const allowedDirDevValue = astraDev.command[allowedDirDevIndex + 1];
      expect(typeof allowedDirDevValue).toBe('string');
      expect(path.isAbsolute(allowedDirDevValue)).toBe(true);
    });
  });

  describe('2. Skill Test: doc-editor.md', () => {
    it('verifies integrations/opencode/agent/doc-editor.md has metadata frontmatter, tool directives, prompt law, and workflow', () => {
      expect(fs.existsSync(SKILL_PATH)).toBe(true);
      const content = fs.readFileSync(SKILL_PATH, 'utf-8');

      // Frontmatter metadata
      expect(content).toMatch(/^---\r?\n/);
      expect(content).toContain('name: doc-editor');
      expect(content).toContain('description:');
      expect(content).toContain('model:');
      expect(content).toContain('tools:');

      // All 14 tools in frontmatter
      const requiredTools = [
        'open_document',
        'inspect_document',
        'search_content',
        'execute_batch',
        'save_document',
        'close_document',
        'new_document',
        'compose_document',
        'export_document',
        'visual_inspect',
        'get_document_history',
        'preview_revert',
        'revert_document',
        'redo_document',
      ];
      for (const tool of requiredTools) {
        expect(content).toContain(tool);
      }

      // 14-tool header reference & UUID-slice session ID format
      expect(content).toContain('The 14 Tools');
      expect(content).not.toContain('Lean-7');
      expect(content).not.toContain('lean-7');
      expect(content).toContain('sess_<uuid>');
      expect(content).not.toContain('randomUUID().slice(0, 8)');
      expect(content).not.toContain('sess_<nanoid>');

      // Canonical engine operations verification against @attentionform/astra-engine
      const canonicalOpsDocumented = [
        'insert',
        'delete',
        'replace',
        'style',
        'callout',
        'table',
        'chart',
        'trackedChange',
        'footnote',
        'pageBreak',
        'field',
      ];
      for (const op of canonicalOpsDocumented) {
        expect(content).toContain(`\`${op}\`:`);
      }

      const engineActionSet = new Set(
        OperationSchema.options.map((schema: any) => schema.shape.action.value)
      );

      // Strict enforcement: zero invented operations exist in doc-editor.md or the engine
      const inventedOperations = [
        'insert_paragraph',
        'insert_heading',
        'insert_table',
        'insert_section',
        'format_run',
        'format_paragraph',
        'insert_chart',
        'insert_callout',
        'tracked_change',
      ];
      for (const op of inventedOperations) {
        expect(content).not.toContain(`\`${op}\``);
        expect(content).not.toContain(`'${op}'`);
        expect(content).not.toContain(`"${op}"`);
        expect(engineActionSet.has(op)).toBe(false);
      }

      // Extract all documented operation example objects from doc-editor.md
      const documentedExamples: any[] = [];
      const inlineObjRegex = /`({[^`]*\baction\b[^`]*})`/g;
      let m: RegExpExecArray | null;
      while ((m = inlineObjRegex.exec(content)) !== null) {
        const raw = m[1];
        const normalized = raw.replace(/([a-zA-Z0-9_$]+)\?:/g, '$1:');
        const parsed = new Function(`return (${normalized});`)();
        documentedExamples.push(parsed);
      }

      // Extract from fenced code blocks (```json or ```) containing action operations
      const fencedBlockRegex = /```(?:json|json5)?\s*([\s\S]*?)\s*```/g;
      while ((m = fencedBlockRegex.exec(content)) !== null) {
        const block = m[1].trim();
        if (block.includes('action') && block.startsWith('{') && block.endsWith('}')) {
          try {
            const normalized = block.replace(/([a-zA-Z0-9_$]+)\?:/g, '$1:');
            const parsed = new Function(`return (${normalized});`)();
            documentedExamples.push(parsed);
          } catch {
            /* best-effort */
          }
        }
      }

      // Ensure all 11 operations are covered by documented examples
      for (const op of canonicalOpsDocumented) {
        const found = documentedExamples.some((ex: any) => {
          if (op === 'pageBreak') {
            return ex.action === 'insert' && ex.type === 'pageBreak';
          }
          return ex.action === op;
        });
        expect(found, `Missing documented example for operation: ${op}`).toBe(true);
      }

      // No carve-outs! All 11 operations must pass OperationSchema.safeParse cleanly!
      for (const example of documentedExamples) {
        const result = OperationSchema.safeParse(example);
        expect(
          result.success,
          `Failed to parse example: ${JSON.stringify(example)} with errors: ${JSON.stringify((result as any).error?.issues)}`
        ).toBe(true);
      }

      // Prompt law requirements
      const contentLower = content.toLowerCase();
      expect(contentLower).toContain('image is approximate');
      expect(content).toContain('XML is truth');
      expect(contentLower).toContain('large defects only');
      expect(contentLower).toContain('never auto-edit from the image alone');

      // Planted defect detection coverage
      expect(contentLower).toContain('missing-chart');
      expect(contentLower).toContain('overflow-table');
      expect(contentLower).toContain('blank-page');
      expect(contentLower).toContain('orphan-heading');

      // Document mutation and TOON fidelity rules
      expect(content).toContain('compact');
      expect(content).toContain('standard');
      expect(content).toContain('full');
      expect(contentLower).toContain('sectionid');

      // DrawingML charts and Tracked Changes
      expect(contentLower).toContain('drawingml');
      expect(content).toContain('w:ins');
      expect(content).toContain('w:del');
      expect(content).toContain('w:delText');

      // Best practice workflow: Inspect -> Search -> Edit -> Validate -> Export
      expect(content).toMatch(/Inspect.*Search.*Edit.*Validate.*Export/is);

      // Data-astra-id sanitization directive
      expect(content).toContain('data-astra-id');
    });
  });

  describe('3. Stdio E2E Bridge Test (Simulating OpenCode Agent)', () => {
    let client: Client;
    let transport: StdioClientTransport;

    beforeAll(async () => {
      expect(fs.existsSync(MCP_SERVER_DIST)).toBe(true);

      // Spawn real @attentionform/astra-mcp-server over stdio with allowed-dir pointing to REPO_ROOT
      transport = new StdioClientTransport({
        command: process.execPath,
        args: [MCP_SERVER_DIST, '--allowed-dir', REPO_ROOT],
        env: {
          ...process.env,
          ASTRA_EXTRA_DOC_ROOTS: `${REPO_ROOT}${path.delimiter}${os.tmpdir()}`,
        },
      });

      client = new Client(
        {
          name: 'opencode-integration-tester',
          version: '1.0.0',
        },
        {
          capabilities: {},
        }
      );

      await client.connect(transport);
    }, 20000);

    afterAll(async () => {
      try {
        await client?.close();
      } catch {
        /* best-effort */
      }
    });

    it('executes full turnkey workflow: listTools -> open -> inspect -> search -> execute_batch -> export -> save -> close', async () => {
      // 1. Calls tools/list and asserts all 19 tools are exposed
      const toolsResult = await client.listTools();
      const toolNames = toolsResult.tools.map((t) => t.name).sort();
      expect(toolNames).toEqual([
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
      ]);
      expect(toolsResult.tools.length).toBe(19);

      // 2. Calls open_document on test-gallery/T1_coffee_onepager.docx
      expect(fs.existsSync(COFFEE_DOC_PATH)).toBe(true);
      const openRes = await client.callTool({
        name: 'open_document',
        arguments: { path: COFFEE_DOC_PATH },
      });
      expect(openRes.isError).toBeFalsy();
      const openData = JSON.parse((openRes.content as Array<{ type: string; text: string }>)[0].text);
      expect(openData.ok).toBe(true);
      const sessionId: string = openData.sessionId;
      expect(typeof sessionId).toBe('string');
      expect(sessionId.length).toBeGreaterThan(0);

      // 3. Calls inspect_document and checks TOON outline
      const inspectRes = await client.callTool({
        name: 'inspect_document',
        arguments: { sessionId, fidelity: 'standard' },
      });
      expect(inspectRes.isError).toBeFalsy();
      const inspectData = JSON.parse((inspectRes.content as Array<{ type: string; text: string }>)[0].text);
      expect(inspectData.ok).toBe(true);
      expect(typeof inspectData.toonText).toBe('string');
      expect(inspectData.toonText.length).toBeGreaterThan(0);
      expect(inspectData.rowCount).toBeGreaterThan(0);
      expect(inspectData.palette).toBeDefined();
      expect(Array.isArray(inspectData.palette.availableStyles)).toBe(true);

      // 4. Calls search_content for "Coffee"
      const searchRes = await client.callTool({
        name: 'search_content',
        arguments: { sessionId, query: 'Coffee' },
      });
      expect(searchRes.isError).toBeFalsy();
      const searchData = JSON.parse((searchRes.content as Array<{ type: string; text: string }>)[0].text);
      expect(searchData.ok).toBe(true);
      expect(Array.isArray(searchData.matches)).toBe(true);
      expect(searchData.matches.length).toBeGreaterThan(0);
      const firstMatch = searchData.matches[0];
      expect(firstMatch.id).toBeTruthy();
      expect(firstMatch.preview.toLowerCase()).toContain('coffee');

      // 5. Calls execute_batch to insert a section or paragraph
      const insertionText = 'OpenCode Turnkey Integration Verified: High-Fidelity OOXML Synthesis.';
      const batchRes = await client.callTool({
        name: 'execute_batch',
        arguments: {
          sessionId,
          operations: [
            {
              action: 'insert',
              type: 'paragraph',
              anchor: firstMatch.id,
              position: 'after',
              text: insertionText,
            },
          ],
        },
      });
      expect(batchRes.isError).toBeFalsy();
      const batchData = JSON.parse((batchRes.content as Array<{ type: string; text: string }>)[0].text);
      expect(batchData.ok).toBe(true);
      expect(batchData.appliedCount).toBe(1);

      // 6. Calls export_document with format: 'markdown' and format: 'html'
      const exportMdRes = await client.callTool({
        name: 'export_document',
        arguments: { sessionId, format: 'markdown' },
      });
      expect(exportMdRes.isError).toBeFalsy();
      const mdData = JSON.parse((exportMdRes.content as Array<{ type: string; text: string }>)[0].text);
      expect(mdData.ok).toBe(true);
      expect(mdData.format).toBe('markdown');
      expect(typeof mdData.content).toBe('string');
      expect(mdData.content).toContain(insertionText);
      expect(mdData.content).not.toContain('data-astra-id');

      const exportHtmlRes = await client.callTool({
        name: 'export_document',
        arguments: { sessionId, format: 'html' },
      });
      expect(exportHtmlRes.isError).toBeFalsy();
      const htmlData = JSON.parse((exportHtmlRes.content as Array<{ type: string; text: string }>)[0].text);
      expect(htmlData.ok).toBe(true);
      expect(htmlData.format).toBe('html');
      expect(typeof htmlData.content).toBe('string');
      expect(htmlData.content).toContain(insertionText);
      expect(htmlData.content).not.toContain('data-astra-id');

      // 7. Calls save_document to a temp directory and validates OOXML with validateDocxBuffer
      const tempOutDocx = path.join(
        os.tmpdir(),
        `opencode-e2e-saved-${process.pid}-${Date.now()}.docx`
      );
      tmpFiles.push(tempOutDocx);

      const saveRes = await client.callTool({
        name: 'save_document',
        arguments: { sessionId, outPath: tempOutDocx },
      });
      expect(saveRes.isError).toBeFalsy();
      const saveData = JSON.parse((saveRes.content as Array<{ type: string; text: string }>)[0].text);
      expect(saveData.ok).toBe(true);
      expect(fs.existsSync(tempOutDocx)).toBe(true);

      const savedBuffer = fs.readFileSync(tempOutDocx);
      const validation = await validateDocxBuffer(savedBuffer);
      expect(validation.ok).toBe(true);

      // Zero leaked data-astra-id attributes in saved package
      const zip = new AdmZip(savedBuffer);
      for (const entry of zip.getEntries()) {
        if (entry.entryName.endsWith('.xml') || entry.entryName.endsWith('.rels')) {
          const xml = entry.getData().toString('utf-8');
          expect(xml).not.toContain('data-astra-id');
        }
      }

      // 8. Calls close_document
      const closeRes = await client.callTool({
        name: 'close_document',
        arguments: { sessionId },
      });
      expect(closeRes.isError).toBeFalsy();
      const closeData = JSON.parse((closeRes.content as Array<{ type: string; text: string }>)[0].text);
      expect(closeData.ok).toBe(true);
      expect(closeData.sessionId).toBe(sessionId);
    }, 30000);
  });

  describe('4. Security & Sandbox Boundary Rejection', () => {
    let client: Client;
    let transport: StdioClientTransport;
    let validSessionId: string;

    beforeAll(async () => {
      expect(fs.existsSync(MCP_SERVER_DIST)).toBe(true);

      // Spawn real @attentionform/astra-mcp-server over stdio with allowed-dir pointing to REPO_ROOT
      transport = new StdioClientTransport({
        command: process.execPath,
        args: [MCP_SERVER_DIST, '--allowed-dir', REPO_ROOT],
        env: {
          ...process.env,
          ASTRA_EXTRA_DOC_ROOTS: `${REPO_ROOT}${path.delimiter}${os.tmpdir()}`,
        },
      });

      client = new Client(
        {
          name: 'opencode-security-tester',
          version: '1.0.0',
        },
        {
          capabilities: {},
        }
      );

      await client.connect(transport);

      // Open valid document to get an active session for save/export negative tests
      const openRes = await client.callTool({
        name: 'open_document',
        arguments: { path: COFFEE_DOC_PATH },
      });
      const openData = JSON.parse((openRes.content as Array<{ type: string; text: string }>)[0].text);
      expect(openData.ok).toBe(true);
      validSessionId = openData.sessionId;
    }, 20000);

    afterAll(async () => {
      try {
        if (validSessionId) {
          await client?.callTool({
            name: 'close_document',
            arguments: { sessionId: validSessionId },
          });
        }
      } catch {
        /* best-effort */
      }
      try {
        await client?.close();
      } catch {
        /* best-effort */
      }
    });

    it('rejects open_document with traversal path (../../outside.docx, /etc/passwd) with E_SANDBOX', async () => {
      const resTraversal = await client.callTool({
        name: 'open_document',
        arguments: { path: '../../outside.docx' },
      });
      expect(resTraversal.isError).toBe(true);
      const textTraversal = (resTraversal.content as Array<{ type: string; text: string }>)[0].text;
      expect(textTraversal).toContain('E_SANDBOX');

      const resPasswd = await client.callTool({
        name: 'open_document',
        arguments: { path: '/etc/passwd' },
      });
      expect(resPasswd.isError).toBe(true);
      const textPasswd = (resPasswd.content as Array<{ type: string; text: string }>)[0].text;
      expect(textPasswd).toContain('E_SANDBOX');
    });

    it('rejects save_document with traversal path outside allowed directory (/etc/evil.docx, ../../escape.docx) with E_SANDBOX', async () => {
      const resOutside = await client.callTool({
        name: 'save_document',
        arguments: { sessionId: validSessionId, outPath: '/etc/evil.docx' },
      });
      expect(resOutside.isError).toBe(true);
      const textOutside = (resOutside.content as Array<{ type: string; text: string }>)[0].text;
      expect(textOutside).toContain('E_SANDBOX');

      const resTraversal = await client.callTool({
        name: 'save_document',
        arguments: { sessionId: validSessionId, outPath: '../../escape.docx' },
      });
      expect(resTraversal.isError).toBe(true);
      const textTraversal = (resTraversal.content as Array<{ type: string; text: string }>)[0].text;
      expect(textTraversal).toContain('E_SANDBOX');
    });

    it('rejects export_document with path traversal with E_SANDBOX', async () => {
      const resTraversal = await client.callTool({
        name: 'export_document',
        arguments: {
          sessionId: validSessionId,
          format: 'markdown',
          outPath: '../../escape.md',
        },
      });
      expect(resTraversal.isError).toBe(true);
      const textTraversal = (resTraversal.content as Array<{ type: string; text: string }>)[0].text;
      expect(textTraversal).toContain('E_SANDBOX');
    });

    it('rejects tool calls with null byte (/tmp/foo\\0.docx) with E_SANDBOX', async () => {
      const resNullOpen = await client.callTool({
        name: 'open_document',
        arguments: { path: '/tmp/foo\0.docx' },
      });
      expect(resNullOpen.isError).toBe(true);
      const textNullOpen = (resNullOpen.content as Array<{ type: string; text: string }>)[0].text;
      expect(textNullOpen).toContain('E_SANDBOX');

      const resNullSave = await client.callTool({
        name: 'save_document',
        arguments: { sessionId: validSessionId, outPath: '/tmp/foo\0.docx' },
      });
      expect(resNullSave.isError).toBe(true);
      const textNullSave = (resNullSave.content as Array<{ type: string; text: string }>)[0].text;
      expect(textNullSave).toContain('E_SANDBOX');
    });
  });
});

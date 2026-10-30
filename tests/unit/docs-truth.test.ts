import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { getServedToolNames } from '@attentionform/astra-mcp-server';
import { OperationSchema } from '@attentionform/astra-engine';
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
  ValidateDocumentSchema,
  GetStatsSchema,
  SanitizeDocumentSchema,
  NormalizeDocumentSchema,
  ImportDocumentSchema,
} from '../../packages/mcp-server/src/lib/mcp/schemas';
import { NewDocumentSchema } from '../../packages/mcp-server/src/tools/new-document';
import { ComposeDocumentSchema } from '../../packages/mcp-server/src/tools/compose-document';
import { ExportDocumentSchema } from '../../packages/mcp-server/src/tools/export';
import { VisualInspectSchema } from '../../packages/mcp-server/src/tools/visual';

const REPO_ROOT = path.resolve(__dirname, '..', '..');
const CLI_README = path.join(REPO_ROOT, 'packages', 'cli', 'README.md');
const CLI_RUN = path.join(REPO_ROOT, 'packages', 'cli', 'src', 'cli', 'run.ts');
const MCP_README = path.join(REPO_ROOT, 'packages', 'mcp-server', 'README.md');
const ROOT_README = path.join(REPO_ROOT, 'README.md');
const ENGINE_README = path.join(REPO_ROOT, 'packages', 'engine', 'README.md');
const GLAMA_JSON = path.join(REPO_ROOT, 'registries', 'glama.json');
const SMITHERY_JSON = path.join(REPO_ROOT, 'registries', 'smithery.json');
const GALLERY_DIR = path.join(REPO_ROOT, 'test-gallery');

const HELP_ALIASES = new Set(['help', 'version', '--help', '-h', '--version', '-v']);

function readText(p: string): string {
  return fs.readFileSync(p, 'utf8');
}

function liveCliCommands(): string[] {
  const src = readText(CLI_RUN);
  const found = new Set<string>();
  const re = /case '([A-Za-z0-9_-]+)'/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(src)) !== null) {
    const name = m[1];
    if (HELP_ALIASES.has(name)) continue;
    if (!/^[a-z]+$/.test(name)) continue;
    found.add(name);
  }
  return [...found].sort();
}

function docCliCommands(): string[] {
  const md = readText(CLI_README);
  const found = new Set<string>();
  const re = /`astra ([a-z0-9_-]+)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(md)) !== null) {
    const name = m[1];
    if (HELP_ALIASES.has(name)) continue;
    if (name === 'command') continue;
    found.add(name);
  }
  return [...found].sort();
}

function docCliHeadings(): string[] {
  const md = readText(CLI_README);
  const found: string[] = [];
  const re = /^### \d+\. .*`astra ([a-z0-9_-]+)`/gm;
  let m: RegExpExecArray | null;
  while ((m = re.exec(md)) !== null) found.push(m[1]);
  return [...new Set(found)].sort();
}

function readmeToolSections(md: string): Map<string, string> {
  const out = new Map<string, string>();
  type Head = { name: string; idx: number; isTool: boolean };
  const heads: Head[] = [];
  const re = /^(#{2,4}) (`[A-Za-z_]+`|[^`\n]+)\s*$/gm;
  let m: RegExpExecArray | null;
  while ((m = re.exec(md)) !== null) {
    const hashes = m[1];
    const title = m[2].trim();
    const toolMatch = hashes === '###' && /^`[a-z_]+`$/.test(title);
    heads.push({
      name: toolMatch ? title.slice(1, -1) : title,
      idx: m.index,
      isTool: toolMatch,
    });
  }
  for (let i = 0; i < heads.length; i++) {
    if (!heads[i].isTool) continue;
    const end = i + 1 < heads.length ? heads[i + 1].idx : md.length;
    out.set(heads[i].name, md.slice(heads[i].idx, end));
  }
  return out;
}

function paramsInSection(section: string): string[] {
  const found = new Set<string>();
  for (const line of section.split('\n')) {
    const t = line.trim();
    if (!t.startsWith('|')) continue;
    const cell = t.split('|').slice(1, -1)[0] ?? '';
    const cm = /`([A-Za-z_][A-Za-z0-9_]*)`/.exec(cell);
    if (cm) found.add(cm[1]);
  }
  return [...found].sort();
}

function schemaKeys(schema: unknown): string[] {
  const s = schema as {
    shape?: Record<string, unknown>;
    _def?: { shape?: Record<string, unknown> | (() => Record<string, unknown>) };
  };
  let shape: Record<string, unknown> = {};
  if (s && typeof s === 'object') {
    if (s.shape && typeof s.shape === 'object') shape = s.shape;
    else if (typeof s._def?.shape === 'function') shape = s._def.shape();
    else if (s._def?.shape && typeof s._def.shape === 'object') shape = s._def.shape;
  }
  return Object.keys(shape).sort();
}

function opCount(): number {
  const s = OperationSchema as unknown as {
    options?: unknown[];
    _def?: { options?: unknown[] };
  };
  if (Array.isArray(s.options)) return s.options.length;
  if (Array.isArray(s._def?.options)) return (s._def?.options as unknown[]).length;
  throw new Error('cannot read OperationSchema options length');
}

function galleryCount(): number {
  return fs.readdirSync(GALLERY_DIR).filter((f) => f.endsWith('.docx')).length;
}

const TOOL_SCHEMAS: Record<string, unknown> = {
  open_document: OpenDocumentSchema,
  inspect_document: InspectDocumentSchema,
  search_content: SearchContentSchema,
  execute_batch: ExecuteBatchSchema,
  save_document: SaveDocumentSchema,
  close_document: CloseDocumentSchema,
  new_document: NewDocumentSchema,
  compose_document: ComposeDocumentSchema,
  export_document: ExportDocumentSchema,
  visual_inspect: VisualInspectSchema,
  get_document_history: GetDocumentHistorySchema,
  preview_revert: PreviewRevertSchema,
  revert_document: RevertDocumentSchema,
  redo_document: RedoDocumentSchema,
  validate_document: ValidateDocumentSchema,
  get_stats: GetStatsSchema,
  sanitize_document: SanitizeDocumentSchema,
  normalize_document: NormalizeDocumentSchema,
  import_document: ImportDocumentSchema,
};

describe('docs truth: cli commands match dispatch', () => {
  it('every README command exists in run.ts and every dispatch target is documented', () => {
    const live = liveCliCommands();
    const doc = docCliCommands();
    const headings = docCliHeadings();
    expect(live.length).toBeGreaterThan(0);
    expect(doc.length).toBeGreaterThan(0);
    const missing = live.filter((c) => !doc.includes(c));
    const extra = doc.filter((c) => !live.includes(c));
    expect({ missing, extra, live, doc }).toEqual({
      missing: [],
      extra: [],
      live,
      doc: live,
    });
    expect(headings).toEqual(live);
  });
});

describe('docs truth: mcp tools and params match live', () => {
  it('README headings plus both registries equal getServedToolNames', () => {
    const live = [...getServedToolNames()].sort();
    expect(live.length).toBeGreaterThan(0);
    const md = readText(MCP_README);
    const sections = readmeToolSections(md);
    const readmeTools = [...sections.keys()]
      .filter((n) => Object.prototype.hasOwnProperty.call(TOOL_SCHEMAS, n))
      .sort();
    expect(readmeTools).toEqual(live);
    const glama = JSON.parse(readText(GLAMA_JSON)) as { tools?: Array<{ name?: string }> };
    const smithery = JSON.parse(readText(SMITHERY_JSON)) as {
      capabilities?: { tools?: Array<{ name?: string }> };
    };
    const glamaNames = (glama.tools ?? []).map((t) => String(t.name)).sort();
    const smitheryNames = (smithery.capabilities?.tools ?? []).map((t) => String(t.name)).sort();
    expect(glamaNames).toEqual(live);
    expect(smitheryNames).toEqual(live);
  });

  it('README param tables match live zod schema keys per tool', () => {
    const md = readText(MCP_README);
    const sections = readmeToolSections(md);
    for (const [tool, schema] of Object.entries(TOOL_SCHEMAS)) {
      const section = sections.get(tool);
      expect(section, `missing README section for ${tool}`).toBeDefined();
      const docParams = paramsInSection(section as string);
      const liveParams = schemaKeys(schema);
      expect({ tool, docParams, liveParams }).toEqual({ tool, docParams: liveParams, liveParams });
    }
  });
});

describe('docs truth: counts match live', () => {
  it('root README tool/op/gallery numbers match live counts', () => {
    const liveTools = getServedToolNames().length;
    const liveOps = opCount();
    const liveGallery = galleryCount();
    const liveCli = liveCliCommands().length;
    expect(liveTools).toBeGreaterThan(0);
    expect(liveOps).toBeGreaterThan(0);
    expect(liveGallery).toBeGreaterThan(0);

    const root = readText(ROOT_README);
    const toolDash: number[] = [];
    {
      const re = /(\d+)-tool\b/g;
      let m: RegExpExecArray | null;
      while ((m = re.exec(root)) !== null) toolDash.push(Number(m[1]));
    }
    expect(toolDash.length).toBeGreaterThan(0);
    for (const n of toolDash) expect(n).toBe(liveTools);

    const toolsWord: number[] = [];
    {
      const re = /(\d+)\s+tools\b/gi;
      let m: RegExpExecArray | null;
      while ((m = re.exec(root)) !== null) {
        const before = root.slice(Math.max(0, m.index - 40), m.index);
        if (/token/i.test(before)) continue;
        toolsWord.push(Number(m[1]));
      }
    }
    expect(toolsWord.length).toBeGreaterThan(0);
    for (const n of toolsWord) expect(n).toBe(liveTools);

    const opsNums: number[] = [];
    {
      const re = /(\d+)\s+(?:fine-grained document |Atomic |Transactional )?operations\b/gi;
      let m: RegExpExecArray | null;
      while ((m = re.exec(root)) !== null) opsNums.push(Number(m[1]));
    }
    expect(opsNums.length).toBeGreaterThan(0);
    for (const n of opsNums) expect(n).toBe(liveOps);

    const badge = /ECMA-376 Gallery:\s*(\d+)\/(\d+)/.exec(root);
    expect(badge).not.toBeNull();
    expect(Number((badge as RegExpExecArray)[1])).toBe(liveGallery);
    expect(Number((badge as RegExpExecArray)[2])).toBe(liveGallery);

    const allTpl = /All (\d+) templates/.exec(root);
    expect(allTpl).not.toBeNull();
    expect(Number((allTpl as RegExpExecArray)[1])).toBe(liveGallery);

    const engine = readText(ENGINE_README);
    const engOps = /(\d+)\s+Atomic Operations/i.exec(engine);
    expect(engOps).not.toBeNull();
    expect(Number((engOps as RegExpExecArray)[1])).toBe(liveOps);

    const cliHeadings = docCliHeadings();
    expect(cliHeadings.length).toBe(liveCli);

    const mcp = readText(MCP_README);
    const mcpAll = /returns all (\d+)/i.exec(mcp);
    expect(mcpAll).not.toBeNull();
    expect(Number((mcpAll as RegExpExecArray)[1])).toBe(liveTools);
  });
});

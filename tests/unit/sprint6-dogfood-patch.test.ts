import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import {
  addComment as addCommentBackend,
  addComment as addCommentEngine,
  createBlankDocxBuffer,
  type DocumentAST,
  DocxParser,
  type DocxZipAdapter,
  type DocxZipAdapter as EngineZipAdapter,
  executeCallout as executeCalloutBackend,
  executeCallout as executeCalloutEngine,
  executeImage as executeImageBackend,
  executeImage as executeImageEngine,
  OperationEngine,
  sanitizeDocumentXml as sanitizeBackend,
  sanitizeDocumentXml as sanitizeEngine,
} from '@attentionform/astra-engine';

import { run } from '../../packages/cli/src/cli';

// ── helpers ─────────────────────────────────────────────────────────────────

function makeMapAdapter(seed: Map<string, string>): DocxZipAdapter {
  const store = new Map<string, string>(seed);
  return {
    hasFile: (p: string): boolean => store.has(p),
    readText: (p: string): string => {
      const v = store.get(p);
      if (v === undefined) throw new Error(`missing ${p}`);
      return v;
    },
    readTextOrNull: (p: string): string | null => store.get(p) ?? null,
    writeText: (p: string, content: string): void => {
      store.set(p, content);
    },
    writeBuffer: (p: string, content: Buffer): void => {
      store.set(p, content.toString('utf8'));
    },
  };
}

function makeEngineMapAdapter(seed: Map<string, string>): EngineZipAdapter {
  const store = new Map<string, string>(seed);
  return {
    hasFile: (p: string): boolean => store.has(p),
    readText: (p: string): string => {
      const v = store.get(p);
      if (v === undefined) throw new Error(`missing ${p}`);
      return v;
    },
    readTextOrNull: (p: string): string | null => store.get(p) ?? null,
    writeText: (p: string, content: string): void => {
      store.set(p, content);
    },
    writeBuffer: (p: string, content: Buffer): void => {
      store.set(p, content.toString('utf8'));
    },
  };
}

function makeParaDoc(): { doc: Document; p: Element } {
  const doc = new DOMParser().parseFromString(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body></w:document>',
    'text/xml'
  );
  const p = doc.getElementsByTagName('w:p')[0] as unknown as Element;
  return { doc, p };
}

function collectParaIds(commentsXml: string): string[] {
  const out: string[] = [];
  const re = /w14:paraId="([^"]+)"/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(commentsXml)) !== null) out.push(m[1] as string);
  const re15 = /w15:paraId="([^"]+)"/g;
  while ((m = re15.exec(commentsXml)) !== null) out.push(m[1] as string);
  return out;
}

function createMinimalDocx(): Buffer {
  const zip = new AdmZip();
  zip.addFile(
    '[Content_Types].xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
    )
  );
  zip.addFile(
    'word/_rels/document.xml.rels',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`
    )
  );
  zip.addFile(
    'word/document.xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p data-astra-id="p_001"><w:r><w:t>anchor</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
    )
  );
  return zip.toBuffer();
}

function mockImageAST(): DocumentAST {
  const doc = new DOMParser().parseFromString(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>',
    'text/xml'
  );
  const bodyP = doc.getElementsByTagName('w:p')[0] as unknown as Element;
  return {
    metadata: {},
    sections: [
      {
        id: 'sec_root',
        heading: null,
        nodes: [
          {
            id: 'p_001',
            type: 'paragraph',
            xml: bodyP,
            metadata: {
              text: 'x',
              hasCitation: false,
              hasComment: false,
              hasBookmark: false,
              hasTrackedChanges: false,
              hasEquation: false,
              hasField: false,
              hasImage: false,
              hasHyperlink: false,
            },
            version: 1,
            checksum: 'c1',
          },
        ],
        pageSetup: {},
        version: 1,
        checksum: 's1',
      },
    ],
    styles: {},
    numbering: { abstractNum: {}, concreteNum: {} },
    relationships: {},
    headers: {},
    footers: {},
    mediaFiles: {},
    settings: { trackChanges: false, defaultAuthor: 'ASTra' },
    nextChangeId: 1,
  } as unknown as DocumentAST;
}

// ── F1: paraId range fuzz ────────────────────────────────────────────────────

describe('sprint6-dogfood F1: comment paraId <0x80000000', () => {
  it('backend addComment fuzzes 100 paraIds in range (8-hex, high bit clear)', () => {
    const adapter = makeMapAdapter(new Map<string, string>());
    const seen = new Set<string>();
    for (let i = 0; i < 100; i++) {
      const { p } = makeParaDoc();
      addCommentBackend(adapter, p, 'Fuzz', `c${i}`);
    }
    const xml = adapter.readText('word/comments.xml');
    const ids = collectParaIds(xml);
    expect(ids.length).toBeGreaterThanOrEqual(100);
    for (const id of ids) {
      expect(id).toMatch(/^[0-9A-F]{8}$/);
      const v = parseInt(id, 16);
      expect(Number.isFinite(v)).toBe(true);
      expect(v).toBeGreaterThan(0);
      expect(v).toBeLessThan(0x80000000);
      seen.add(id);
    }
    // Randomness sanity: 100 draws must not all collide.
    expect(seen.size).toBeGreaterThan(90);
  });

  it('engine addComment mirrors the mask (50 draws in range)', () => {
    const adapter = makeEngineMapAdapter(new Map<string, string>());
    for (let i = 0; i < 50; i++) {
      const { p } = makeParaDoc();
      addCommentEngine(adapter as unknown as Parameters<typeof addCommentEngine>[0], p, 'Fuzz', `e${i}`);
    }
    const xml = adapter.readText('word/comments.xml');
    const ids = collectParaIds(xml);
    expect(ids.length).toBeGreaterThanOrEqual(50);
    for (const id of ids) {
      const v = parseInt(id, 16);
      expect(v).toBeGreaterThan(0);
      expect(v).toBeLessThan(0x80000000);
    }
  });
});

// ── F2: callout tblCellMar order + validate ok:true all themes ───────────────

describe('sprint6-dogfood F2: callout tblCellMar top,left,bottom,right + validate', () => {
  const themes = ['indigo', 'emerald', 'amber', 'slate', 'rose'] as const;

  it.each(themes)('backend callout theme %s emits schema order and validates ok:true', async (theme) => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const anchor = ast.sections[0]?.nodes[0]?.id;
    if (!anchor) throw new Error('blank template has no anchor node');
    const engine = new OperationEngine(ast, parser);
    const res = await engine.executeBatch({
      operations: [{ action: 'callout', anchor, title: `T ${theme}`, text: `body ${theme}`, theme } as unknown as Parameters<OperationEngine['executeBatch']>[0]['operations'][number]],
    });
    expect(res.success).toBe(true);
    const node = ast.sections[0].nodes.find((n) => (n.metadata as { callout?: { title?: string } })?.callout?.title === `T ${theme}`);
    expect(node).toBeDefined();
    const xml = new XMLSerializer().serializeToString(node!.xml);
    const mar = xml.slice(xml.indexOf('<w:tblCellMar'), xml.indexOf('</w:tblCellMar>'));
    const itop = mar.indexOf('<w:top');
    const ileft = mar.indexOf('<w:left');
    const ibottom = mar.indexOf('<w:bottom');
    const iright = mar.indexOf('<w:right');
    expect(itop).toBeGreaterThanOrEqual(0);
    expect(ileft).toBeGreaterThanOrEqual(0);
    expect(ibottom).toBeGreaterThanOrEqual(0);
    expect(iright).toBeGreaterThanOrEqual(0);
    expect(itop).toBeLessThan(ileft);
    expect(ileft).toBeLessThan(ibottom);
    expect(ibottom).toBeLessThan(iright);
    const validation = await parser.validateOOXML();
    expect(validation.ok).toBe(true);
  });

  it('engine executeCallout mirrors schema order', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const anchor = ast.sections[0]?.nodes[0]?.id;
    if (!anchor) throw new Error('blank template has no anchor node');
    await executeCalloutBackend(
      { action: 'callout', anchor, title: 'B', text: 'backend', theme: 'amber' },
      ast,
      parser
    );
    await executeCalloutEngine(
      { action: 'callout', anchor, title: 'E', text: 'engine', theme: 'amber' } as unknown as Parameters<typeof executeCalloutEngine>[0],
      ast as unknown as Parameters<typeof executeCalloutEngine>[1],
      parser as unknown as Parameters<typeof executeCalloutEngine>[2]
    );
    const ser = (x: Element): string => new XMLSerializer().serializeToString(x);
    for (const n of ast.sections[0].nodes.filter((x) => (x.metadata as { callout?: unknown })?.callout !== undefined)) {
      const xml = ser(n.xml);
      const mar = xml.slice(xml.indexOf('<w:tblCellMar'), xml.indexOf('</w:tblCellMar>'));
      expect(mar.indexOf('<w:top')).toBeLessThan(mar.indexOf('<w:left'));
      expect(mar.indexOf('<w:left')).toBeLessThan(mar.indexOf('<w:bottom'));
      expect(mar.indexOf('<w:bottom')).toBeLessThan(mar.indexOf('<w:right'));
    }
  });
});

// ── F3: people.xml prefix preserved ──────────────────────────────────────────

describe('sprint6-dogfood F3: sanitizer preserves w:/w15: author prefix', () => {
  const peopleXml =
    `<w15:people xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">` +
    `<w15:person w15:author="Jane Doe"><w15:presenceInfo/></w15:person>` +
    `<w:ins w:author="Legacy Author" w:initials="LA"/>` +
    `</w15:people>`;

  it('backend keeps w15:author (never rewrites to w:author)', () => {
    const r = sanitizeBackend(peopleXml, { replacementAuthor: 'ASTra' });
    // w15:author + w:author + w:initials = 3 purged author slots.
    expect(r.purgedAuthors).toBe(3);
    expect(r.xml).toContain('w15:author="ASTra"');
    expect(r.xml).toContain('w:author="ASTra"');
    expect(r.xml).not.toContain('Jane Doe');
    expect(r.xml).not.toContain('Legacy Author');
    // Exactly one w15:author and one w:author — no prefix collapse.
    expect((r.xml.match(/w15:author="/g) ?? []).length).toBe(1);
    expect((r.xml.match(/ w:author="/g) ?? []).length).toBe(1);
  });

  it('engine mirrors the prefix-preserving callback', () => {
    const r = sanitizeEngine(peopleXml, { replacementAuthor: 'ASTra' });
    expect(r.purgedAuthors).toBe(3);
    expect(r.xml).toContain('w15:author="ASTra"');
    expect(r.xml).toContain('w:author="ASTra"');
    expect(r.xml).not.toContain('Jane Doe');
  });
});

// ── F5: image.delete missing target throws ───────────────────────────────────

describe('sprint6-dogfood F5: image.delete missing target throws like tables', () => {
  it('backend delete ghost throws with remedy (no silent success)', () => {
    const ast = mockImageAST();
    expect(() =>
      executeImageBackend({ action: 'image', operation: 'delete', target: 'ghost_img_999' }, ast)
    ).toThrow(/Target image node 'ghost_img_999' not found/);
    expect(() =>
      executeImageBackend({ action: 'image', operation: 'delete', target: 'ghost_img_999' }, ast)
    ).toThrow(/Remedy:/);
    // No mutation on miss.
    expect(ast.sections[0].nodes.length).toBe(1);
  });

  it('engine delete ghost throws with remedy', () => {
    const ast = mockImageAST();
    expect(() =>
      executeImageEngine(
        { action: 'image', operation: 'delete', target: 'ghost_img_999' },
        ast as unknown as Parameters<typeof executeImageEngine>[1]
      )
    ).toThrow(/Target image node 'ghost_img_999' not found/);
  });
});

// ── CLI --json codes ─────────────────────────────────────────────────────────

describe('sprint6-dogfood CLI: batch errors carry E_VALIDATION/E_QUOTA in human + --json', () => {
  const tmp: string[] = [];
  let logSpy: ReturnType<typeof vi.spyOn>;
  let errSpy: ReturnType<typeof vi.spyOn>;
  let logs: string[];
  let errs: string[];

  beforeEach(() => {
    logs = [];
    errs = [];
    logSpy = vi.spyOn(console, 'log').mockImplementation((...a: unknown[]) => {
      logs.push(a.map(String).join(' '));
    });
    errSpy = vi.spyOn(console, 'error').mockImplementation((...a: unknown[]) => {
      errs.push(a.map(String).join(' '));
    });
  });

  afterEach(() => {
    logSpy.mockRestore();
    errSpy.mockRestore();
    for (const f of tmp.splice(0)) {
      try {
        fs.unlinkSync(f);
      } catch {
        /* best-effort */
      }
    }
  });

  function writeTmpDocx(): string {
    const buf = new DocxParser(createMinimalDocx()).rezip();
    const file = path.join(os.tmpdir(), `s6-patch-${process.pid}-${Date.now()}-${tmp.length}.docx`);
    fs.writeFileSync(file, buf);
    tmp.push(file);
    return file;
  }

  function writeOps(payload: unknown): string {
    const file = path.join(os.tmpdir(), `s6-ops-${process.pid}-${Date.now()}-${tmp.length}.json`);
    fs.writeFileSync(file, JSON.stringify(payload));
    tmp.push(file);
    return file;
  }

  function lastJson(): Record<string, unknown> | null {
    for (let i = logs.length - 1; i >= 0; i--) {
      try {
        return JSON.parse(logs[i] as string) as Record<string, unknown>;
      } catch {
        /* not json */
      }
    }
    return null;
  }

  it('invalid op --json carries E_VALIDATION in json.code + human stderr', async () => {
    const file = writeTmpDocx();
    const ops = writeOps({ operations: [{ action: 'definitely_not_a_real_op' }] });
    const code = await run(['--json', 'batch', file, ops]);
    expect(code).toBe(1);
    const payload = lastJson();
    expect(payload).not.toBeNull();
    expect(payload?.code).toBe('E_VALIDATION');
    expect(String(payload?.error ?? '')).toMatch(/E_VALIDATION/);
    expect(errs.join('\n')).toMatch(/E_VALIDATION/);
  });

  it('cap overflow (501 ops) --json carries E_QUOTA in json.code + human stderr', async () => {
    const file = writeTmpDocx();
    const many = Array.from({ length: 501 }, (_, i) => ({ action: 'replace', target: 'p_001', text: `x${i}` }));
    const ops = writeOps({ operations: many, sequential: true });
    const code = await run(['--json', 'batch', file, ops]);
    expect(code).toBe(1);
    const payload = lastJson();
    expect(payload).not.toBeNull();
    expect(payload?.code).toBe('E_QUOTA');
    expect(String(payload?.error ?? '')).toMatch(/E_QUOTA/);
    expect(errs.join('\n')).toMatch(/E_QUOTA/);
  });
});

import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import path from 'path';
import fs from 'fs';
import fsp from 'fs/promises';
import os from 'os';
import { DOMParser } from '@xmldom/xmldom';
import {
  ensureAstraGitRepo,
  getRepoDir,
  initAstraRepo,
  DocxParser,
  DocumentAST,
  convertAstToHtml,
  convertAstToMarkdown,
  parseFontSize,
  replaceOpSchema,
  styleOpSchema,
  executeStyle,
  executeReplace,
  replaceText,
} from '../../packages/engine/src';
import { renderTable } from '../../packages/engine/src/export/html';
import { renderTableMarkdown } from '../../packages/engine/src/export/markdown';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';
import { LEAN6_HANDLERS } from '../../packages/mcp-server/src/lib/mcp-tool-adapter';
import { stripAnsi, shouldStripColor, formatCliOutput } from '../../packages/cli/src/cli';

function blankDocxBuf(): Buffer {
  return DocxParser.createNewDocument('blank').rezip();
}

function parseXml(xmlStr: string): Element {
  const doc = new DOMParser().parseFromString(xmlStr, 'text/xml');
  return doc.documentElement;
}

function buildTestAst(nodes: any[]): DocumentAST {
  return {
    metadata: {},
    settings: { trackChanges: false, defaultAuthor: 'Author' },
    nextChangeId: 1,
    styles: {},
    numbering: { abstractNum: {}, concreteNum: {} },
    relationships: {},
    headers: {},
    footers: {},
    mediaFiles: {},
    sections: [
      {
        id: 'sec_1',
        heading: null,
        pageSetup: null,
        nodes,
        version: 1,
        checksum: 'sec_chk_1',
      },
    ],
  } as unknown as DocumentAST;
}

describe('Wave 2E: Appendix Lows Sweep (T-4, L-4, L-3, L-8, L-13, T-13)', () => {
  // =========================================================================
  // T-4: .astra/ POSIX Permissions (packages/engine/src/git.ts)
  // =========================================================================
  describe('T-4: .astra/ POSIX 0o700 Directory Permissions', () => {
    let tmpDir: string;
    let docPath: string;

    beforeEach(async () => {
      tmpDir = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-t4-test-'));
      docPath = path.join(tmpDir, 'document.docx');
    });

    afterEach(async () => {
      await fsp.rm(tmpDir, { recursive: true, force: true });
    });

    it('ensureAstraGitRepo sets directory mode to 0o700 on POSIX', async () => {
      const repoDir = await ensureAstraGitRepo(docPath);
      expect(fs.existsSync(repoDir)).toBe(true);

      if (process.platform !== 'win32') {
        const stat = await fsp.stat(repoDir);
        const mode = stat.mode & 0o777;
        expect(mode).toBe(0o700);

        const astraDir = path.dirname(repoDir);
        const astraStat = await fsp.stat(astraDir);
        expect(astraStat.mode & 0o777).toBe(0o700);
      }
    });

    it('initAstraRepo creates repo directory with 0o700 permissions on POSIX', async () => {
      const initialXml = '<w:document><w:body><w:p>Initial</w:p></w:body></w:document>';
      await initAstraRepo(docPath, initialXml);
      const repoDir = getRepoDir(docPath);

      if (process.platform !== 'win32') {
        const stat = await fsp.stat(repoDir);
        expect(stat.mode & 0o777).toBe(0o700);
      }
    });
  });

  // =========================================================================
  // L-4: MCP search_content Field Harmonization (matches & results)
  // =========================================================================
  describe('L-4: MCP search_content Field Harmonization', () => {
    it('returns both matches and results: matches for literal query', async () => {
      const store = createMcpSessionStore();
      const h = await store.openBuffer(blankDocxBuf(), 'search-test.docx');
      const pId = h.ast.sections[0].nodes[0]?.id;

      // Insert search target
      if (pId) {
        await LEAN6_HANDLERS.execute_batch(store, {
          sessionId: h.sessionId,
          operations: [
            {
              action: 'insert',
              type: 'paragraph',
              anchor: pId,
              position: 'after',
              text: 'Harmonized MCP search target paragraph',
            },
          ],
        });
      }

      const res = await LEAN6_HANDLERS.search_content(store, {
        sessionId: h.sessionId,
        query: 'Harmonized',
      });

      expect(res.ok).toBe(true);
      if (res.ok) {
        expect(res.matches).toBeDefined();
        expect(res.results).toBeDefined();
        expect(res.matches.length).toBeGreaterThanOrEqual(1);
        expect(res.results).toEqual(res.matches);
        expect(res.matches[0].preview).toContain('Harmonized');
      }
    });

    it('returns both matches and results: matches for regex query', async () => {
      const store = createMcpSessionStore();
      const h = await store.openBuffer(blankDocxBuf(), 'search-regex.docx');
      const pId = h.ast.sections[0].nodes[0]?.id;

      if (pId) {
        await LEAN6_HANDLERS.execute_batch(store, {
          sessionId: h.sessionId,
          operations: [
            {
              action: 'insert',
              type: 'paragraph',
              anchor: pId,
              position: 'after',
              text: 'AlphaBeta123PatternFound',
            },
          ],
        });
      }

      const res = await LEAN6_HANDLERS.search_content(store, {
        sessionId: h.sessionId,
        query: 'AlphaBeta\\d+Pattern',
        regex: true,
      });

      expect(res.ok).toBe(true);
      if (res.ok) {
        expect(res.matches).toBeDefined();
        expect(res.results).toBeDefined();
        expect(res.matches.length).toBe(1);
        expect(res.results).toEqual(res.matches);
        expect(res.matches[0].preview).toContain('AlphaBeta123PatternFound');
      }
    });

    it('returns empty matches and results arrays on no match', async () => {
      const store = createMcpSessionStore();
      const h = await store.openBuffer(blankDocxBuf(), 'search-none.docx');

      const res = await LEAN6_HANDLERS.search_content(store, {
        sessionId: h.sessionId,
        query: 'NonExistentString12345XYZ',
      });

      expect(res.ok).toBe(true);
      if (res.ok) {
        expect(res.matches).toEqual([]);
        expect(res.results).toEqual([]);
      }
    });
  });

  // =========================================================================
  // L-3: Honor Table captionPosition: 'below' in Exporters
  // =========================================================================
  describe("L-3: Honor Table captionPosition: 'below' in Exporters", () => {
    const tableNodeBelow = {
      id: 'tbl_1',
      type: 'table',
      caption: 'Table 1: Financial Performance',
      metadata: {
        captionPosition: 'below',
      },
      rows: [
        {
          id: 'row_0',
          cells: [
            { id: 'c_0', text: 'Metric', metadata: {} },
            { id: 'c_1', text: 'Value', metadata: {} },
          ],
        },
        {
          id: 'row_1',
          cells: [
            { id: 'c_2', text: 'Revenue', metadata: {} },
            { id: 'c_3', text: '$1.2M', metadata: {} },
          ],
        },
      ],
    };

    const tableNodeAbove = {
      id: 'tbl_2',
      type: 'table',
      caption: 'Table 2: Overview Summary',
      metadata: {
        captionPosition: 'above',
      },
      rows: [
        {
          id: 'row_0',
          cells: [
            { id: 'c_0', text: 'Item', metadata: {} },
            { id: 'c_1', text: 'Qty', metadata: {} },
          ],
        },
      ],
    };

    it('HTML renderTable places caption below </table> when captionPosition is below', () => {
      const html = renderTable(tableNodeBelow);
      expect(html).toContain('class="table-caption"');
      const tableEndIdx = html.indexOf('</table>');
      const captionIdx = html.indexOf('<p class="table-caption">Table 1: Financial Performance</p>');

      expect(tableEndIdx).toBeGreaterThan(-1);
      expect(captionIdx).toBeGreaterThan(-1);
      expect(captionIdx).toBeGreaterThan(tableEndIdx);
    });

    it('HTML renderTable places caption above <table ...> when captionPosition is above or default', () => {
      const html = renderTable(tableNodeAbove);
      const tableStartIdx = html.indexOf('<table');
      const captionIdx = html.indexOf('<p class="table-caption">Table 2: Overview Summary</p>');

      expect(tableStartIdx).toBeGreaterThan(-1);
      expect(captionIdx).toBeGreaterThan(-1);
      expect(captionIdx).toBeLessThan(tableStartIdx);
    });

    it('HTML convertAstToHtml honors captionPosition below', () => {
      const ast = buildTestAst([tableNodeBelow]);
      const fullHtml = convertAstToHtml(ast);
      const tableEndIdx = fullHtml.indexOf('</table>');
      const captionIdx = fullHtml.indexOf('Table 1: Financial Performance');

      expect(captionIdx).toBeGreaterThan(tableEndIdx);
    });

    it('Markdown renderTableMarkdown places caption after table when captionPosition is below', () => {
      const md = renderTableMarkdown(tableNodeBelow);
      const tableRowIdx = md.indexOf('| Metric | Value |');
      const captionIdx = md.indexOf('Table 1: Financial Performance');

      expect(tableRowIdx).toBeGreaterThan(-1);
      expect(captionIdx).toBeGreaterThan(-1);
      expect(captionIdx).toBeGreaterThan(tableRowIdx);
    });

    it('Markdown convertAstToMarkdown places caption before table when captionPosition is above', () => {
      const ast = buildTestAst([tableNodeAbove]);
      const md = convertAstToMarkdown(ast);
      const tableRowIdx = md.indexOf('| Item | Qty |');
      const captionIdx = md.indexOf('Table 2: Overview Summary');

      expect(tableRowIdx).toBeGreaterThan(-1);
      expect(captionIdx).toBeGreaterThan(-1);
      expect(captionIdx).toBeLessThan(tableRowIdx);
    });

    it('Markdown convertAstToMarkdown places caption after table when captionPosition is below', () => {
      const ast = buildTestAst([tableNodeBelow]);
      const md = convertAstToMarkdown(ast);
      const tableRowIdx = md.indexOf('| Metric | Value |');
      const captionIdx = md.indexOf('Table 1: Financial Performance');

      expect(tableRowIdx).toBeGreaterThan(-1);
      expect(captionIdx).toBeGreaterThan(-1);
      expect(captionIdx).toBeGreaterThan(tableRowIdx);
    });
  });

  // =========================================================================
  // L-8: Accept "pt" Suffix for Font Size sz
  // =========================================================================
  describe('L-8: Accept "pt" Suffix for Font Size sz / fontSize', () => {
    it('parseFontSize parses "12pt" -> 24 half-points and "10.5pt" -> 21 half-points', () => {
      expect(parseFontSize('12pt')).toBe(24);
      expect(parseFontSize('12PT')).toBe(24);
      expect(parseFontSize('10.5pt')).toBe(21);
      expect(parseFontSize('10.5 pt')).toBe(21);
      expect(parseFontSize('8.25pt')).toBe(17); // Math.round(8.25 * 2) = 17
      expect(parseFontSize('0pt')).toBe(0);
    });

    it('parseFontSize passes numeric half-points through unchanged', () => {
      expect(parseFontSize(24)).toBe(24);
      expect(parseFontSize(21)).toBe(21);
      expect(parseFontSize(0)).toBe(0);
      expect(parseFontSize(undefined)).toBeUndefined();
    });

    it('parseFontSize rejects invalid string formats gracefully', () => {
      expect(parseFontSize('invalid')).toBeUndefined();
      expect(parseFontSize('12px')).toBeUndefined();
      expect(parseFontSize('pt12')).toBeUndefined();
      expect(parseFontSize('')).toBeUndefined();
      expect(parseFontSize('   ')).toBeUndefined();
    });

    it('parseFontSize rejects bare numeric strings like the schema gate', () => {
      for (const bare of ['+5', '1e3', '.5', '24', '-5', '12', '0']) {
        expect(parseFontSize(bare), bare).toBeUndefined();
      }
      expect(parseFontSize('+5pt')).toBeUndefined();
      expect(parseFontSize(0)).toBe(0);
      expect(parseFontSize('0pt')).toBe(0);
    });

    it('operation schemas accept string font size with pt suffix', () => {
      const styleParsed = styleOpSchema.safeParse({
        action: 'style',
        target: 'p_1',
        sz: '12pt',
      });
      expect(styleParsed.success).toBe(true);

      const styleParsedFontSize = styleOpSchema.safeParse({
        action: 'style',
        target: 'p_1',
        fontSize: '10.5pt',
      });
      expect(styleParsedFontSize.success).toBe(true);

      const replaceParsed = replaceOpSchema.safeParse({
        action: 'replace',
        target: 'p_1',
        text: 'Replaced text',
        sz: '14pt',
      });
      expect(replaceParsed.success).toBe(true);

      const replaceParsedFontSize = replaceOpSchema.safeParse({
        action: 'replace',
        target: 'p_1',
        text: 'Replaced text',
        fontSize: '12pt',
      });
      expect(replaceParsedFontSize.success).toBe(true);
      // Bare numeric strings fail the schema gate, matching runtime.
      expect(styleOpSchema.safeParse({ action: 'style', target: 'p_1', sz: '24' }).success).toBe(false);
      expect(replaceOpSchema.safeParse({ action: 'replace', target: 'p_1', text: 'x', fontSize: '24' }).success).toBe(false);
    });

    it('executeStyle applies "12pt" as 24 half-points to <w:sz> and <w:szCs>', () => {
      const pXml = parseXml(
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Hello World</w:t></w:r></w:p>'
      );
      const node = {
        id: 'p_1',
        type: 'paragraph' as const,
        xml: pXml,
        metadata: { text: 'Hello World' },
      };
      const ast = buildTestAst([node]);

      executeStyle(
        {
          action: 'style',
          target: 'p_1',
          sz: '12pt',
        },
        ast
      );

      const szEl = pXml.getElementsByTagName('w:sz')[0];
      const szCsEl = pXml.getElementsByTagName('w:szCs')[0];
      expect(szEl).toBeDefined();
      expect(szEl?.getAttribute('w:val')).toBe('24');
      expect(szCsEl?.getAttribute('w:val')).toBe('24');
    });

    it('executeStyle applies "10.5pt" as 21 half-points to <w:sz> and <w:szCs>', () => {
      const pXml = parseXml(
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Quarter Point</w:t></w:r></w:p>'
      );
      const node = {
        id: 'p_2',
        type: 'paragraph' as const,
        xml: pXml,
        metadata: { text: 'Quarter Point' },
      };
      const ast = buildTestAst([node]);

      executeStyle(
        {
          action: 'style',
          target: 'p_2',
          fontSize: '10.5pt',
        },
        ast
      );

      const szEl = pXml.getElementsByTagName('w:sz')[0];
      const szCsEl = pXml.getElementsByTagName('w:szCs')[0];
      expect(szEl?.getAttribute('w:val')).toBe('21');
      expect(szCsEl?.getAttribute('w:val')).toBe('21');
    });

    it('replaceText applies "12pt" as 24 half-points to run properties', () => {
      const pXml = parseXml(
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Old text</w:t></w:r></w:p>'
      );
      replaceText(pXml, 'New text', true, '12pt');

      const szEl = pXml.getElementsByTagName('w:sz')[0];
      const szCsEl = pXml.getElementsByTagName('w:szCs')[0];
      expect(szEl?.getAttribute('w:val')).toBe('24');
      expect(szCsEl?.getAttribute('w:val')).toBe('24');
    });

    it('unified semantics: bare 12 → w:sz 12 identical on style AND replace paths', () => {
      const styleXml = parseXml(
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Hello</w:t></w:r></w:p>'
      );
      const styleAst = buildTestAst([{ id: 'p_1', type: 'paragraph', xml: styleXml, metadata: { text: 'Hello' } }]);
      executeStyle({ action: 'style', target: 'p_1', properties: { fontSize: 12 } }, styleAst);

      const repXml = parseXml(
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Old</w:t></w:r></w:p>'
      );
      const repAst = buildTestAst([{ id: 'p_1', type: 'paragraph', xml: repXml, metadata: { text: 'Old' } }]);
      executeReplace({ action: 'replace', target: 'p_1', text: 'New', fontSize: 12 }, repAst);

      const styleSz = styleXml.getElementsByTagName('w:sz')[0]?.getAttribute('w:val');
      const repSz = repXml.getElementsByTagName('w:sz')[0]?.getAttribute('w:val');
      expect(styleSz).toBe('12');
      expect(repSz).toBe('12');
      expect(styleSz).toBe(repSz);
    });

    it('unified semantics: "12pt" → w:sz 24 identical on style AND replace paths', () => {
      const styleXml = parseXml(
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Hello</w:t></w:r></w:p>'
      );
      const styleAst = buildTestAst([{ id: 'p_1', type: 'paragraph', xml: styleXml, metadata: { text: 'Hello' } }]);
      executeStyle({ action: 'style', target: 'p_1', properties: { fontSize: '12pt' } }, styleAst);

      const repXml = parseXml(
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Old</w:t></w:r></w:p>'
      );
      const repAst = buildTestAst([{ id: 'p_1', type: 'paragraph', xml: repXml, metadata: { text: 'Old' } }]);
      executeReplace({ action: 'replace', target: 'p_1', text: 'New', fontSize: '12pt' }, repAst);

      const styleSz = styleXml.getElementsByTagName('w:sz')[0]?.getAttribute('w:val');
      const repSz = repXml.getElementsByTagName('w:sz')[0]?.getAttribute('w:val');
      expect(styleSz).toBe('24');
      expect(repSz).toBe('24');
      expect(styleSz).toBe(repSz);
    });
  });

  // =========================================================================
  // L-13: DocxParser Path-String Overload
  // =========================================================================
  describe('L-13: DocxParser Path-String Overload', () => {
    let tmpDocxPath: string;

    beforeEach(() => {
      tmpDocxPath = path.join(os.tmpdir(), `astra-l13-${Date.now()}-${Math.random().toString(36).slice(2)}.docx`);
      fs.writeFileSync(tmpDocxPath, blankDocxBuf());
    });

    afterEach(() => {
      try {
        fs.unlinkSync(tmpDocxPath);
      } catch {
        /* best effort */
      }
    });

    it('DocxParser accepts a string file path and parses successfully', async () => {
      const parser = new DocxParser(tmpDocxPath);
      const ast = await parser.getAST();
      expect(ast).toBeDefined();
      expect(ast.sections.length).toBeGreaterThan(0);
    });

    it('DocxParser throws descriptive Error("File not found: " + input) for non-existent path', () => {
      const fakePath = '/tmp/definitely_not_existing_astra_file_123456789.docx';
      expect(() => new DocxParser(fakePath)).toThrow(`File not found: ${fakePath}`);
    });

    it('DocxParser retains Buffer constructor functionality', async () => {
      const buf = blankDocxBuf();
      const parser = new DocxParser(buf);
      const ast = await parser.getAST();
      expect(ast).toBeDefined();
      expect(ast.sections.length).toBeGreaterThan(0);
    });
  });

  // =========================================================================
  // T-13: Respect Non-TTY and NO_COLOR in CLI Output
  // =========================================================================
  describe('T-13: CLI Non-TTY and NO_COLOR Color Stripping', () => {
    const originalNoColor = process.env.NO_COLOR;

    afterEach(() => {
      if (originalNoColor !== undefined) {
        process.env.NO_COLOR = originalNoColor;
      } else {
        delete process.env.NO_COLOR;
      }
    });

    it('ANSI_REGEX detects and stripAnsi strips ANSI escape codes', () => {
      const colored = '\u001b[32mSuccess!\u001b[0m \u001b[1;31mError here\u001b[0m';
      expect(stripAnsi(colored)).toBe('Success! Error here');
    });

    it('shouldStripColor returns true when process.env.NO_COLOR is defined', () => {
      process.env.NO_COLOR = '1';
      expect(shouldStripColor({ isTTY: true })).toBe(true);
      expect(shouldStripColor({ isTTY: false })).toBe(true);
    });

    it('shouldStripColor returns true when stream is not a TTY', () => {
      delete process.env.NO_COLOR;
      expect(shouldStripColor({ isTTY: false })).toBe(true);
      expect(shouldStripColor({})).toBe(true);
    });

    it('shouldStripColor returns false when stream is TTY and NO_COLOR is undefined', () => {
      delete process.env.NO_COLOR;
      expect(shouldStripColor({ isTTY: true })).toBe(false);
    });

    it('formatCliOutput strips ANSI when stream is not a TTY or NO_COLOR is set', () => {
      const colored = '\u001b[34mInfo message\u001b[0m';
      const nonTtyStream = { isTTY: false };
      expect(formatCliOutput(colored, nonTtyStream)).toBe('Info message');

      const ttyStream = { isTTY: true };
      delete process.env.NO_COLOR;
      expect(formatCliOutput(colored, ttyStream)).toBe(colored);

      process.env.NO_COLOR = '';
      // no-color.org: only non-empty NO_COLOR disables color, so empty value preserves color
      expect(formatCliOutput(colored, ttyStream)).toBe(colored);
    });
  });
});

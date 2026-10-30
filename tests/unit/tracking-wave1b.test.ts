/**
 * Wave 1B — track-changes revision emission for replace + table updateCell.
 *
 * - Fresh (version 1) replace with tracking on must emit w:ins + w:del
 *   instead of silently overwriting text.
 * - Tracking off must stay silent (no revision markup).
 * - updateCell with tracking on must emit revision markup; tracking off
 *   must keep the silent text-swap bytes.
 *
 * Markup assertions read the live DOM serialization or SAVED bytes
 * (rezip -> word/document.xml), never just appliedCount.
 */
import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { XMLSerializer } from '@xmldom/xmldom';
import {
  DocxParser,
  createBlankDocxBuffer,
  convertAstToHtml,
  convertAstToMarkdown,
  executeInsert,
  executeReplace,
  executeTable,
  validateOperationStrict,
} from '@attentionform/astra-engine';

const serializer = new XMLSerializer();

function savedDocumentXml(parser: DocxParser, ast: any): string {
  const buf = parser.rezip(ast);
  return new AdmZip(buf).readAsText('word/document.xml');
}

/** Blank doc + one inserted paragraph with known text (a fresh version-1 node). */
async function freshParagraph(oldText: string, trackChanges: boolean) {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  ast.settings.trackChanges = trackChanges;
  const anchor = ast.sections[0].nodes[0].id;
  await executeInsert({
    action: 'insert', type: 'paragraph', anchor, position: 'after', text: oldText,
  } as any, ast, parser);
  const node = ast.sections[0].nodes.find((n: any) => n.metadata?.text === oldText);
  expect(node).toBeDefined();
  expect(node.version).toBe(1);
  return { parser, ast, nodeId: node.id as string };
}

/** Blank doc + one 2-row x 2-col table (headers H1/H2, data R1/R2). */
async function freshTable(trackChanges: boolean) {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  ast.settings.trackChanges = trackChanges;
  const anchor = ast.sections[0].nodes[0].id;
  await executeInsert({
    action: 'insert', type: 'table', anchor, position: 'after',
    headers: ['H1', 'H2'], rows: [['R1', 'R2']],
  } as any, ast, parser);
  const tblNode = ast.sections[0].nodes.find((n: any) => n.type === 'table');
  expect(tblNode).toBeDefined();
  return { parser, ast, tableId: tblNode.id as string };
}

describe('tracking wave 1B: replace revision emission', () => {
  it('fresh-node replace with tracking on yields w:ins + w:del', () => {
    return (async () => {
      const { ast, nodeId } = await freshParagraph('ORIGINAL TEXT HERE', true);
      executeReplace({ action: 'replace', target: nodeId, text: 'COMPLETELY NEW WORDING' } as any, ast);
      const node = ast.sections[0].nodes.find((n: any) => n.id === nodeId);
      const xml = serializer.serializeToString(node.xml);
      expect(xml).toContain('w:ins');
      expect(xml).toContain('w:del');
      expect(xml).toContain('w:delText');
      // Word-level diff emits one run per word, so assert each word lands
      // inside an insertion run rather than one contiguous string.
      for (const word of ['COMPLETELY', 'NEW', 'WORDING']) {
        expect(xml).toContain(word);
      }
      expect(node.metadata.text).toBe('COMPLETELY NEW WORDING');
    })();
  });

  it('replace with tracking off stays silent', () => {
    return (async () => {
      const { ast, nodeId } = await freshParagraph('ORIGINAL TEXT HERE', false);
      executeReplace({ action: 'replace', target: nodeId, text: 'COMPLETELY NEW WORDING' } as any, ast);
      const node = ast.sections[0].nodes.find((n: any) => n.id === nodeId);
      const xml = serializer.serializeToString(node.xml);
      expect(xml).not.toContain('w:ins');
      expect(xml).not.toContain('w:del');
      expect(xml).not.toContain('w:delText');
      expect(xml).toContain('COMPLETELY NEW WORDING');
    })();
  });

  it('op-level trackChanges:false opts out even when the document toggle is on', () => {
    return (async () => {
      const { ast, nodeId } = await freshParagraph('ORIGINAL TEXT HERE', true);
      executeReplace(
        { action: 'replace', target: nodeId, text: 'COMPLETELY NEW WORDING', trackChanges: false } as any,
        ast,
      );
      const node = ast.sections[0].nodes.find((n: any) => n.id === nodeId);
      const xml = serializer.serializeToString(node.xml);
      expect(xml).not.toContain('w:ins');
      expect(xml).not.toContain('w:del');
      expect(xml).toContain('COMPLETELY NEW WORDING');
    })();
  });
});

describe('tracking wave 1B: updateCell revision emission', () => {
  it('updateCell with tracking on emits cell revision markup', async () => {
    const { parser, ast, tableId } = await freshTable(true);
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'updateCell', rowIndex: 2, colIndex: 1, text: 'REVISED' },
    } as any, ast);
    const saved = savedDocumentXml(parser, ast);
    // Element-name regexes: plain 'w:ins' would false-positive on
    // w:insideH / w:insideV table-border elements.
    expect(saved).toMatch(/<w:ins[\s>]/);
    expect(saved).toMatch(/<w:del[\s>]/);
    expect(saved).toContain('REVISED');
    // Old cell text survives only as deleted revision content.
    expect(saved).toMatch(/<w:delText[\s>]/);
  });

  it('updateCell with tracking off keeps the silent text-swap bytes', async () => {
    const { parser, ast, tableId } = await freshTable(false);
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'updateCell', rowIndex: 2, colIndex: 1, text: 'REVISED' },
    } as any, ast);
    const saved = savedDocumentXml(parser, ast);
    expect(saved).toContain('REVISED');
    // Element-name regexes: plain 'w:ins' would false-positive on
    // w:insideH / w:insideV table-border elements.
    expect(saved).not.toMatch(/<w:ins[\s>]/);
    expect(saved).not.toMatch(/<w:del[\s>]/);
    expect(saved).not.toMatch(/<w:delText[\s>]/);
    expect(saved).not.toContain('cellIns');
    expect(saved).not.toContain('cellDel');
  });
});

/** Blank-document default: no toggle touched, so the parsed default applies. */
async function defaultParagraph(oldText: string) {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const anchor = ast.sections[0].nodes[0].id;
  await executeInsert({
    action: 'insert', type: 'paragraph', anchor, position: 'after', text: oldText,
  } as any, ast, parser);
  const node = ast.sections[0].nodes.find((n: any) => n.metadata?.text === oldText);
  expect(node).toBeDefined();
  return { parser, ast, nodeId: node.id as string };
}

async function defaultTable() {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const anchor = ast.sections[0].nodes[0].id;
  await executeInsert({
    action: 'insert', type: 'table', anchor, position: 'after',
    headers: ['H1', 'H2'], rows: [['R1', 'R2']],
  } as any, ast, parser);
  const tblNode = ast.sections[0].nodes.find((n: any) => n.type === 'table');
  expect(tblNode).toBeDefined();
  return { parser, ast, tableId: tblNode.id as string };
}

/** Blank docx bytes plus a word/settings.xml carrying the given inner flag. */
function bufferWithSettingsInner(inner: string): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  const settingsXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${inner}</w:settings>`;
  zip.addFile('word/settings.xml', Buffer.from(settingsXml, 'utf8'));
  const ct = zip.readAsText('[Content_Types].xml');
  if (!ct.includes('/word/settings.xml')) {
    zip.updateFile(
      '[Content_Types].xml',
      Buffer.from(
        ct.replace(
          '</Types>',
          '  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>\n</Types>',
        ),
        'utf8',
      ),
    );
  }
  return zip.toBuffer();
}

describe('blank-document default: edits stay silent unless opted in', () => {
  it('parsed blank defaults tracking off', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    expect(ast.settings.trackChanges).toBe(false);
    expect(ast.settings.defaultAuthor).toBe('ASTra');
  });

  it('blank-doc replace yields clean w:t with zero ins/del', async () => {
    const { parser, ast, nodeId } = await defaultParagraph('ORIGINAL TEXT HERE');
    executeReplace({ action: 'replace', target: nodeId, text: 'COMPLETELY NEW WORDING' } as any, ast);
    const node = ast.sections[0].nodes.find((n: any) => n.id === nodeId);
    const xml = serializer.serializeToString(node.xml);
    expect(xml).toContain('<w:t');
    expect(xml).toContain('COMPLETELY NEW WORDING');
    expect(xml).not.toMatch(/<w:ins[\s>]/);
    expect(xml).not.toMatch(/<w:del[\s>]/);
    expect(xml).not.toMatch(/<w:delText[\s>]/);
    const saved = savedDocumentXml(parser, ast);
    expect(saved).toContain('COMPLETELY NEW WORDING');
    expect(saved.match(/<w:ins[\s>]/g) || []).toHaveLength(0);
    expect(saved.match(/<w:del[\s>]/g) || []).toHaveLength(0);
    expect(saved.match(/<w:delText[\s>]/g) || []).toHaveLength(0);
  });

  it('explicit opt-in redlines with author and timestamp', async () => {
    const { ast, nodeId } = await defaultParagraph('ORIGINAL TEXT HERE');
    executeReplace(
      { action: 'replace', target: nodeId, text: 'COMPLETELY NEW WORDING', trackChanges: true } as any,
      ast,
    );
    const node = ast.sections[0].nodes.find((n: any) => n.id === nodeId);
    const xml = serializer.serializeToString(node.xml);
    expect(xml).toMatch(/<w:ins[\s>]/);
    expect(xml).toMatch(/<w:del[\s>]/);
    expect(xml).toMatch(/<w:delText[\s>]/);
    expect(xml).toContain('w:author="ASTra"');
    const date = xml.match(/w:date="([^"]+)"/)?.[1];
    expect(date).toBeTruthy();
    expect(Number.isFinite(Date.parse(date!))).toBe(true);
    expect(xml).toMatch(/w:id="\d+"/);
    expect(node.metadata.text).toBe('COMPLETELY NEW WORDING');
  });

  it('params-nested opt-in redlines too', async () => {
    const { ast, nodeId } = await defaultParagraph('ORIGINAL TEXT HERE');
    executeReplace(
      { action: 'replace', target: nodeId, text: 'COMPLETELY NEW WORDING', params: { trackChanges: true } } as any,
      ast,
    );
    const node = ast.sections[0].nodes.find((n: any) => n.id === nodeId);
    const xml = serializer.serializeToString(node.xml);
    expect(xml).toMatch(/<w:ins[\s>]/);
    expect(xml).toMatch(/<w:del[\s>]/);
  });
});

describe('settings flag: real tracking signal opts back in', () => {
  it('flag-bearing doc parses tracking on and tracks a plain replace', async () => {
    const parser = new DocxParser(bufferWithSettingsInner('<w:trackRevisions/>'));
    const ast: any = await parser.getAST();
    expect(ast.settings.trackChanges).toBe(true);
    const target = ast.sections[0].nodes[0];
    const oldText: string = target.metadata.text;
    executeReplace({ action: 'replace', target: target.id, text: `${oldText} appended` } as any, ast);
    const xml = serializer.serializeToString(target.xml);
    expect(xml).toMatch(/<w:ins[\s>]/);
    expect(xml).toContain('appended');
    expect(xml).toContain('w:author="ASTra"');
  });

  it('explicit val=false flag keeps tracking off', async () => {
    const parser = new DocxParser(bufferWithSettingsInner('<w:trackRevisions w:val="false"/>'));
    const ast: any = await parser.getAST();
    expect(ast.settings.trackChanges).toBe(false);
    const target = ast.sections[0].nodes[0];
    const oldText: string = target.metadata.text;
    executeReplace({ action: 'replace', target: target.id, text: `${oldText} appended` } as any, ast);
    const xml = serializer.serializeToString(target.xml);
    expect(xml).not.toMatch(/<w:ins[\s>]/);
    expect(xml).not.toMatch(/<w:del[\s>]/);
  });
});

describe('table cell parity: same default and opt-in rule as paragraphs', () => {
  it('blank-doc updateCell stays a silent swap with zero ins/del', async () => {
    const { parser, ast, tableId } = await defaultTable();
    expect(ast.settings.trackChanges).toBe(false);
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'updateCell', rowIndex: 2, colIndex: 1, text: 'REVISED' },
    } as any, ast);
    const saved = savedDocumentXml(parser, ast);
    expect(saved).toContain('REVISED');
    expect(saved.match(/<w:ins[\s>]/g) || []).toHaveLength(0);
    expect(saved.match(/<w:del[\s>]/g) || []).toHaveLength(0);
    expect(saved.match(/<w:delText[\s>]/g) || []).toHaveLength(0);
    expect(saved).toMatch(/<w:t[^>]*>REVISED<\/w:t>/);
  });

  it('explicit opt-in updateCell redlines with author and timestamp', async () => {
    const { parser, ast, tableId } = await defaultTable();
    await executeTable({
      action: 'table', target: tableId, trackChanges: true,
      params: { op: 'updateCell', rowIndex: 2, colIndex: 1, text: 'REVISED' },
    } as any, ast);
    const saved = savedDocumentXml(parser, ast);
    expect(saved).toMatch(/<w:ins[\s>]/);
    expect(saved).toMatch(/<w:del[\s>]/);
    expect(saved).toMatch(/<w:delText[\s>]/);
    expect(saved).toContain('REVISED');
    expect(saved).toContain('w:author="ASTra"');
    const date = saved.match(/w:date="([^"]+)"/)?.[1];
    expect(date).toBeTruthy();
    expect(Number.isFinite(Date.parse(date!))).toBe(true);
  });

  it('flag-bearing doc tracks a plain updateCell', async () => {
    const setup = new DocxParser(bufferWithSettingsInner('<w:trackRevisions/>'));
    const setupAst: any = await setup.getAST();
    expect(setupAst.settings.trackChanges).toBe(true);
    const anchor = setupAst.sections[0].nodes[0].id;
    await executeInsert({
      action: 'insert', type: 'table', anchor, position: 'after',
      headers: ['H1', 'H2'], rows: [['R1', 'R2']],
    } as any, setupAst, setup);
    const tblNode = setupAst.sections[0].nodes.find((n: any) => n.type === 'table');
    await executeTable({
      action: 'table', target: tblNode.id,
      params: { op: 'updateCell', rowIndex: 2, colIndex: 1, text: 'REVISED' },
    } as any, setupAst);
    const buf = setup.rezip(setupAst);
    const saved = new AdmZip(buf).readAsText('word/document.xml');
    expect(saved).toMatch(/<w:ins[\s>]/);
    expect(saved).toMatch(/<w:del[\s>]/);
    expect(saved).toContain('REVISED');
  });
});

describe('tracked paragraph insert', () => {
  it('schema gate accepts trackChanges on paragraph inserts and rejects non-boolean values', () => {
    const ok = validateOperationStrict({
      action: 'insert', type: 'paragraph', anchor: 'p1', position: 'after',
      text: 'Hello', trackChanges: true,
    });
    expect(ok.ok).toBe(true);
    expect(() => validateOperationStrict({
      action: 'insert', type: 'paragraph', anchor: 'p1', position: 'after',
      text: 'Hello', trackChanges: 'yes',
    })).toThrow(/E_VALIDATION/);
  });

  it('wraps inserted runs in w:ins with author, date, and change id', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const anchor = ast.sections[0].nodes[0].id;
    await executeInsert({
      action: 'insert', type: 'paragraph', anchor, position: 'after',
      text: 'Tracked hello', trackChanges: true,
    } as any, ast, parser);
    const node = ast.sections[0].nodes.find((n: any) => n.metadata?.text === 'Tracked hello');
    expect(node).toBeDefined();
    expect(node.metadata.hasTrackedChanges).toBe(true);
    const saved = savedDocumentXml(parser, ast);
    expect(saved).toMatch(/<w:ins w:author="ASTra" w:date="[^"]+" w:id="\d+"/);
    expect(saved).toContain('<w:t>Tracked hello</w:t>');
  });

  it('fails closed on tracked non-paragraph inserts without mutating the document', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const before = savedDocumentXml(parser, ast);
    const anchor = ast.sections[0].nodes[0].id;
    await expect(executeInsert({
      action: 'insert', type: 'table', anchor, position: 'after',
      headers: ['A'], rows: [['1']], trackChanges: true,
    } as any, ast, parser)).rejects.toThrow(/E_VALIDATION/);
    expect(savedDocumentXml(parser, ast)).toBe(before);
    await expect(executeInsert({
      action: 'insert', type: 'paragraph', anchor, position: 'after',
      text: 'x', trackChanges: 'yes',
    } as any, ast, parser)).rejects.toThrow(/E_VALIDATION/);
  });

  it('explicit false stays silent with no revision markup', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const anchor = ast.sections[0].nodes[0].id;
    await executeInsert({
      action: 'insert', type: 'paragraph', anchor, position: 'after',
      text: 'Silent hello', trackChanges: false,
    } as any, ast, parser);
    const node = ast.sections[0].nodes.find((n: any) => n.metadata?.text === 'Silent hello');
    expect(node.metadata.hasTrackedChanges).toBe(false);
    expect(savedDocumentXml(parser, ast)).not.toContain('<w:ins');
  });

  it('tracked insert renders redline in HTML and accepted text in Markdown', async () => {    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const anchor = ast.sections[0].nodes[0].id;
    await executeInsert({
      action: 'insert', type: 'paragraph', anchor, position: 'after',
      text: 'Review me', trackChanges: true,
    } as any, ast, parser);
    const html = convertAstToHtml(ast, 'tracked');
    expect(html).toContain('<ins class="tracked-ins">');
    expect(html).toContain('Review me');
    const md = convertAstToMarkdown(ast);
    expect(md).toContain('Review me');
  });

  it('wraps multi-line inserts in one w:ins holding the break runs', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const anchor = ast.sections[0].nodes[0].id;
    await executeInsert({
      action: 'insert', type: 'paragraph', anchor, position: 'after',
      text: 'first\nsecond', trackChanges: true,
    } as any, ast, parser);
    const saved = savedDocumentXml(parser, ast);
    const insBody = saved.match(/<w:ins[\s>][\s\S]*?<\/w:ins>/)?.[0] || '';
    expect(insBody).toContain('<w:t>first</w:t>');
    expect(insBody).toContain('<w:t>second</w:t>');
    expect((insBody.match(/<w:br[\s/>]/g) || []).length).toBe(1);
  });
});

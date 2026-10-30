import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import {
  createWordElement,
  DocumentManager,
  DocxParser,
  extractStylePalette,
  formatPaletteForPrompt,
  lintAndFixTypography,
  OperationEngine,
  PPR_CHILD_ORDER,
  TRPR_CHILD_ORDER,
  validateDocxBuffer,
} from '@attentionform/astra-engine';

/**
 * Helper to verify that an element's child XML elements strictly adhere
 * to the specified ECMA-376 schema sequence order without out-of-order children.
 */
function assertStrictChildOrder(element: Element, expectedOrder: readonly string[]): void {
  let lastIndex = -1;
  for (let i = 0; i < element.childNodes.length; i++) {
    const child = element.childNodes[i];
    if (child.nodeType === 1) {
      const tagName = (child as Element).tagName;
      const orderIndex = expectedOrder.indexOf(tagName);
      expect(orderIndex, `Element <${tagName}> must be recognized in schema order definition`).toBeGreaterThanOrEqual(0);
      expect(
        orderIndex,
        `Element <${tagName}> (order index ${orderIndex}) must appear after previous element (order index ${lastIndex})`
      ).toBeGreaterThan(lastIndex);
      lastIndex = orderIndex;
    }
  }
}

/**
 * Helper to create a valid ECMA-376 table for typographic linter stress tests.
 */
function createTestTable(
  doc: Document,
  numRows: number,
  numCols: number,
  options?: { multiline?: boolean; nestedTable?: boolean }
): Element {
  const tbl = createWordElement(doc, 'w:tbl');
  const tblPr = createWordElement(doc, 'w:tblPr');
  const tblW = createWordElement(doc, 'w:tblW');
  tblW.setAttribute('w:w', '5000');
  tblW.setAttribute('w:type', 'dxa');
  tblPr.appendChild(tblW);
  tbl.appendChild(tblPr);

  const tblGrid = createWordElement(doc, 'w:tblGrid');
  const colWidth = Math.floor(5000 / numCols);
  for (let c = 0; c < numCols; c++) {
    const gridCol = createWordElement(doc, 'w:gridCol');
    gridCol.setAttribute('w:w', String(colWidth));
    tblGrid.appendChild(gridCol);
  }
  tbl.appendChild(tblGrid);

  for (let r = 0; r < numRows; r++) {
    const tr = createWordElement(doc, 'w:tr');
    const trPr = createWordElement(doc, 'w:trPr');
    const trHeight = createWordElement(doc, 'w:trHeight');
    trHeight.setAttribute('w:val', '320');
    trPr.appendChild(trHeight);
    tr.appendChild(trPr);

    for (let c = 0; c < numCols; c++) {
      const tc = createWordElement(doc, 'w:tc');
      const tcPr = createWordElement(doc, 'w:tcPr');
      const tcW = createWordElement(doc, 'w:tcW');
      tcW.setAttribute('w:w', String(colWidth));
      tcW.setAttribute('w:type', 'dxa');
      tcPr.appendChild(tcW);
      tc.appendChild(tcPr);

      // Add primary paragraph to cell
      const p1 = createWordElement(doc, 'w:p');
      const r1 = createWordElement(doc, 'w:r');
      const t1 = createWordElement(doc, 'w:t');
      t1.appendChild(doc.createTextNode(`Row ${r + 1} Col ${c + 1}`));
      r1.appendChild(t1);
      p1.appendChild(r1);
      tc.appendChild(p1);

      if (options?.multiline && r > 0) {
        const p2 = createWordElement(doc, 'w:p');
        const r2 = createWordElement(doc, 'w:r');
        const t2 = createWordElement(doc, 'w:t');
        t2.appendChild(doc.createTextNode(`Row ${r + 1} Col ${c + 1} Multiline paragraph`));
        r2.appendChild(t2);
        p2.appendChild(r2);
        tc.appendChild(p2);
      }

      if (options?.nestedTable && r === 1 && c === 0) {
        // Create nested table with 2 rows, 2 cols
        const nestedTbl = createTestTable(doc, 2, 2);
        tc.appendChild(nestedTbl);
        // Schema requirement: tc must end with w:p
        const pAfter = createWordElement(doc, 'w:p');
        tc.appendChild(pAfter);
      }

      tr.appendChild(tc);
    }
    tbl.appendChild(tr);
  }

  return tbl;
}

describe('Sprint 2 Moat Capabilities (MOAT-1, MOAT-2, MOAT-3, MOAT-0)', () => {
  let tmpDir: string;

  beforeEach(() => {
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-moat-test-'));
  });

  afterEach(() => {
    try {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    } catch {
      // best-effort cleanup
    }
  });

  describe('MOAT-1: Style Palette Introspection', () => {
    it('extracts styles, fonts, and brand colors from T1_coffee_onepager gallery document', async () => {
      const galleryPath = path.resolve('test-gallery/T1_coffee_onepager.docx');
      expect(fs.existsSync(galleryPath)).toBe(true);

      const buf = fs.readFileSync(galleryPath);
      const parser = new DocxParser(buf);
      await parser.getAST();

      const palette = parser.getStylePalette();
      expect(palette).toBeDefined();
      expect(typeof palette.primaryFont).toBe('string');
      expect(palette.primaryFont.length).toBeGreaterThan(0);
      expect(Array.isArray(palette.availableStyles)).toBe(true);
      expect(palette.availableStyles.length).toBeGreaterThan(0);

      // Verify normal and heading styles are discovered
      const styleIds = palette.availableStyles.map(s => s.id);
      expect(styleIds).toContain('Normal');

      // Verify formatPaletteForPrompt produces readable token-efficient summary
      const promptSummary = formatPaletteForPrompt(palette);
      expect(promptSummary).toContain('Primary Font:');
      expect(promptSummary).toContain('Available Heading Styles:');
    });

    it('extracts style palette from pure styles.xml string', () => {
      const sampleStylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:docDefaults>
            <w:rPrDefault>
              <w:rPr>
                <w:rFonts w:ascii="Inter" w:hAnsi="Inter"/>
                <w:sz w:val="22"/>
              </w:rPr>
            </w:rPrDefault>
          </w:docDefaults>
          <w:style w:type="paragraph" w:styleId="Heading1">
            <w:name w:val="heading 1"/>
            <w:rPr>
              <w:rFonts w:ascii="Playfair Display"/>
              <w:color w:val="1E3A8A"/>
              <w:sz w:val="40"/>
              <w:b/>
            </w:rPr>
          </w:style>
          <w:style w:type="paragraph" w:styleId="CorporateCallout">
            <w:name w:val="Corporate Callout"/>
            <w:rPr>
              <w:color w:val="2563EB"/>
              <w:sz w:val="24"/>
              <w:i/>
            </w:rPr>
          </w:style>
        </w:styles>`;

      const palette = extractStylePalette(sampleStylesXml);
      expect(palette.primaryFont).toBe('Inter');
      expect(palette.headingFont).toBe('Playfair Display');
      expect(palette.brandColors).toContain('1E3A8A');
      expect(palette.brandColors).toContain('2563EB');
      expect(palette.headingStyles).toContain('Heading1');

      const calloutStyle = palette.availableStyles.find(s => s.id === 'CorporateCallout');
      expect(calloutStyle).toBeDefined();
      expect(calloutStyle?.color).toBe('2563EB');
      expect(calloutStyle?.italic).toBe(true);
      expect(calloutStyle?.fontSizePt).toBe(12);
    });

    it('extracts style palette against every single .docx in test-gallery/*.docx (all 9 gallery files)', async () => {
      const galleryFiles = fs.readdirSync('test-gallery').filter(f => f.endsWith('.docx'));
      expect(galleryFiles.length).toBe(9);

      for (const file of galleryFiles) {
        const filePath = path.resolve('test-gallery', file);
        const buf = fs.readFileSync(filePath);
        const parser = new DocxParser(buf);
        await parser.getAST();

        const palette = parser.getStylePalette();
        expect(palette, `Palette should be defined for ${file}`).toBeDefined();

        // Assert that primaryFont is non-empty for all valid docs
        expect(typeof palette.primaryFont, `primaryFont must be string for ${file}`).toBe('string');
        expect(palette.primaryFont.trim().length, `primaryFont must not be empty for ${file}`).toBeGreaterThan(0);

        // Assert that brandColors contains valid 6-character hex strings without '#'
        for (const color of palette.brandColors) {
          expect(color, `brandColor in ${file} must not start with #`).not.toMatch(/^#/);
          expect(color, `brandColor ${color} in ${file} must be valid 6-char hex`).toMatch(/^[0-9A-Fa-f]{6}$/);
          expect(color, `brandColor in ${file} must not be pure black or white`).not.toBe('000000');
          expect(color, `brandColor in ${file} must not be pure black or white`).not.toBe('FFFFFF');
        }

        // Assert formatPaletteForPrompt output is compact, readable, and under 500 tokens (< ~2000 chars)
        const promptSummary = formatPaletteForPrompt(palette);
        expect(promptSummary.length, `Prompt summary for ${file} must be non-empty`).toBeGreaterThan(0);
        expect(promptSummary.length, `Prompt summary for ${file} must be under 2000 characters (~500 tokens)`).toBeLessThan(2000);
        const approxTokenCount = promptSummary.split(/\s+/).length * 1.3;
        expect(approxTokenCount, `Estimated token count for ${file} must be under 500`).toBeLessThan(500);
      }
    });

    it('handles MOAT-1 edge cases: empty <w:styles>, styles with no colors, auto color, and docDefaults overrides', () => {
      // 1. Empty <w:styles>
      const emptyStylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        </w:styles>`;
      const emptyPalette = extractStylePalette(emptyStylesXml);
      expect(emptyPalette.primaryFont).toBe('Calibri');
      expect(emptyPalette.brandColors).toEqual([]);
      expect(emptyPalette.availableStyles).toEqual([]);
      expect(emptyPalette.headingStyles).toEqual([]);
      const emptyPrompt = formatPaletteForPrompt(emptyPalette);
      expect(emptyPrompt).toContain('Primary Font: Calibri');
      expect(emptyPrompt).toContain('Available Heading Styles: None');
      expect(emptyPrompt.length).toBeLessThan(500);

      // 2. Styles with no colors
      const noColorStylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:style w:type="paragraph" w:styleId="BodyPlain">
            <w:name w:val="Body Plain"/>
            <w:rPr>
              <w:rFonts w:ascii="Arial"/>
              <w:sz w:val="24"/>
              <w:b/>
            </w:rPr>
          </w:style>
        </w:styles>`;
      const noColorPalette = extractStylePalette(noColorStylesXml);
      expect(noColorPalette.brandColors).toEqual([]);
      expect(noColorPalette.availableStyles.length).toBe(1);
      expect(noColorPalette.availableStyles[0].color).toBeUndefined();

      // 3. Styles with 'auto' color and 'AUTO' uppercase
      const autoColorStylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:style w:type="paragraph" w:styleId="StyleAuto">
            <w:name w:val="Style Auto"/>
            <w:rPr>
              <w:color w:val="auto"/>
            </w:rPr>
          </w:style>
          <w:style w:type="paragraph" w:styleId="StyleAutoCaps">
            <w:name w:val="Style Auto Caps"/>
            <w:rPr>
              <w:color w:val="AUTO"/>
            </w:rPr>
          </w:style>
        </w:styles>`;
      const autoPalette = extractStylePalette(autoColorStylesXml);
      expect(autoPalette.brandColors).toEqual([]);
      expect(autoPalette.availableStyles[0].color).toBeUndefined();
      expect(autoPalette.availableStyles[1].color).toBeUndefined();

      // 4. docDefaults overrides
      const docDefaultsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:docDefaults>
            <w:rPrDefault>
              <w:rPr>
                <w:rFonts w:ascii="Fira Code" w:hAnsi="Fira Code"/>
              </w:rPr>
            </w:rPrDefault>
          </w:docDefaults>
        </w:styles>`;
      const docDefPalette = extractStylePalette(docDefaultsXml);
      expect(docDefPalette.primaryFont).toBe('Fira Code');
    });
  });

  describe('MOAT-2: Typographic Linter Engine', () => {
    it('applies keepNext on headings, cantSplit on rows, and tblHeader on table header', async () => {
      const dm = new DocumentManager();
      await dm.createNewDocument('report');

      const parser = dm.getParser();
      expect(parser).toBeDefined();

      const docDom = parser.getDocumentDom();

      // Run linter directly
      const lintResult = lintAndFixTypography(docDom);
      expect(lintResult.applied).toBe(true);
      expect(lintResult.fixes.length).toBeGreaterThan(0);

      // Verify that heading has keepNext
      const pNodes = docDom.getElementsByTagName('w:p');
      let foundHeadingWithKeepNext = false;
      for (let i = 0; i < pNodes.length; i++) {
        const p = pNodes[i];
        const pPr = p.getElementsByTagName('w:pPr')[0];
        if (pPr) {
          const pStyle = pPr.getElementsByTagName('w:pStyle')[0];
          if (pStyle && /heading/i.test(pStyle.getAttribute('w:val') || '')) {
            const keepNext = pPr.getElementsByTagName('w:keepNext');
            if (keepNext.length > 0) foundHeadingWithKeepNext = true;
          }
        }
      }
      expect(foundHeadingWithKeepNext).toBe(true);

      // Save document and verify that validator passes ok: true
      const savePath = path.join(tmpDir, 'typography_tested.docx');
      const saveRes = await dm.saveAsDocument(savePath);
      expect(saveRes.typographyFixes).toBeDefined();
      expect(fs.existsSync(savePath)).toBe(true);

      const savedBuf = fs.readFileSync(savePath);
      const validation = await validateDocxBuffer(savedBuf);
      expect(validation.ok).toBe(true);
    });

    it('creates document with 10 tables, nested tables, orphan headings (H1, H2, H3, Title), and multiline table rows, verifies schema order, and passes validator', async () => {
      const dm = new DocumentManager();
      await dm.createNewDocument('blank');

      const parser = dm.getParser();
      const dom = parser.getDocumentDom();
      const body = dom.getElementsByTagName('w:body')[0];
      expect(body).toBeDefined();

      // Insert orphan headings with deliberately pre-existing out-of-order pPr children
      const headingSpecs = [
        { style: 'Heading1', text: 'Section 1: Executive Overview' },
        { style: 'Heading 2', text: 'Section 1.1: Financial Performance' },
        { style: 'Heading3', text: 'Section 1.1.1: Detailed Breakdowns' },
        { style: 'Title', text: 'Global Enterprise Audit Report' }
      ];

      const sectPr = body.getElementsByTagName('w:sectPr')[0];

      for (const spec of headingSpecs) {
        const p = createWordElement(dom, 'w:p');
        const pPr = createWordElement(dom, 'w:pPr');
        
        // Intentionally insert w:jc before w:pStyle to stress schema sorting
        const jc = createWordElement(dom, 'w:jc');
        jc.setAttribute('w:val', 'left');
        pPr.appendChild(jc);

        const pStyle = createWordElement(dom, 'w:pStyle');
        pStyle.setAttribute('w:val', spec.style);
        pPr.appendChild(pStyle);

        p.appendChild(pPr);

        const r = createWordElement(dom, 'w:r');
        const t = createWordElement(dom, 'w:t');
        t.appendChild(dom.createTextNode(spec.text));
        r.appendChild(t);
        p.appendChild(r);

        if (sectPr) body.insertBefore(p, sectPr);
        else body.appendChild(p);
      }

      // Insert 10 tables, including multiline rows and a nested table in Table 3
      const tables: Element[] = [];
      for (let i = 0; i < 10; i++) {
        const hasNested = (i === 2); // Table 3 has a nested table
        const tbl = createTestTable(dom, 4, 3, { multiline: true, nestedTable: hasNested });
        if (sectPr) body.insertBefore(tbl, sectPr);
        else body.appendChild(tbl);
        tables.push(tbl);
      }

      // Execute typographic linter
      const lintRes = lintAndFixTypography(dom);
      expect(lintRes.applied).toBe(true);
      expect(lintRes.fixes.length).toBeGreaterThan(0);

      // Verify <w:keepNext> on ALL headings and verify PPR_CHILD_ORDER strictly
      const allParagraphs = Array.from(dom.getElementsByTagName('w:p'));
      const headingParagraphs = allParagraphs.filter(p => {
        const pPr = p.getElementsByTagName('w:pPr')[0];
        if (!pPr) return false;
        const pStyle = pPr.getElementsByTagName('w:pStyle')[0];
        if (!pStyle) return false;
        const val = pStyle.getAttribute('w:val') || '';
        return /heading\s*\d|heading\d|title/i.test(val);
      });

      expect(headingParagraphs.length).toBe(6);
      for (const h of headingParagraphs) {
        const pPr = h.getElementsByTagName('w:pPr')[0];
        expect(pPr).toBeDefined();

        const keepNext = pPr.getElementsByTagName('w:keepNext');
        expect(keepNext.length, 'Every heading paragraph must contain <w:keepNext>').toBe(1);

        // Verify that pPr child elements match PPR_CHILD_ORDER strictly
        assertStrictChildOrder(pPr, PPR_CHILD_ORDER);
      }

      // Also specifically assert that our 4 inserted headings have keepNext
      for (const spec of headingSpecs) {
        const found = headingParagraphs.find(p => p.textContent?.includes(spec.text));
        expect(found, `Inserted heading "${spec.text}" must be present`).toBeDefined();
        const pPr = found!.getElementsByTagName('w:pPr')[0];
        expect(pPr.getElementsByTagName('w:keepNext').length).toBe(1);
      }

      // Verify <w:cantSplit> is on all rows and <w:tblHeader> is on first rows
      const allTables = Array.from(dom.getElementsByTagName('w:tbl'));
      // 10 top-level tables + 1 nested table = 11 tables total
      expect(allTables.length).toBe(11);

      for (const tbl of allTables) {
        const rows = Array.from(tbl.childNodes).filter(
          (c): c is Element => c.nodeType === 1 && (c as Element).tagName === 'w:tr'
        );
        expect(rows.length).toBeGreaterThan(0);

        for (let r = 0; r < rows.length; r++) {
          const row = rows[r];
          const trPr = row.getElementsByTagName('w:trPr')[0];
          expect(trPr, `Row ${r} must have w:trPr`).toBeDefined();

          const cantSplit = trPr.getElementsByTagName('w:cantSplit');
          expect(cantSplit.length, `Row ${r} must have w:cantSplit`).toBe(1);

          const tblHeader = trPr.getElementsByTagName('w:tblHeader');
          if (r === 0) {
            expect(tblHeader.length, 'First row of table must have w:tblHeader').toBe(1);
          } else {
            expect(tblHeader.length, `Row ${r} must NOT have w:tblHeader`).toBe(0);
          }

          // Verify TRPR_CHILD_ORDER is 100% compliant with ECMA-376
          assertStrictChildOrder(trPr, TRPR_CHILD_ORDER);
        }
      }

      // Save document and validate with validateDocxBuffer: MUST pass with ok: true
      const savePath = path.join(tmpDir, 'brutal_typography_10_tables.docx');
      await dm.saveAsDocument(savePath);
      expect(fs.existsSync(savePath)).toBe(true);

      const savedBuf = fs.readFileSync(savePath);
      const valRes = await validateDocxBuffer(savedBuf);
      expect(valRes.ok, `Docx validation must pass: ${JSON.stringify(valRes.errors)}`).toBe(true);
    });

    it('verifies idempotency: running lintAndFixTypography twice produces zero duplicate tags and zero new fixes', async () => {
      const dm = new DocumentManager();
      await dm.createNewDocument('report');

      const dom = dm.getParser().getDocumentDom();

      // First run: fixes applied
      const firstRun = lintAndFixTypography(dom);
      expect(firstRun.applied).toBe(true);
      expect(firstRun.fixes.length).toBeGreaterThan(0);

      // Second run: MUST NOT apply any new fixes
      const secondRun = lintAndFixTypography(dom);
      expect(secondRun.applied).toBe(false);
      expect(secondRun.fixes.length).toBe(0);

      // Verify no duplicate tags exist in any pPr or trPr
      const pPrNodes = dom.getElementsByTagName('w:pPr');
      for (let i = 0; i < pPrNodes.length; i++) {
        const keepNext = pPrNodes[i].getElementsByTagName('w:keepNext');
        expect(keepNext.length).toBeLessThanOrEqual(1);
        const widowControl = pPrNodes[i].getElementsByTagName('w:widowControl');
        expect(widowControl.length).toBeLessThanOrEqual(1);
      }

      const trPrNodes = dom.getElementsByTagName('w:trPr');
      for (let i = 0; i < trPrNodes.length; i++) {
        const cantSplit = trPrNodes[i].getElementsByTagName('w:cantSplit');
        expect(cantSplit.length).toBeLessThanOrEqual(1);
        const tblHeader = trPrNodes[i].getElementsByTagName('w:tblHeader');
        expect(tblHeader.length).toBeLessThanOrEqual(1);
      }
    });

    it('respects applyTypography: false opt-out in DocumentManager without altering the document', async () => {
      const dm = new DocumentManager();
      await dm.createNewDocument('blank');

      const savePath = path.join(tmpDir, 'opt_out.docx');
      const saveRes = await dm.saveAsDocument(savePath, { applyTypography: false });
      expect(saveRes.typographyFixes.length).toBe(0);

      const savedBuf = fs.readFileSync(savePath);
      const validation = await validateDocxBuffer(savedBuf);
      expect(validation.ok).toBe(true);

      // Verify DOM was not modified with keepNext or cantSplit
      const dom = dm.getParser().getDocumentDom();
      const cantSplitNodes = dom.getElementsByTagName('w:cantSplit');
      expect(cantSplitNodes.length).toBe(0);
    });
  });

  describe('MOAT-3: Word Redlining & Track Changes', () => {
    it('executes tracked insert and delete operations producing valid w:ins and w:del markup', async () => {
      const dm = new DocumentManager();
      await dm.createNewDocument('blank');

      const parser = dm.getParser();
      const ast = await parser.getAST();
      const firstSection = ast.sections[0];
      const targetNodeId = firstSection.nodes[0].id;

      const engine = new OperationEngine(ast, parser);

      // 1. Perform tracked change insert
      const insertRes = await engine.executeBatch({
        operations: [
          {
            action: 'trackedChange',
            operation: 'insert',
            target: targetNodeId,
            text: ' Tracked insertion text.',
            author: 'Audit Reviewer',
            date: '2026-09-07T12:00:00.000Z'
          }
        ],
        sequential: true
      });
      expect(insertRes.success).toBe(true);

      // Verify DOM has w:ins with correct attributes
      const dom = parser.getDocumentDom();
      const insNodes = dom.getElementsByTagName('w:ins');
      expect(insNodes.length).toBeGreaterThan(0);
      const insEl = insNodes[0];
      expect(insEl.getAttribute('w:author')).toBe('Audit Reviewer');
      expect(insEl.getAttribute('w:date')).toBe('2026-09-07T12:00:00.000Z');
      expect(insEl.textContent).toContain('Tracked insertion text.');

      // 2. Perform tracked change delete
      const deleteRes = await engine.executeBatch({
        operations: [
          {
            action: 'trackedChange',
            operation: 'delete',
            target: targetNodeId,
            author: 'Audit Reviewer',
            date: '2026-09-07T12:05:00.000Z'
          }
        ],
        sequential: true
      });
      expect(deleteRes.success).toBe(true);

      const delNodes = dom.getElementsByTagName('w:del');
      expect(delNodes.length).toBeGreaterThan(0);
      const delEl = delNodes[0];
      expect(delEl.getAttribute('w:author')).toBe('Audit Reviewer');

      // 3. Save and verify validator passes with ok: true
      const savePath = path.join(tmpDir, 'tracked_changes.docx');
      await dm.saveAsDocument(savePath);

      const savedBuf = fs.readFileSync(savePath);
      const valRes = await validateDocxBuffer(savedBuf);
      expect(valRes.ok).toBe(true);
    });

    it('executes multiple sequential tracked changes (insert, delete) with unique non-colliding integer IDs and validates docx', async () => {
      const dm = new DocumentManager();
      await dm.createNewDocument('blank');

      const parser = dm.getParser();
      let ast = await parser.getAST();
      const engine = new OperationEngine(ast, parser);

      // First add several paragraph nodes to manipulate
      const addRes = await engine.executeBatch({
        operations: [
          { action: 'insert', type: 'paragraph', content: 'Initial paragraph alpha.' },
          { action: 'insert', type: 'paragraph', content: 'Initial paragraph beta.' },
          { action: 'insert', type: 'paragraph', content: 'Initial paragraph gamma.' }
        ],
        sequential: true
      });
      expect(addRes.success).toBe(true);

      // Re-read AST to get minted node IDs
      ast = await parser.getAST();
      const section = ast.sections[0];
      const pNodes = section.nodes.filter(n => n.type === 'paragraph');
      expect(pNodes.length).toBeGreaterThanOrEqual(3);

      const id0 = pNodes[0].id;
      const id1 = pNodes[1].id;
      const id2 = pNodes[2].id;

      // Execute 6 sequential tracked changes: insert, delete, insert, delete, insert, delete
      const trackedOps = [
        {
          action: 'trackedChange' as const,
          operation: 'insert' as const,
          target: id0,
          text: ' [Tracked Insert 1]',
          author: 'Senior Auditor',
          date: '2026-09-07T12:10:00.000Z'
        },
        {
          action: 'trackedChange' as const,
          operation: 'delete' as const,
          target: id1,
          author: 'Senior Auditor',
          date: '2026-09-07T12:11:00.000Z'
        },
        {
          action: 'trackedChange' as const,
          operation: 'insert' as const,
          target: id2,
          text: ' [Tracked Insert 2]',
          author: 'Peer Reviewer',
          date: '2026-09-07T12:12:00.000Z'
        },
        {
          action: 'trackedChange' as const,
          operation: 'delete' as const,
          target: id0,
          author: 'Peer Reviewer',
          date: '2026-09-07T12:13:00.000Z'
        },
        {
          action: 'trackedChange' as const,
          operation: 'insert' as const,
          target: id2,
          text: ' [Tracked Insert 3]',
          author: 'Compliance Officer',
          date: '2026-09-07T12:14:00.000Z'
        },
        {
          action: 'trackedChange' as const,
          operation: 'delete' as const,
          target: id2,
          author: 'Compliance Officer',
          date: '2026-09-07T12:15:00.000Z'
        }
      ];

      for (const op of trackedOps) {
        const opRes = await engine.executeBatch({
          operations: [op],
          sequential: true
        });
        expect(opRes.success, `Tracked operation ${op.operation} failed`).toBe(true);
      }

      const dom = parser.getDocumentDom();

      // Collect all tracked change elements (w:ins and w:del)
      const insEls = Array.from(dom.getElementsByTagName('w:ins'));
      const delEls = Array.from(dom.getElementsByTagName('w:del'));
      const allChanges = [...insEls, ...delEls];
      expect(allChanges.length).toBe(6);

      // Verify ID generation: IDs must be unique positive integers without collision
      const allocatedIds: string[] = [];
      for (const ch of allChanges) {
        const changeId = ch.getAttribute('w:id');
        expect(changeId).toBeDefined();
        expect(changeId).toMatch(/^\d+$/);
        const intId = parseInt(changeId!, 10);
        expect(intId).toBeGreaterThan(0);
        allocatedIds.push(changeId!);
      }

      // Assert zero collisions in allocated IDs
      const uniqueIds = new Set(allocatedIds);
      expect(uniqueIds.size, 'Every tracked change w:id must be strictly unique').toBe(allocatedIds.length);

      // Verify deleteTrackedChange creates <w:del> with <w:r><w:delText xml:space="preserve">...</w:delText></w:r>
      for (const del of delEls) {
        expect(del.getAttribute('w:author')).toBeDefined();
        expect(del.getAttribute('w:date')).toBeDefined();
        const runs = del.getElementsByTagName('w:r');
        expect(runs.length).toBeGreaterThan(0);
        const delText = runs[0].getElementsByTagName('w:delText');
        expect(delText.length).toBeGreaterThan(0);
        expect(delText[0].getAttribute('xml:space')).toBe('preserve');
      }

      // Verify insertTrackedChange creates <w:ins> with <w:r><w:t xml:space="preserve">...</w:t></w:r>
      for (const ins of insEls) {
        expect(ins.getAttribute('w:author')).toBeDefined();
        expect(ins.getAttribute('w:date')).toBeDefined();
        const runs = ins.getElementsByTagName('w:r');
        expect(runs.length).toBeGreaterThan(0);
        const tNodes = runs[0].getElementsByTagName('w:t');
        expect(tNodes.length).toBeGreaterThan(0);
        expect(tNodes[0].getAttribute('xml:space')).toBe('preserve');
      }

      // Validate modified document with validateDocxBuffer: MUST pass with ok: true
      const savePath = path.join(tmpDir, 'brutal_sequential_tracked_changes.docx');
      await dm.saveAsDocument(savePath);
      expect(fs.existsSync(savePath)).toBe(true);

      const savedBuf = fs.readFileSync(savePath);
      const valRes = await validateDocxBuffer(savedBuf);
      expect(valRes.ok, `Validation failed: ${JSON.stringify(valRes.errors)}`).toBe(true);
    });
  });

  describe('MOAT-0: Zero Raw createElement in Canonical Homes', () => {
    it('verifies 100% of w: element creations in packages/engine/src use createWordElement', () => {
      const canonicalDirs = [path.resolve('packages/engine/src')];
      const filesWithRawWordCreate: string[] = [];

      function scanDir(dir: string) {
        const entries = fs.readdirSync(dir, { withFileTypes: true });
        for (const entry of entries) {
          const fullPath = path.join(dir, entry.name);
          if (entry.isDirectory()) {
            scanDir(fullPath);
          } else if (entry.isFile() && entry.name.endsWith('.ts')) {
            const content = fs.readFileSync(fullPath, 'utf-8');
            if (content.includes(".createElement('w:") || content.includes('.createElement("w:')) {
              filesWithRawWordCreate.push(fullPath);
            }
          }
        }
      }

      for (const dir of canonicalDirs) {
        expect(fs.existsSync(dir), `canonical home missing: ${dir}`).toBe(true);
        scanDir(dir);
      }
      expect(filesWithRawWordCreate).toEqual([]);
    });
  });
});

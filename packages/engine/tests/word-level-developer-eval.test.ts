/**
 * Word Control (word-level span styling) — Comprehensive Developer Evaluation Suite
 *
 * Evaluates:
 * 1. TypeScript API & Type Ergonomics (StyleOperation vs Operation)
 * 2. Pre-existing complex formatting preservation on run splitting
 * 3. Handling of tabs (w:tab) and line breaks (w:br)
 * 4. Mixed paragraph-level (pPr) and word-level (rPr) styling in a single operation
 * 5. Multiple consecutive styling operations on the same paragraph
 * 6. OOXML schema conformance via validateOOXML()
 * 7. TOON fidelity across tiers (compact, standard, full)
 */
import { describe, it, expect } from 'vitest';
import { XMLSerializer } from '@xmldom/xmldom';
import { DocxParser, createBlankDocxBuffer } from '../src/parser';
import { OperationEngine } from '../src/operation-engine';
import { StyleOperation } from '../src/operations/style';
import { Operation, validateOperationStrict } from '../src/operation-schema';
import { ToonGenerator } from '../src/toon';
import { PPR_CHILD_ORDER, RPR_CHILD_ORDER } from '../src/ooxml/schema-order';
import type { Node } from '../src/ast';

const ser = new XMLSerializer();

async function freshDoc(title = 'Dev Eval Test') {
  const parser = new DocxParser(createBlankDocxBuffer(title));
  const ast = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  return { parser, ast, engine };
}

function lastNode(ast: { sections: Array<{ nodes: Node[] }> }): Node {
  const nodes = ast.sections[0].nodes;
  return nodes[nodes.length - 1];
}

function runText(run: Element): string {
  let out = '';
  const kids = run.childNodes;
  for (let i = 0; i < kids.length; i++) {
    const c = kids[i] as unknown as Element;
    if (!c || c.nodeType !== 1) continue;
    if (c.tagName === 'w:t') out += c.textContent || '';
    else if (c.tagName === 'w:tab') out += '\t';
    else if (c.tagName === 'w:br') out += '\n';
  }
  return out;
}

function runsOf(node: Node): Array<{ text: string; xml: string; el: Element }> {
  const els = Array.from(node.xml.getElementsByTagName('w:r')) as unknown as Element[];
  return els.map((r) => ({ text: runText(r), xml: ser.serializeToString(r), el: r }));
}

function paraText(node: Node): string {
  return runsOf(node).map((r) => r.text).join('');
}

async function insertParagraph(engine: OperationEngine, ast: { sections: Array<{ nodes: Node[] }> }, text: string): Promise<Node> {
  const anchor = lastNode(ast).id;
  await engine.execute({ action: 'insert', type: 'paragraph', anchor, position: 'after', text } as never);
  return lastNode(ast);
}

async function insertXml(engine: OperationEngine, ast: { sections: Array<{ nodes: Node[] }> }, xml: string): Promise<Node> {
  const anchor = lastNode(ast).id;
  await engine.execute({ action: 'insert', type: 'xml', anchor, position: 'after', xml } as never);
  return lastNode(ast);
}

function getRPrChildTagNames(runEl: Element): string[] {
  const rPr = runEl.getElementsByTagName('w:rPr')[0];
  if (!rPr) return [];
  const tags: string[] = [];
  for (let i = 0; i < rPr.childNodes.length; i++) {
    const c = rPr.childNodes[i] as Element;
    if (c && c.nodeType === 1) tags.push(c.tagName);
  }
  return tags;
}

function getPPrChildTagNames(pEl: Element): string[] {
  const pPr = pEl.getElementsByTagName('w:pPr')[0];
  if (!pPr) return [];
  const tags: string[] = [];
  for (let i = 0; i < pPr.childNodes.length; i++) {
    const c = pPr.childNodes[i] as Element;
    if (c && c.nodeType === 1) tags.push(c.tagName);
  }
  return tags;
}

describe('Word Control Developer Evaluation', () => {

  // =========================================================================
  // 1. Developer API & Type Ergonomics
  // =========================================================================
  describe('1. Developer API & Type Ergonomics', () => {
    it('evaluates Operation type ergonomics (discriminated union from operation-schema)', () => {
      // Testing with Operation type:
      const op: Operation = {
        action: 'style',
        target: 'p-1',
        match: 'Riverbend',
        matchAll: true,
        matchCase: false,
        color: '047857',
        bold: true,
        fontSize: '12pt',
      };

      expect(op.action).toBe('style');
      if (op.action === 'style') {
        // Here match, matchAll, matchCase are fully typed as:
        // match?: string | undefined
        // matchAll?: boolean | undefined
        // matchCase?: boolean | undefined
        const m: string | undefined = op.match;
        const ma: boolean | undefined = op.matchAll;
        const mc: boolean | undefined = op.matchCase;
        expect(m).toBe('Riverbend');
        expect(ma).toBe(true);
        expect(mc).toBe(false);
      }
    });

    it('exposes match precisely on StyleOperation', () => {
      // StyleOperation extends BaseOperation with StyleProperties, so flat
      // match/matchAll/matchCase carry precise types.
      const styleOp: StyleOperation = {
        action: 'style',
        target: 'p-1',
        match: 'Riverbend',
        matchAll: true,
        matchCase: true,
      };
      const m: string | undefined = styleOp.match;
      const ma: boolean | undefined = styleOp.matchAll;
      const mc: boolean | undefined = styleOp.matchCase;
      expect([m, ma, mc]).toEqual(['Riverbend', true, true]);

      // The nested properties bag stays typed as well.
      const nested: StyleOperation = {
        action: 'style',
        target: 'p-1',
        properties: {
          match: 'Riverbend',
          matchAll: true,
          matchCase: true,
        },
      };
      const propMatch: string | undefined = nested.properties?.match;
      expect(propMatch).toBe('Riverbend');
    });

    it('verifies schema gate validation with validateOperationStrict', () => {
      // Valid flat word-style operation
      const valid = validateOperationStrict({
        action: 'style',
        target: 'p-1',
        match: 'Riverbend',
        color: '047857',
        bold: true,
      });
      expect(valid.ok).toBe(true);

      // Invalid empty match
      expect(() => {
        validateOperationStrict({
          action: 'style',
          target: 'p-1',
          match: '',
          bold: true,
        });
      }).toThrow(/E_VALIDATION/);

      // Unknown property rejected by strict schema gate
      expect(() => {
        validateOperationStrict({
          action: 'style',
          target: 'p-1',
          match: 'Riverbend',
          unknownTypoProperty: true,
        });
      }).toThrow(/unrecognized_keys/);
    });
  });

  // =========================================================================
  // 2. Empirical Engine Testing: Pre-existing Complex Formatting
  // =========================================================================
  describe('2. Pre-existing Complex Formatting Preservation', () => {
    it('preserves existing w:rPr (font size, style, font family, color) when splitting a formatted run', async () => {
      const { parser, ast, engine } = await freshDoc();

      // Create a paragraph with 3 pre-existing formatted runs:
      // Run 1: "Leading text " (Times New Roman, 24 hps / 12pt, italic)
      // Run 2: "PreExistingStyledWord" (Calibri, 16 hps / 8pt, bold, color 112233)
      // Run 3: " trailing text" (Arial, 20 hps / 10pt, underline)
      const xml = 
        '<w:p>' +
          '<w:r>' +
            '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>' +
            '<w:t xml:space="preserve">Leading text </w:t>' +
          '</w:r>' +
          '<w:r>' +
            '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="112233"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>' +
            '<w:t>PreExistingStyledWord</w:t>' +
          '</w:r>' +
          '<w:r>' +
            '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:u w:val="single"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>' +
            '<w:t xml:space="preserve"> trailing text</w:t>' +
          '</w:r>' +
        '</w:p>';

      const para = await insertXml(engine, ast, xml);
      const runsBefore = runsOf(para);
      expect(runsBefore.length).toBe(3);

      // Apply word styling to "Existing" within "PreExistingStyledWord":
      // Add highlight 'yellow' and strikethrough: true
      await engine.execute({
        action: 'style',
        target: para.id,
        match: 'Existing',
        highlight: 'yellow',
        strikethrough: true,
      } as never);

      const runsAfter = runsOf(para);
      // Run 2 should be split into 3 segments: "Pre", "Existing", "StyledWord"
      // Total runs: Run 1 + 3 segments + Run 3 = 5 runs
      expect(runsAfter.length).toBe(5);
      expect(runsAfter.map(r => r.text)).toEqual([
        'Leading text ',
        'Pre',
        'Existing',
        'StyledWord',
        ' trailing text',
      ]);

      // Check Run 0 (untouched):
      expect(runsAfter[0].xml).toContain('w:ascii="Times New Roman"');
      expect(runsAfter[0].xml).toContain('<w:i/>');
      expect(runsAfter[0].xml).toContain('w:val="24"');

      // Check Segment "Pre" (untouched prefix of Run 2):
      // Must preserve Calibri, bold, color 112233, sz 16
      expect(runsAfter[1].xml).toContain('w:ascii="Calibri"');
      expect(runsAfter[1].xml).toContain('<w:b/>');
      expect(runsAfter[1].xml).toContain('w:val="112233"');
      expect(runsAfter[1].xml).toContain('w:val="16"');
      expect(runsAfter[1].xml).not.toContain('w:highlight');
      expect(runsAfter[1].xml).not.toContain('<w:strike/>');

      // Check Segment "Existing" (matched segment):
      // Must inherit Calibri, bold, color 112233, sz 16 AND gain highlight yellow and strike!
      expect(runsAfter[2].xml).toContain('w:ascii="Calibri"');
      expect(runsAfter[2].xml).toContain('<w:b/>');
      expect(runsAfter[2].xml).toContain('w:val="112233"');
      expect(runsAfter[2].xml).toContain('w:val="16"');
      expect(runsAfter[2].xml).toContain('w:val="yellow"');
      expect(runsAfter[2].xml).toContain('<w:strike/>');

      // Check Segment "StyledWord" (untouched suffix of Run 2):
      expect(runsAfter[3].xml).toContain('w:ascii="Calibri"');
      expect(runsAfter[3].xml).toContain('<w:b/>');
      expect(runsAfter[3].xml).toContain('w:val="112233"');
      expect(runsAfter[3].xml).toContain('w:val="16"');
      expect(runsAfter[3].xml).not.toContain('w:highlight');
      expect(runsAfter[3].xml).not.toContain('<w:strike/>');

      // Check Run 4 (untouched trailing):
      expect(runsAfter[4].xml).toContain('w:ascii="Arial"');
      expect(runsAfter[4].xml).toContain('w:val="single"');
      expect(runsAfter[4].xml).toContain('w:val="20"');

      // Validate OOXML conformance
      const valResult = await parser.validateOOXML();
      expect(valResult.ok).toBe(true);
      expect(valResult.errors).toEqual([]);
    });

    it('preserves existing formatting across a cross-run span', async () => {
      const { parser, ast, engine } = await freshDoc();

      // Run 1: "Green " (color 00FF00)
      // Run 2: "Blue" (color 0000FF)
      const xml = 
        '<w:p>' +
          '<w:r>' +
            '<w:rPr><w:color w:val="00FF00"/></w:rPr>' +
            '<w:t xml:space="preserve">Green </w:t>' +
          '</w:r>' +
          '<w:r>' +
            '<w:rPr><w:color w:val="0000FF"/></w:rPr>' +
            '<w:t>Blue</w:t>' +
          '</w:r>' +
        '</w:p>';

      const para = await insertXml(engine, ast, xml);

      // Match "een Bl": spans tail of Run 1 and head of Run 2
      await engine.execute({
        action: 'style',
        target: para.id,
        match: 'een Bl',
        bold: true,
      } as never);

      const runs = runsOf(para);
      expect(runs.map(r => r.text)).toEqual(['Gr', 'een ', 'Bl', 'ue']);

      // 'Gr' should have color 00FF00, NOT bold
      expect(runs[0].xml).toContain('w:val="00FF00"');
      expect(runs[0].xml).not.toContain('<w:b/>');

      // 'een ' should have color 00FF00 AND bold
      expect(runs[1].xml).toContain('w:val="00FF00"');
      expect(runs[1].xml).toContain('<w:b/>');

      // 'Bl' should have color 0000FF AND bold
      expect(runs[2].xml).toContain('w:val="0000FF"');
      expect(runs[2].xml).toContain('<w:b/>');

      // 'ue' should have color 0000FF, NOT bold
      expect(runs[3].xml).toContain('w:val="0000FF"');
      expect(runs[3].xml).not.toContain('<w:b/>');

      const valResult = await parser.validateOOXML();
      expect(valResult.ok).toBe(true);
    });
  });

  // =========================================================================
  // 3. Empirical Engine Testing: Tabs (w:tab) and Line Breaks (w:br)
  // =========================================================================
  describe('3. Paragraphs Containing Tabs and Line Breaks', () => {
    it('accurately counts tabs and breaks so word offsets align perfectly', async () => {
      const { parser, ast, engine } = await freshDoc();

      // Structure:
      // Run with: "Col1" + <w:tab/> + "Col2" + <w:br/> + "Row2"
      const xml = 
        '<w:p>' +
          '<w:r>' +
            '<w:t>Col1</w:t>' +
            '<w:tab/>' +
            '<w:t>Col2</w:t>' +
            '<w:br/>' +
            '<w:t>Row2</w:t>' +
          '</w:r>' +
        '</w:p>';

      const para = await insertXml(engine, ast, xml);

      // Overall text is "Col1\tCol2\nRow2"
      expect(paraText(para)).toBe('Col1\tCol2\nRow2');

      // Test 1: Style "Col2" (comes after \t)
      await engine.execute({
        action: 'style',
        target: para.id,
        match: 'Col2',
        color: 'FF0000',
      } as never);

      let runs = runsOf(para);
      expect(runs.map(r => r.text)).toEqual(['Col1\t', 'Col2', '\nRow2']);
      expect(runs[1].xml).toContain('w:val="FF0000"');
      expect(runs[0].xml).not.toContain('w:val="FF0000"');
      expect(runs[2].xml).not.toContain('w:val="FF0000"');
      expect(runs[0].xml).toContain('<w:tab/>');
      expect(runs[2].xml).toContain('<w:br/>');

      // Test 2: Style "Row2" (comes after \n)
      await engine.execute({
        action: 'style',
        target: para.id,
        match: 'Row2',
        bold: true,
      } as never);

      runs = runsOf(para);
      expect(runs.map(r => r.text)).toEqual(['Col1\t', 'Col2', '\n', 'Row2']);
      expect(runs[3].xml).toContain('<w:b/>');
      expect(runs[2].xml).toContain('<w:br/>');
      expect(runs[2].xml).not.toContain('<w:b/>');

      // Validate OOXML
      const valResult = await parser.validateOOXML();
      expect(valResult.ok).toBe(true);
    });

    it('handles matching a substring that spans across a tab character', async () => {
      const { parser, ast, engine } = await freshDoc();

      const xml = 
        '<w:p>' +
          '<w:r>' +
            '<w:t>Left</w:t>' +
            '<w:tab/>' +
            '<w:t>Right</w:t>' +
          '</w:r>' +
        '</w:p>';

      const para = await insertXml(engine, ast, xml);

      // Match "ft\tRi"
      await engine.execute({
        action: 'style',
        target: para.id,
        match: 'ft\tRi',
        italic: true,
      } as never);

      const runs = runsOf(para);
      expect(runs.map(r => r.text)).toEqual(['Le', 'ft\tRi', 'ght']);
      expect(runs[1].xml).toContain('<w:i/>');
      expect(runs[1].xml).toContain('<w:tab/>');
      expect(runs[0].xml).not.toContain('<w:i/>');
      expect(runs[2].xml).not.toContain('<w:i/>');

      const valResult = await parser.validateOOXML();
      expect(valResult.ok).toBe(true);
    });

    it('handles matching a substring that spans across a line break', async () => {
      const { parser, ast, engine } = await freshDoc();

      const xml = 
        '<w:p>' +
          '<w:r>' +
            '<w:t>Line1</w:t>' +
            '<w:br/>' +
            '<w:t>Line2</w:t>' +
          '</w:r>' +
        '</w:p>';

      const para = await insertXml(engine, ast, xml);

      // Match "e1\nLi"
      await engine.execute({
        action: 'style',
        target: para.id,
        match: 'e1\nLi',
        underline: true,
      } as never);

      const runs = runsOf(para);
      expect(runs.map(r => r.text)).toEqual(['Lin', 'e1\nLi', 'ne2']);
      expect(runs[1].xml).toContain('<w:u/>');
      expect(runs[1].xml).toContain('<w:br/>');
      expect(runs[0].xml).not.toContain('<w:u/>');
      expect(runs[2].xml).not.toContain('<w:u/>');

      const valResult = await parser.validateOOXML();
      expect(valResult.ok).toBe(true);
    });
  });

  // =========================================================================
  // 4. Empirical Engine Testing: Paragraph + Run Styling in Same Op
  // =========================================================================
  describe('4. Paragraph and Word-Level Styling in the Same Operation', () => {
    it('applies paragraph properties paragraph-wide while run properties apply strictly to matched word', async () => {
      const { parser, ast, engine } = await freshDoc();
      const para = await insertParagraph(engine, ast, 'Prefix HighlightMe Suffix');

      // Combined op: paragraph props (alignment, spacingBefore) + word prop (color)
      await engine.execute({
        action: 'style',
        target: para.id,
        match: 'HighlightMe',
        color: 'FF0000',
        alignment: 'center',
        spacingBefore: 240,
        spacingAfter: 120,
      } as never);

      // 1. Check paragraph properties in w:pPr
      const pPr = para.xml.getElementsByTagName('w:pPr')[0];
      expect(pPr).toBeDefined();

      const jc = pPr.getElementsByTagName('w:jc')[0];
      expect(jc).toBeDefined();
      expect(jc.getAttribute('w:val')).toBe('center');

      const spacing = pPr.getElementsByTagName('w:spacing')[0];
      expect(spacing).toBeDefined();
      expect(spacing.getAttribute('w:before')).toBe('240');
      expect(spacing.getAttribute('w:after')).toBe('120');

      // 2. Check run splitting and run properties
      const runs = runsOf(para);
      expect(runs.map(r => r.text)).toEqual(['Prefix ', 'HighlightMe', ' Suffix']);

      // Run 0 ('Prefix '): NO color
      expect(runs[0].xml).not.toContain('w:color');

      // Run 1 ('HighlightMe'): color FF0000
      expect(runs[1].xml).toContain('w:val="FF0000"');

      // Run 2 (' Suffix'): NO color
      expect(runs[2].xml).not.toContain('w:color');

      // Validate OOXML conformance
      const valResult = await parser.validateOOXML();
      expect(valResult.ok).toBe(true);
      expect(valResult.errors).toEqual([]);
    });

    it('validates schema ordering in w:pPr when mixed styling is applied', async () => {
      const { parser, ast, engine } = await freshDoc();
      const para = await insertParagraph(engine, ast, 'The quick brown fox');

      await engine.execute({
        action: 'style',
        target: para.id,
        match: 'brown',
        color: '795548',
        alignment: 'right',
        spacingBefore: 200,
        indentLeft: 720,
      } as never);

      const pPrTags = getPPrChildTagNames(para.xml);
      // In PPR_CHILD_ORDER: w:spacing, w:ind, w:jc
      const jcIndex = pPrTags.indexOf('w:jc');
      const spacingIndex = pPrTags.indexOf('w:spacing');
      const indIndex = pPrTags.indexOf('w:ind');

      expect(spacingIndex).toBeGreaterThanOrEqual(0);
      expect(indIndex).toBeGreaterThanOrEqual(0);
      expect(jcIndex).toBeGreaterThanOrEqual(0);

      // Verify schema ordering: w:spacing precedes w:ind precedes w:jc in ECMA-376
      const orderSpacing = PPR_CHILD_ORDER.indexOf('w:spacing');
      const orderInd = PPR_CHILD_ORDER.indexOf('w:ind');
      const orderJc = PPR_CHILD_ORDER.indexOf('w:jc');

      expect(orderSpacing).toBeLessThan(orderInd);
      expect(orderInd).toBeLessThan(orderJc);

      expect(spacingIndex).toBeLessThan(indIndex);
      expect(indIndex).toBeLessThan(jcIndex);

      const valResult = await parser.validateOOXML();
      expect(valResult.ok).toBe(true);
    });
  });

  // =========================================================================
  // 5. Empirical Engine Testing: Multiple Consecutive Operations
  // =========================================================================
  describe('5. Multiple Consecutive Styling Operations on the Same Paragraph', () => {
    it('survives consecutive operations without index corruption or offset drift (left-to-right)', async () => {
      const { parser, ast, engine } = await freshDoc();
      const para = await insertParagraph(engine, ast, 'One Two Three Four Five');

      // Step 1: Style "Two"
      await engine.execute({
        action: 'style',
        target: para.id,
        match: 'Two',
        bold: true,
        color: 'FF0000',
      } as never);

      expect(runsOf(para).map(r => r.text)).toEqual(['One ', 'Two', ' Three Four Five']);

      // Step 2: Style "Four" (after "Two")
      await engine.execute({
        action: 'style',
        target: para.id,
        match: 'Four',
        italic: true,
        color: '0000FF',
      } as never);

      expect(runsOf(para).map(r => r.text)).toEqual([
        'One ',
        'Two',
        ' Three ',
        'Four',
        ' Five',
      ]);

      // Step 3: Style "Three" (in between "Two" and "Four")
      await engine.execute({
        action: 'style',
        target: para.id,
        match: 'Three',
        underline: true,
        color: '00FF00',
      } as never);

      const runs = runsOf(para);
      expect(runs.map(r => r.text)).toEqual([
        'One ',
        'Two',
        ' ',
        'Three',
        ' ',
        'Four',
        ' Five',
      ]);

      // Verify formatting on each segment
      expect(runs[1].xml).toContain('<w:b/>');
      expect(runs[1].xml).toContain('w:val="FF0000"');

      expect(runs[3].xml).toContain('<w:u/>');
      expect(runs[3].xml).toContain('w:val="00FF00"');

      expect(runs[5].xml).toContain('<w:i/>');
      expect(runs[5].xml).toContain('w:val="0000FF"');

      // Verify overall paragraph text is preserved exactly
      expect(paraText(para)).toBe('One Two Three Four Five');

      const valResult = await parser.validateOOXML();
      expect(valResult.ok).toBe(true);
    });

    it('survives consecutive operations styled in reverse order (right-to-left)', async () => {
      const { parser, ast, engine } = await freshDoc();
      const para = await insertParagraph(engine, ast, 'Apple Banana Cherry Date');

      // Step 1: Style "Date" (rightmost)
      await engine.execute({ action: 'style', target: para.id, match: 'Date', bold: true } as never);

      // Step 2: Style "Cherry"
      await engine.execute({ action: 'style', target: para.id, match: 'Cherry', italic: true } as never);

      // Step 3: Style "Banana"
      await engine.execute({ action: 'style', target: para.id, match: 'Banana', underline: true } as never);

      // Step 4: Style "Apple" (leftmost)
      await engine.execute({ action: 'style', target: para.id, match: 'Apple', highlight: 'yellow' } as never);

      const runs = runsOf(para);
      expect(runs.map(r => r.text)).toEqual([
        'Apple',
        ' ',
        'Banana',
        ' ',
        'Cherry',
        ' ',
        'Date',
      ]);

      expect(runs[0].xml).toContain('w:val="yellow"');
      expect(runs[2].xml).toContain('<w:u/>');
      expect(runs[4].xml).toContain('<w:i/>');
      expect(runs[6].xml).toContain('<w:b/>');

      expect(paraText(para)).toBe('Apple Banana Cherry Date');

      const valResult = await parser.validateOOXML();
      expect(valResult.ok).toBe(true);
    });

    it('styles a sub-word within an already-styled word run', async () => {
      const { parser, ast, engine } = await freshDoc();
      const para = await insertParagraph(engine, ast, 'Supercalifragilistic');

      // First style the whole word bold + color blue
      await engine.execute({
        action: 'style',
        target: para.id,
        match: 'Supercalifragilistic',
        bold: true,
        color: '0000FF',
      } as never);

      // Then style "cali" inside it to italic + red
      await engine.execute({
        action: 'style',
        target: para.id,
        match: 'cali',
        italic: true,
        color: 'FF0000',
      } as never);

      const runs = runsOf(para);
      expect(runs.map(r => r.text)).toEqual(['Super', 'cali', 'fragilistic']);

      // 'Super': bold, color 0000FF, NOT italic
      expect(runs[0].xml).toContain('<w:b/>');
      expect(runs[0].xml).toContain('w:val="0000FF"');
      expect(runs[0].xml).not.toContain('<w:i/>');

      // 'cali': bold (inherited), italic (added), color FF0000 (overridden)
      expect(runs[1].xml).toContain('<w:b/>');
      expect(runs[1].xml).toContain('<w:i/>');
      expect(runs[1].xml).toContain('w:val="FF0000"');

      // 'fragilistic': bold, color 0000FF, NOT italic
      expect(runs[2].xml).toContain('<w:b/>');
      expect(runs[2].xml).toContain('w:val="0000FF"');
      expect(runs[2].xml).not.toContain('<w:i/>');

      const valResult = await parser.validateOOXML();
      expect(valResult.ok).toBe(true);
    });
  });

  // =========================================================================
  // 6. OOXML & TOON Fidelity
  // =========================================================================
  describe('6. OOXML & TOON Fidelity', () => {
    it('verifies RPR schema ordering on all generated split runs', async () => {
      const { parser, ast, engine } = await freshDoc();
      const para = await insertParagraph(engine, ast, 'Ordering check run');

      // Apply multiple run properties
      await engine.execute({
        action: 'style',
        target: para.id,
        match: 'check',
        fontFamily: 'Georgia',
        bold: true,
        italic: true,
        underline: true,
        color: '123456',
        fontSize: '14pt',
        highlight: 'yellow',
        vertAlign: 'superscript',
      } as never);

      const runs = runsOf(para);
      const styledRun = runs[1];
      expect(styledRun.text).toBe('check');

      const tags = getRPrChildTagNames(styledRun.el);

      // Verify each tag in tags appears in ascending order according to RPR_CHILD_ORDER
      let lastIndex = -1;
      for (const tag of tags) {
        const schemaIdx = RPR_CHILD_ORDER.indexOf(tag);
        expect(schemaIdx).toBeGreaterThanOrEqual(0);
        expect(schemaIdx).toBeGreaterThan(lastIndex);
        lastIndex = schemaIdx;
      }

      const valResult = await parser.validateOOXML();
      expect(valResult.ok).toBe(true);
    });

    it('verifies TOON output representation before and after word styling across tiers', async () => {
      const { ast, engine } = await freshDoc();
      const para = await insertParagraph(engine, ast, 'ASTra word styling in action');

      const toonGen = new ToonGenerator();

      // Check standard tier before styling
      const standardBefore = toonGen.toon(ast, { fidelity: 'standard' });
      const rowBefore = standardBefore.rows.find(r => r.id === para.id);
      expect(rowBefore).toBeDefined();
      expect(rowBefore?.text).toBe('ASTra word styling in action');
      expect(rowBefore?.flags).not.toContain('[B]');
      expect(rowBefore?.flags).not.toContain('[C:');

      // Apply word styling
      await engine.execute({
        action: 'style',
        target: para.id,
        match: 'word styling',
        bold: true,
        color: 'FF0000',
      } as never);

      // 1. Check compact tier
      const compactAfter = toonGen.toon(ast, { fidelity: 'compact' });
      const rowCompact = compactAfter.rows.find(r => r.id === para.id);
      expect(rowCompact?.text).toBe('ASTra word styling in action');
      expect(rowCompact?.flags).not.toContain('[B]'); // compact drops formatting flags

      // 2. Check standard tier
      const standardAfter = toonGen.toon(ast, { fidelity: 'standard' });
      const rowStandard = standardAfter.rows.find(r => r.id === para.id);
      expect(rowStandard?.text).toBe('ASTra word styling in action'); // Readable text preserved without duplicates!
      expect(rowStandard?.flags).toContain('[B]');
      expect(rowStandard?.flags).toContain('[C:FF0000]');

      // 3. Check full tier
      const fullAfter = toonGen.toon(ast, { fidelity: 'full', sectionId: ast.sections[0].id });
      const rowFull = fullAfter.rows.find(r => r.id === para.id);
      expect(rowFull?.text).toBe('ASTra word styling in action');
      expect(rowFull?.detail).toBeDefined();

      // Full tier detail includes rPr dumps for all runs
      expect(rowFull?.detail).toContain('rPr');
      expect(rowFull?.detail).toContain('<w:b/>');
      expect(rowFull?.detail).toContain('w:val="FF0000"');
    });

    it('validates entire DOCX buffer via parser.validateOOXML() after complete workflow', async () => {
      const { parser, ast, engine } = await freshDoc('Full Workflow Validation');
      
      const p1 = await insertParagraph(engine, ast, 'First paragraph with some text');
      const p2 = await insertParagraph(engine, ast, 'Second paragraph with Tab\tSeparated\nValues');

      await engine.execute({
        action: 'style',
        target: p1.id,
        match: 'some text',
        bold: true,
        italic: true,
        color: '2E7D32',
      } as never);

      await engine.execute({
        action: 'style',
        target: p2.id,
        match: 'Separated',
        highlight: 'cyan',
        alignment: 'center',
        spacingBefore: 180,
      } as never);

      const validation = await parser.validateOOXML();
      expect(validation.ok).toBe(true);
      expect(validation.errors).toEqual([]);
    });
  });
});

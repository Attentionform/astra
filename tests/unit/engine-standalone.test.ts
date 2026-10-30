import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import {
  DocxParser,
  createBlankDocxBuffer,
  OperationEngine,
  lintAndFixTypography,
  extractStylePalette,
  ToonGenerator,
  sanitizeDocxBuffer,
  validateDocxBuffer,
} from '@attentionform/astra-engine';

describe('@attentionform/astra-engine Headless Standalone Engine Suite', () => {
  it('verifies zero Electron runtime dependency', () => {
    // Standalone engine must never require electron
    expect(process.versions.electron).toBeUndefined();
    expect(process.env.ELECTRON_RUN_AS_NODE).toBeUndefined();

    // Verify no electron modules leaked into require cache
    const loadedModules = Object.keys((typeof require !== 'undefined' && require.cache) || {});
    const electronLoaded = loadedModules.some(
      (mod) => mod.includes('/electron/') || mod.endsWith('/electron')
    );
    expect(electronLoaded).toBe(false);
  });

  it('executes complete headless OOXML lifecycle: create, insert paragraphs, tables, 3D charts with XLSX workbooks, lint typography, extract styles, generate tiered TOON, sanitize metadata, and validate schema', async () => {
    // 1. Create a fresh document in-memory
    const initialBuffer = createBlankDocxBuffer('Standalone OOXML Report');
    const parser = new DocxParser(initialBuffer);
    let ast = await parser.getAST();
    expect(ast.sections.length).toBeGreaterThan(0);

    const engine = new OperationEngine(ast, parser);

    // 2. Insert paragraphs (Title, H1, Body)
    const anchorId = ast.sections[0].nodes[0]?.id || 'p_001';

    await engine.execute({
      action: 'insert',
      type: 'paragraph',
      position: 'after',
      anchor: anchorId,
      text: 'Executive Summary',
    });

    await engine.execute({
      action: 'insert',
      type: 'paragraph',
      position: 'after',
      anchor: anchorId,
      text: 'This document was authored entirely via the headless @attentionform/astra-engine standalone package.',
    });

    // Refresh AST
    ast = await parser.getAST();
    expect(ast.sections[0].nodes.length).toBeGreaterThanOrEqual(3);

    // 3. Insert a structured table with headers and rows
    const preTableLastNode = ast.sections[0].nodes[ast.sections[0].nodes.length - 1];
    await engine.execute({
      action: 'insert',
      type: 'table',
      position: 'after',
      anchor: preTableLastNode.id,
      headers: ['Metric', 'Baseline', 'Target', 'Status'],
      rows: [
        ['Cold Start (ms)', '850', '200', 'Achieved (120ms)'],
        ['Memory Footprint', '350MB', '< 100MB', 'Achieved (42MB)'],
        ['OOXML Schema Compliance', '98.5%', '100%', 'Verified (100%)'],
        ['Headless Portability', 'Electron Only', 'Pure Node.js', 'Complete']
      ]
    });

    // Refresh AST
    ast = await parser.getAST();
    const hasTable = ast.sections.some(sec => sec.nodes.some(n => n.type === 'table'));
    expect(hasTable).toBe(true);

    // 4. Insert DrawingML 3D Chart with embedded real XLSX workbook binary
    const lastNode = ast.sections[0].nodes[ast.sections[0].nodes.length - 1];
    await engine.execute({
      action: 'chart',
      operation: 'insert',
      chartType: 'bar3D',
      title: 'Headless Engine Performance Benchmarks',
      view3D: { rotX: 30, rotY: 45, perspective: 30, rAngAx: true },
      shape: 'cylinder',
      categories: ['Parse', 'Mutate', 'Lint', 'Validate'],
      series: [
        { name: 'ASTra v1 (Electron)', values: [120, 85, 45, 95], color: 'EF4444' },
        { name: '@attentionform/astra-engine (Pure)', values: [32, 18, 12, 28], color: '10B981' }
      ],
      anchor: lastNode.id,
      position: 'after'
    });

    // 5. Inspect OOXML package zip parts for chart and embedded workbook
    const zip = parser.getZipAdapter();
    expect(zip.hasFile('word/charts/chart1.xml')).toBe(true);
    expect(zip.hasFile('word/embeddings/Microsoft_Excel_Worksheet1.xlsx')).toBe(true);
    expect(zip.hasFile('word/charts/_rels/chart1.xml.rels')).toBe(true);

    // Verify embedded Excel binary content
    const docxZip = new AdmZip(parser.rezip());
    const xlsxEntry = docxZip.getEntry('word/embeddings/Microsoft_Excel_Worksheet1.xlsx');
    expect(xlsxEntry).not.toBeNull();
    const xlsxBuf = xlsxEntry!.getData();
    expect(xlsxBuf.length).toBeGreaterThan(500);
    const excelZip = new AdmZip(xlsxBuf);
    expect(excelZip.getEntry('[Content_Types].xml')).not.toBeNull();
    expect(excelZip.getEntry('xl/workbook.xml')).not.toBeNull();

    // Verify chartSpace XML contains 3D controls and externalData rel
    const chartXml = zip.readText('word/charts/chart1.xml');
    expect(chartXml).toContain('<c:view3D>');
    expect(chartXml).toContain('<c:bar3DChart>');
    expect(chartXml).toContain('<c:shape val="cylinder"/>');
    expect(chartXml).toContain('<c:externalData');

    // 6. Run typographic linter & schema guard
    const lintRes = lintAndFixTypography(parser.getDom(), {
      keepNextHeadings: true,
      cantSplitRows: true,
      repeatTableHeader: true,
    });
    expect(lintRes.applied).toBe(true);
    expect(Array.isArray(lintRes.fixes)).toBe(true);

    // 7. Extract style palette
    const palette = parser.getStylePalette();
    expect(palette).toBeDefined();
    expect(typeof palette.primaryFont).toBe('string');
    expect(Array.isArray(palette.brandColors)).toBe(true);
    expect(Array.isArray(palette.availableStyles)).toBe(true);

    if (parser.getZipAdapter().hasFile('word/styles.xml')) {
      const directPalette = extractStylePalette(parser.getZipAdapter().readText('word/styles.xml'));
      expect(directPalette).toBeDefined();
    }

    // 8. Generate tiered TOON (Token-Optimized Outline Notation)
    ast = await parser.getAST();
    const toonGen = new ToonGenerator();
    const compactTable = toonGen.toon(ast, { fidelity: 'compact' });
    const standardTable = toonGen.toon(ast, { fidelity: 'standard' });
    const fullTable = toonGen.toon(ast, { fidelity: 'full', sectionId: ast.sections[0].id });

    const compactText = toonGen.renderToonText(compactTable);
    const standardText = toonGen.renderToonText(standardTable);
    const fullText = toonGen.renderToonText(fullTable);

    expect(typeof compactText).toBe('string');
    expect(typeof standardText).toBe('string');
    expect(typeof fullText).toBe('string');
    expect(compactTable.tier).toBe('compact');
    expect(standardTable.tier).toBe('standard');
    expect(fullTable.tier).toBe('full');
    expect(standardText).toContain('Executive Summary');

    // 9. Sanitize document metadata
    const rawDocxBuf = parser.rezip();
    const { buffer: sanitizedBuffer, report } = sanitizeDocxBuffer(rawDocxBuf);
    expect(sanitizedBuffer).toBeDefined();
    expect(sanitizedBuffer.length).toBeGreaterThan(1000);
    expect(report.ok).toBe(true);

    // 10. Strict Schema Validation Gate: validateDocxBuffer must be ok: true
    const validation = await validateDocxBuffer(sanitizedBuffer);
    expect(validation.ok).toBe(true);
    expect(validation.errors).toHaveLength(0);
  });
});

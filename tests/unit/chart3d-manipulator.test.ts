import { describe, it, expect, beforeEach } from 'vitest';
import { DOMParser } from '@xmldom/xmldom';
import AdmZip from 'adm-zip';
import {
  Chart3DManipulator,
  Chart3DType,
  DocumentAST,
  DocxParser,
  SUPPORTED_CHART_TYPES,
  executeChart,
  generateEmbeddedWorkbook,
  is3DChartType,
  validateDocxBuffer,
} from '@attentionform/astra-engine';

type Equal<A, B> = (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2 ? true : false;
type ExpectTrue<T extends true> = T;
/** Compile-time pin: the requestable 3D union is exactly the runtime list. */
export type Chart3DUnionMatchesRuntime = ExpectTrue<Equal<Chart3DType, (typeof SUPPORTED_CHART_TYPES)[number]>>;

describe('Chart3DManipulator Unit Suite', () => {
  let ast: DocumentAST;
  let domParser: DOMParser;

  beforeEach(() => {
    domParser = new DOMParser();
    const mockDoc = domParser.parseFromString(
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p data-astra-id="p_001">
            <w:r><w:t>Baseline paragraph for chart anchor</w:t></w:r>
          </w:p>
        </w:body>
      </w:document>`,
      'text/xml'
    );

    const bodyP = mockDoc.getElementsByTagName('w:p')[0];

    ast = {
      metadata: {},
      sections: [{
        id: 'sec_root',
        heading: null,
        nodes: [{
          id: 'p_001',
          type: 'paragraph',
          xml: bodyP,
          metadata: {
            text: 'Baseline paragraph for chart anchor',
            hasCitation: false,
            hasComment: false,
            hasBookmark: false,
            hasTrackedChanges: false,
            hasEquation: false,
            hasField: false,
            hasImage: false,
            hasHyperlink: false
          },
          version: 1,
          checksum: '111'
        }],
        pageSetup: null,
        version: 1,
        checksum: '222'
      }],
      styles: {},
      numbering: { abstractNum: {}, concreteNum: {} },
      relationships: {},
      headers: {},
      footers: {},
      mediaFiles: {},
      settings: { trackChanges: true, defaultAuthor: 'ASTra' },
      nextChangeId: 1
    };
  });

  it('should generate ISO/IEC 29500-1 Clause 21.2 compliant 3D Bar Chart XML with c:view3D controls', () => {
    const manipulator = new Chart3DManipulator(ast);
    const chartXml = manipulator.generate3DChartXml({
      chartType: 'bar3D',
      title: 'Quarterly Revenue 3D',
      view3D: { rotX: 30, rotY: 45, perspective: 40, rAngAx: true },
      shape: 'cylinder',
      categories: ['Q1', 'Q2', 'Q3', 'Q4'],
      series: [{ name: 'Revenue', values: [10, 20, 30, 40], color: '0066CC' }]
    });

    expect(chartXml).toContain('<c:chartSpace');
    expect(chartXml).toContain('<c:view3D>');
    expect(chartXml).toContain('<c:rotX val="30"/>');
    expect(chartXml).toContain('<c:rotY val="45"/>');
    expect(chartXml).toContain('<c:perspective val="40"/>');
    expect(chartXml).toContain('<c:bar3DChart>');
    expect(chartXml).toContain('<c:shape val="cylinder"/>');
    expect(chartXml).toContain('Quarterly Revenue 3D');
  });

  it('should generate 3D Pie Chart XML with valid c:pie3DChart structure', () => {
    const manipulator = new Chart3DManipulator(ast);
    const chartXml = manipulator.generate3DChartXml({
      chartType: 'pie3D',
      title: 'Market Share 3D',
      view3D: { rotX: 25, rotY: 10, perspective: 30 },
      categories: ['Product A', 'Product B'],
      series: [{ name: 'Share', values: [60, 40] }]
    });

    expect(chartXml).toContain('<c:pie3DChart>');
    expect(chartXml).toContain('<c:varyColors val="1"/>');
    expect(chartXml).toContain('<c:v>Product A</c:v>');
    expect(chartXml).toContain('<c:v>60</c:v>');
  });

  it('should generate ISO/IEC 29500-1 Clause 21.2 compliant 3D Line Chart XML with c:serAx and validate in document', async () => {
    const manipulator = new Chart3DManipulator(ast);
    const chartXml = manipulator.generate3DChartXml({
      chartType: 'line3D',
      title: 'Growth Trajectory 3D',
      view3D: { rotX: 20, rotY: 30, perspective: 30 },
      categories: ['Jan', 'Feb', 'Mar', 'Apr'],
      series: [{ name: 'Growth', values: [10, 25, 40, 75], color: '10B981' }]
    });

    expect(chartXml).toContain('<c:line3DChart>');
    expect(chartXml).toContain('<c:serAx>');
    expect(chartXml).toContain('<c:axId val="1003"/>');

    const parser = DocxParser.createNewDocument('blank');
    const docAst = await parser.getAST();
    const docManipulator = new Chart3DManipulator(docAst, parser);

    const result = docManipulator.insert3DChart({
      chartType: 'line3D',
      title: 'Growth Trajectory 3D',
      categories: ['Jan', 'Feb', 'Mar', 'Apr'],
      series: [{ name: 'Growth', values: [10, 25, 40, 75], color: '10B981' }]
    });

    expect(result.chartPath).toBe('word/charts/chart1.xml');
    const docxBuf = parser.rezip();

    const bufValidation = await validateDocxBuffer(docxBuf);
    expect(bufValidation.ok).toBe(true);
    expect(bufValidation.errors).toEqual([]);

    const ooxmlValidation = await parser.validateOOXML();
    expect(ooxmlValidation.ok).toBe(true);
    expect(ooxmlValidation.errors).toEqual([]);
  });

  it('should update 3D perspective controls on existing chart XML', () => {
    const manipulator = new Chart3DManipulator(ast);
    const initialXml = manipulator.generate3DChartXml({
      chartType: 'bar3D',
      view3D: { rotX: 10, rotY: 10, perspective: 20 }
    });

    const updatedXml = manipulator.updateView3D(initialXml, {
      rotX: 45,
      rotY: 90,
      perspective: 50,
      hPercent: 120
    });

    expect(updatedXml).toContain('<c:rotX val="45"/>');
    expect(updatedXml).toContain('<c:rotY val="90"/>');
    expect(updatedXml).toContain('<c:perspective val="50"/>');
    expect(updatedXml).toContain('<c:hPercent val="120"/>');
  });

  it('should generate DrawingML inline embedding element with EMU dimensions', () => {
    const manipulator = new Chart3DManipulator(ast);
    const embedXml = manipulator.generateDrawingEmbedXml('rIdC1', 1, 'Chart 1', 5486400, 3200400);

    expect(embedXml).toContain('<w:drawing>');
    expect(embedXml).toContain('<wp:extent cx="5486400" cy="3200400"/>');
    expect(embedXml).toContain('<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">');
    expect(embedXml).toContain('<c:chart r:id="rIdC1"/>');
  });

  describe('MOAT-4: Chart Type Honesty & Error Contract', () => {
    it('should throw structured E_VALIDATION error when unsupported chart type area3D is passed to generate3DChartXml', () => {
      const manipulator = new Chart3DManipulator(ast);
      try {
        manipulator.generate3DChartXml({
          chartType: 'area3D' as any,
          title: 'Area 3D Test'
        });
        expect.unreachable('Should have thrown an unsupported chart type error');
      } catch (err: any) {
        expect(err).toBeInstanceOf(Error);
        expect(err.message).toContain('Unsupported chart type "area3D"');
        expect(err.message).toContain('Supported types: bar3D, line3D, pie3D, column3D');
        expect(err.code).toBe('E_VALIDATION');
        expect(err.retryable).toBe(false);
        expect(err.remedy).toContain('Use one of the supported chart types');
      }
    });

    it('should throw structured E_VALIDATION error when unsupported chart type surface3D is passed to insert3DChart', () => {
      const manipulator = new Chart3DManipulator(ast);
      try {
        manipulator.insert3DChart({
          chartType: 'surface3D' as any,
          title: 'Surface 3D Test'
        });
        expect.unreachable('Should have thrown an unsupported chart type error');
      } catch (err: any) {
        expect(err).toBeInstanceOf(Error);
        expect(err.message).toContain('Unsupported chart type "surface3D"');
        expect(err.code).toBe('E_VALIDATION');
        expect(err.retryable).toBe(false);
        expect(err.remedy).toContain('Use one of the supported chart types');
      }
    });

    it('should throw structured E_VALIDATION error when executeChart is called with unsupported chart type', () => {
      try {
        executeChart(
          {
            action: 'chart',
            chartType: 'area3D' as any,
            title: 'Unsupported Area 3D'
          },
          ast
        );
        expect.unreachable('Should have thrown');
      } catch (err: any) {
        expect(err).toBeInstanceOf(Error);
        expect(err.message).toContain('Unsupported chart type "area3D"');
        expect(err.code).toBe('E_VALIDATION');
        expect(err.retryable).toBe(false);
        expect(err.remedy).toContain('Use one of the supported chart types');
      }
    });
  });

  describe('MOAT-4: Embedded XLSX Workbook & OOXML Validation', () => {
    it('should generate valid SpreadsheetML workbook buffer containing series names, categories, and values', () => {
      const series = [
        { name: 'Revenue', values: [100.5, 200.25], color: '0066CC' },
        { name: 'Expenses', values: [80.1, 150.75], color: 'CC0000' }
      ];
      const categories = ['Q1 2026', 'Q2 2026'];

      const workbookBuf = generateEmbeddedWorkbook(series, categories);
      expect(workbookBuf).toBeInstanceOf(Buffer);
      expect(workbookBuf.length).toBeGreaterThan(0);

      const zip = new AdmZip(workbookBuf);
      expect(zip.getEntry('[Content_Types].xml')).not.toBeNull();
      expect(zip.getEntry('_rels/.rels')).not.toBeNull();
      expect(zip.getEntry('xl/workbook.xml')).not.toBeNull();
      expect(zip.getEntry('xl/_rels/workbook.xml.rels')).not.toBeNull();
      expect(zip.getEntry('xl/styles.xml')).not.toBeNull();
      expect(zip.getEntry('xl/worksheets/sheet1.xml')).not.toBeNull();

      const sheetXml = zip.getEntry('xl/worksheets/sheet1.xml')?.getData().toString('utf8');
      expect(sheetXml).toContain('Revenue');
      expect(sheetXml).toContain('Expenses');
      expect(sheetXml).toContain('Q1 2026');
      expect(sheetXml).toContain('Q2 2026');
      expect(sheetXml).toContain('100.5');
      expect(sheetXml).toContain('200.25');
      expect(sheetXml).toContain('80.1');
      expect(sheetXml).toContain('150.75');
    });

    it('should embed workbook, wire relationships, append c:externalData, and pass validateOOXML and validateDocxBuffer', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const docAst = await parser.getAST();
      const manipulator = new Chart3DManipulator(docAst, parser);

      const result = manipulator.insert3DChart({
        chartType: 'column3D',
        title: '3D Enterprise Growth',
        categories: ['2023', '2024', '2025', '2026'],
        series: [
          { name: 'ARR', values: [10, 25, 50, 100], color: '10B981' },
          { name: 'Headcount', values: [15, 30, 60, 120], color: '6366F1' }
        ]
      });

      expect(result.chartPath).toBe('word/charts/chart1.xml');
      expect(result.workbookPath).toBe('word/embeddings/Microsoft_Excel_Worksheet1.xlsx');
      expect(result.drawingNodeId).toBeDefined();

      const docxBuf = parser.rezip();
      const zip = new AdmZip(docxBuf);

      // Assert workbook file exists in zip
      expect(zip.getEntry('word/embeddings/Microsoft_Excel_Worksheet1.xlsx')).not.toBeNull();

      // Assert chart rels points to embedded workbook
      const chartRelsXml = zip.getEntry('word/charts/_rels/chart1.xml.rels')?.getData().toString('utf8');
      expect(chartRelsXml).toBeDefined();
      expect(chartRelsXml).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"');
      expect(chartRelsXml).toContain('Target="../embeddings/Microsoft_Excel_Worksheet1.xlsx"');

      // Assert chart.xml contains c:externalData referencing rId1 per ECMA-376 §21.2
      const chartXml = zip.getEntry('word/charts/chart1.xml')?.getData().toString('utf8');
      expect(chartXml).toContain('<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>');

      // Assert [Content_Types].xml contains override and Default for xlsx
      const contentTypesXml = zip.getEntry('[Content_Types].xml')?.getData().toString('utf8');
      expect(contentTypesXml).toContain('PartName="/word/embeddings/Microsoft_Excel_Worksheet1.xlsx"');
      expect(contentTypesXml).toContain('ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"');
      expect(contentTypesXml).toContain('Extension="xlsx"');

      // Assert validateDocxBuffer returns ok: true
      const bufValidation = await validateDocxBuffer(docxBuf);
      expect(bufValidation.ok).toBe(true);
      expect(bufValidation.errors).toEqual([]);

      // Assert validateOOXML returns ok: true
      const ooxmlValidation = await parser.validateOOXML();
      expect(ooxmlValidation.ok).toBe(true);
      expect(ooxmlValidation.errors).toEqual([]);
    });
  });

  it('assigns distinct palette fills per uncolored series; explicit wins', () => {
    const manipulator = new Chart3DManipulator(ast);
    const xml = manipulator.generate3DChartXml({
      chartType: 'bar3D',
      title: 'Two series',
      categories: ['A', 'B'],
      series: [
        { name: 'Planned', values: [1, 2] },
        { name: 'Completed', values: [3, 4] },
      ],
    });
    const fills = [...xml.matchAll(/<c:ser>[\s\S]*?<a:srgbClr val="([0-9A-Fa-f]{6})"/g)].map((m) => m[1].toUpperCase());
    expect(fills).toEqual(['0066CC', '10B981']);

    const mixed = manipulator.generate3DChartXml({
      chartType: 'bar3D',
      title: 'Mixed',
      categories: ['A', 'B'],
      series: [
        { name: 'Custom', values: [1, 2], color: 'FF0000' },
        { name: 'Defaulted', values: [3, 4] },
      ],
    });
    const mixedFills = [...mixed.matchAll(/<c:ser>[\s\S]*?<a:srgbClr val="([0-9A-Fa-f]{6})"/g)].map((m) => m[1].toUpperCase());
    expect(mixedFills).toEqual(['FF0000', '10B981']);
  });
});

describe('Chart3DType truth (union = runtime support)', () => {
  it('names exactly the four runtime-supported 3D types', () => {
    expect([...SUPPORTED_CHART_TYPES]).toEqual(['bar3D', 'line3D', 'pie3D', 'column3D']);
    for (const chartType of SUPPORTED_CHART_TYPES) expect(is3DChartType(chartType)).toBe(true);
    for (const foreign of ['area3D', 'surface3D']) expect(is3DChartType(foreign)).toBe(false);
  });
});

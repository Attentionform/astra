import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import {
  DocxParser,
  DocumentAST,
  executeChart,
  executeChartUpdate,
  validateOperationStrict,
  OperationSchema,
  OperationEngine,
  validateDocxBuffer,
} from '@attentionform/astra-engine';

function chartXmlFromBuffer(buf: Buffer, name = 'chart1.xml'): string {
  const zip = new AdmZip(buf);
  const entry = zip.getEntry(`word/charts/${name}`);
  if (!entry) throw new Error(`${name} missing from rezipped docx`);
  return entry.getData().toString('utf8');
}

function listChartEntriesFromBuffer(buf: Buffer): string[] {
  const zip = new AdmZip(buf);
  return zip.getEntries()
    .map((e) => e.entryName)
    .filter((name) => name.startsWith('word/charts/chart') && name.endsWith('.xml'));
}

function readWorkbookSheetFromBuffer(buf: Buffer, workbookName = 'Microsoft_Excel_Worksheet1.xlsx'): string {
  const zip = new AdmZip(buf);
  const wbEntry = zip.getEntry(`word/embeddings/${workbookName}`);
  if (!wbEntry) throw new Error(`word/embeddings/${workbookName} missing from docx`);
  const wbZip = new AdmZip(wbEntry.getData());
  const sheetEntry = wbZip.getEntry('xl/worksheets/sheet1.xml');
  if (!sheetEntry) throw new Error('xl/worksheets/sheet1.xml missing from embedded workbook');
  return sheetEntry.getData().toString('utf8');
}

describe('Wave 1C: Chart in-place update without duplication (P0-4)', () => {
  describe('Strict Operation Schema Gate (chartOp target & params validation)', () => {
    it('accepts flat target and params in chartOp for update', () => {
      const validFlat = {
        action: 'chart',
        operation: 'update',
        target: 'chart_node_1',
        title: 'Q3 Financials',
        categories: ['Jul', 'Aug', 'Sep'],
        series: [{ name: 'Revenue', values: [100, 200, 300], color: '10B981' }],
      };
      const flatRes = validateOperationStrict(validFlat);
      expect(flatRes.ok).toBe(true);
      expect(flatRes.operation).toMatchObject(validFlat);

      const validNested = {
        action: 'chart',
        operation: 'update',
        params: {
          target: 'chart_node_1',
          title: 'Nested Title',
          series: [{ name: 'EBITDA', values: [50, 75] }],
        },
      };
      const nestedRes = validateOperationStrict(validNested);
      expect(nestedRes.ok).toBe(true);
    });

    it('rejects empty string target or invalid series at schema gate with E_VALIDATION', () => {
      expect(() =>
        validateOperationStrict({
          action: 'chart',
          operation: 'update',
          target: '',
          title: 'Empty target',
        })
      ).toThrow(/E_VALIDATION/);

      expect(() =>
        validateOperationStrict({
          action: 'chart',
          operation: 'update',
          target: 'chart_1',
          series: [{ name: 'Bad', values: ['not-a-number' as unknown as number] }],
        })
      ).toThrow(/E_VALIDATION/);

      expect(
        OperationSchema.safeParse({
          action: 'chart',
          operation: 'update',
          target: 'chart_1',
          unknownExtraKey: 123,
        }).success
      ).toBe(false);
    });
  });

  describe('2D Chart In-Place Update', () => {
    it('update without series colors assigns the palette and keeps existing colors', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();
      executeChart(
        {
          action: 'chart',
          chartType: 'column',
          title: 'Palette update',
          categories: ['A', 'B'],
          series: [{ name: 'One', values: [1, 2], color: 'FF0000' }],
          id: 'chart_pal_node',
        },
        ast,
        parser
      );
      // New uncolored series take palette entries; the kept series keeps red.
      executeChart(
        {
          action: 'chart',
          operation: 'update',
          target: 'chart_pal_node',
          series: [
            { name: 'One', values: [1, 2] },
            { name: 'Two', values: [3, 4] },
          ],
        },
        ast,
        parser
      );
      const updatedXml = chartXmlFromBuffer(parser.rezip(), 'chart1.xml');
      const fills = [...updatedXml.matchAll(/<c:ser>[\s\S]*?<a:srgbClr val="([0-9A-Fa-f]{6})"/g)].map((m) => m[1].toUpperCase());
      expect(fills).toEqual(['FF0000', '10B981']);
    });

    it('updates 2D column chart in-place: title, categories, series, workbook, AST metadata, without part duplication', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();

      // 1. Insert initial 2D chart
      executeChart(
        {
          action: 'chart',
          chartType: 'column',
          title: 'Original 2D Revenue',
          categories: ['Jan', 'Feb', 'Mar'],
          series: [{ name: 'Sales', values: [10, 20, 30], color: '0066CC' }],
          id: 'chart_2d_node',
        },
        ast,
        parser
      );

      const initialChartNode = ast.sections[0].nodes.find((n) => n.id === 'chart_2d_node');
      expect(initialChartNode).toBeDefined();
      expect(initialChartNode!.type).toBe('chart');
      expect(initialChartNode!.metadata.title).toBe('Original 2D Revenue');
      expect(initialChartNode!.version).toBe(1);
      const initialChecksum = initialChartNode!.checksum;

      const bufBefore = parser.rezip();
      const chartPartsBefore = listChartEntriesFromBuffer(bufBefore);
      expect(chartPartsBefore).toEqual(['word/charts/chart1.xml']);

      // 2. Perform in-place update with target
      executeChart(
        {
          action: 'chart',
          operation: 'update',
          target: 'chart_2d_node',
          title: 'Updated 2D Growth Analysis',
          categories: ['Q1', 'Q2', 'Q3', 'Q4'],
          series: [
            { name: 'ARR', values: [100, 250, 500, 1000], color: '10B981' },
            { name: 'Headcount', values: [15, 30, 60, 120], color: '6366F1' },
          ],
        },
        ast,
        parser
      );

      // 3. Verify single chart part in zip (NO chart2.xml created)
      const bufAfter = parser.rezip();
      const chartPartsAfter = listChartEntriesFromBuffer(bufAfter);
      expect(chartPartsAfter).toEqual(['word/charts/chart1.xml']);
      expect(chartPartsAfter.length).toBe(1);

      // 4. Verify word/charts/chart1.xml updated content
      const updatedXml = chartXmlFromBuffer(bufAfter, 'chart1.xml');
      expect(updatedXml).toContain('<c:barChart>');
      expect(updatedXml).toContain('<c:barDir val="col"/>');
      expect(updatedXml).not.toContain('3DChart');
      expect(updatedXml).toContain('Updated 2D Growth Analysis');
      expect(updatedXml).not.toContain('Original 2D Revenue');
      expect(updatedXml).toContain('<c:v>Q1</c:v>');
      expect(updatedXml).toContain('<c:v>Q4</c:v>');
      expect(updatedXml).toContain('<c:v>ARR</c:v>');
      expect(updatedXml).toContain('<c:v>Headcount</c:v>');
      expect(updatedXml).toContain('<c:v>1000</c:v>');
      expect(updatedXml).toContain('<c:v>120</c:v>');
      expect(updatedXml).toContain('<a:srgbClr val="10B981"/>');
      expect(updatedXml).toContain('<a:srgbClr val="6366F1"/>');

      // 5. Verify embedded workbook updated with new series and categories
      const sheetXml = readWorkbookSheetFromBuffer(bufAfter, 'Microsoft_Excel_Worksheet1.xlsx');
      expect(sheetXml).toContain('ARR');
      expect(sheetXml).toContain('Headcount');
      expect(sheetXml).toContain('Q1');
      expect(sheetXml).toContain('Q4');
      expect(sheetXml).toContain('<v>1000</v>');
      expect(sheetXml).toContain('<v>120</v>');

      // 6. Verify AST node updated in-place (same node object mutated, version incremented, checksum updated)
      expect(initialChartNode!.metadata.title).toBe('Updated 2D Growth Analysis');
      expect(initialChartNode!.metadata.text).toBe('Updated 2D Growth Analysis');
      expect(initialChartNode!.metadata.categories).toEqual(['Q1', 'Q2', 'Q3', 'Q4']);
      expect(initialChartNode!.metadata.chartType).toBe('column');
      expect(initialChartNode!.type).toBe('chart');
      expect(initialChartNode!.version).toBe(2);
      expect(initialChartNode!.checksum).not.toBe(initialChecksum);

      // 7. Verify OOXML and DocxBuffer validation passes
      const docxVal = await validateDocxBuffer(bufAfter);
      expect(docxVal.ok).toBe(true);
      expect(docxVal.errors).toEqual([]);

      const ooxmlVal = await parser.validateOOXML();
      expect(ooxmlVal.ok).toBe(true);
      expect(ooxmlVal.errors).toEqual([]);
    });

    it('preserves existing categories and series when fields are omitted during update', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();

      executeChart(
        {
          action: 'chart',
          chartType: 'bar',
          title: 'Original Bar Chart',
          categories: ['Alpha', 'Beta', 'Gamma'],
          series: [{ name: 'Throughput', values: [12, 24, 36], color: 'F59E0B' }],
          id: 'chart_bar_node',
        },
        ast,
        parser
      );

      // Update title only, omitting categories and series
      executeChart(
        {
          action: 'chart',
          operation: 'update',
          target: 'chart_bar_node',
          title: 'Renamed Bar Chart Only',
        },
        ast,
        parser
      );

      const buf = parser.rezip();
      const updatedXml = chartXmlFromBuffer(buf, 'chart1.xml');
      expect(updatedXml).toContain('Renamed Bar Chart Only');
      expect(updatedXml).toContain('<c:v>Alpha</c:v>');
      expect(updatedXml).toContain('<c:v>Beta</c:v>');
      expect(updatedXml).toContain('<c:v>Gamma</c:v>');
      expect(updatedXml).toContain('<c:v>Throughput</c:v>');
      expect(updatedXml).toContain('<c:v>36</c:v>');

      const node = ast.sections[0].nodes.find((n) => n.id === 'chart_bar_node')!;
      expect(node.metadata.title).toBe('Renamed Bar Chart Only');
      expect(node.metadata.categories).toEqual(['Alpha', 'Beta', 'Gamma']);
      expect(node.version).toBe(2);
    });
  });

  describe('3D Chart In-Place Update', () => {
    it('updates 3D bar chart in-place with new shape, view3D, categories, and series', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();

      executeChart(
        {
          action: 'chart',
          chartType: 'bar3D',
          title: 'Original 3D Bar',
          shape: 'box',
          categories: ['Stage 1', 'Stage 2'],
          series: [{ name: 'Latency', values: [50, 80], color: '2563EB' }],
          id: 'chart_3d_node',
        },
        ast,
        parser
      );

      const chartNode = ast.sections[0].nodes.find((n) => n.id === 'chart_3d_node')!;
      expect(chartNode.type).toBe('chart3d');
      expect(chartNode.version).toBe(1);

      // Perform update: change shape to cone, update series and categories, set view3D
      executeChart(
        {
          action: 'chart',
          operation: 'update',
          target: 'chart_3d_node',
          title: '3D Cylinder Analysis',
          shape: 'cylinder',
          categories: ['Phase A', 'Phase B', 'Phase C'],
          series: [
            { name: 'Optimization', values: [15, 35, 75], color: '10B981' },
          ],
          view3D: { rotX: 25, rotY: 40, perspective: 35 },
        },
        ast,
        parser
      );

      const buf = parser.rezip();
      const chartParts = listChartEntriesFromBuffer(buf);
      expect(chartParts).toEqual(['word/charts/chart1.xml']);

      const updatedXml = chartXmlFromBuffer(buf, 'chart1.xml');
      expect(updatedXml).toContain('<c:bar3DChart>');
      expect(updatedXml).toContain('<c:shape val="cylinder"/>');
      expect(updatedXml).toContain('<c:rotX val="25"/>');
      expect(updatedXml).toContain('<c:rotY val="40"/>');
      expect(updatedXml).toContain('3D Cylinder Analysis');
      expect(updatedXml).toContain('<c:v>Phase A</c:v>');
      expect(updatedXml).toContain('<c:v>Optimization</c:v>');
      expect(updatedXml).toContain('<c:v>75</c:v>');

      // Verify embedded workbook
      const sheetXml = readWorkbookSheetFromBuffer(buf, 'Microsoft_Excel_Worksheet1.xlsx');
      expect(sheetXml).toContain('Phase A');
      expect(sheetXml).toContain('Phase C');
      expect(sheetXml).toContain('Optimization');
      expect(sheetXml).toContain('<v>75</v>');

      // Verify AST node
      expect(chartNode.metadata.title).toBe('3D Cylinder Analysis');
      expect(chartNode.metadata.categories).toEqual(['Phase A', 'Phase B', 'Phase C']);
      expect(chartNode.type).toBe('chart3d');
      expect(chartNode.version).toBe(2);

      const docxVal = await validateDocxBuffer(buf);
      expect(docxVal.ok).toBe(true);
      expect(docxVal.errors).toEqual([]);
    });

    it('supports 2D to 3D and 3D to 2D type conversion in-place without duplicate parts', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();

      // Start as 2D line chart
      executeChart(
        {
          action: 'chart',
          chartType: 'line',
          title: 'Trend Line',
          categories: ['W1', 'W2'],
          series: [{ name: 'Clicks', values: [500, 750] }],
          id: 'convertible_chart',
        },
        ast,
        parser
      );

      const node = ast.sections[0].nodes.find((n) => n.id === 'convertible_chart')!;
      expect(node.type).toBe('chart');
      expect(node.metadata.chartType).toBe('line');

      // Convert to 3D column3D
      executeChart(
        {
          action: 'chart',
          operation: 'update',
          target: 'convertible_chart',
          chartType: 'column3D',
          title: '3D Converted Chart',
        },
        ast,
        parser
      );

      expect(node.type).toBe('chart3d');
      expect(node.metadata.chartType).toBe('column3D');
      expect(node.version).toBe(2);

      let buf = parser.rezip();
      expect(listChartEntriesFromBuffer(buf)).toEqual(['word/charts/chart1.xml']);
      let xml = chartXmlFromBuffer(buf, 'chart1.xml');
      expect(xml).toContain('<c:bar3DChart>');
      expect(xml).toContain('<c:barDir val="col"/>');

      // Convert back to 2D pie
      executeChart(
        {
          action: 'chart',
          operation: 'update',
          target: 'convertible_chart',
          chartType: 'pie',
          title: '2D Pie Converted',
        },
        ast,
        parser
      );

      expect(node.type).toBe('chart');
      expect(node.metadata.chartType).toBe('pie');
      expect(node.version).toBe(3);

      buf = parser.rezip();
      expect(listChartEntriesFromBuffer(buf)).toEqual(['word/charts/chart1.xml']);
      xml = chartXmlFromBuffer(buf, 'chart1.xml');
      expect(xml).toContain('<c:pieChart>');
      expect(xml).not.toContain('3DChart');

      const docxVal = await validateDocxBuffer(buf);
      expect(docxVal.ok).toBe(true);
      expect(docxVal.errors).toEqual([]);
    });
  });

  describe('Nested params handling (OP-16)', () => {
    it('accepts parameters nested under params: {} and updates accurately', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();

      executeChart(
        {
          action: 'chart',
          chartType: 'column',
          title: 'Before Nested Update',
          categories: ['A', 'B'],
          series: [{ name: 'V', values: [1, 2] }],
          id: 'nested_test_node',
        },
        ast,
        parser
      );

      // Invoke update with nested params
      executeChart(
        {
          action: 'chart',
          operation: 'update',
          params: {
            target: 'nested_test_node',
            title: 'After Nested Update',
            categories: ['X', 'Y', 'Z'],
            series: [{ name: 'NewMetric', values: [100, 200, 300] }],
          },
        },
        ast,
        parser
      );

      const node = ast.sections[0].nodes.find((n) => n.id === 'nested_test_node')!;
      expect(node.metadata.title).toBe('After Nested Update');
      expect(node.metadata.categories).toEqual(['X', 'Y', 'Z']);
      expect(node.version).toBe(2);

      const buf = parser.rezip();
      expect(listChartEntriesFromBuffer(buf)).toEqual(['word/charts/chart1.xml']);
      const xml = chartXmlFromBuffer(buf, 'chart1.xml');
      expect(xml).toContain('After Nested Update');
      expect(xml).toContain('<c:v>NewMetric</c:v>');
    });
  });

  describe('params-nested insert', () => {
    it('insert honors title/categories/series nested under params (2D)', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();
      const anchor = ast.sections[0].nodes[0].id;

      executeChart(
        {
          action: 'chart',
          operation: 'insert',
          chartType: 'column',
          anchor,
          params: {
            title: 'Nested',
            categories: ['A'],
            series: [{ name: 'S', values: [1] }],
          },
        },
        ast,
        parser
      );

      const node = ast.sections[0].nodes.find((n) => n.type === 'chart')!;
      expect(node).toBeDefined();
      expect(node.metadata.title).toBe('Nested');
      expect(node.metadata.text).toBe('Nested');
      expect(node.metadata.categories).toEqual(['A']);
      expect(node.metadata.series).toMatchObject([{ name: 'S', values: [1] }]);
      expect(node.metadata.chartData).toMatchObject({ title: 'Nested' });

      const xml = chartXmlFromBuffer(parser.rezip(), 'chart1.xml');
      expect(xml).toContain('Nested');
      expect(xml).toContain('<c:v>A</c:v>');
      expect(xml).toContain('<c:v>S</c:v>');
      expect(xml).not.toContain('Time Saved by SDLC Stage (Hours)');
    });

    it('insert honors params-nested data on the 3D path', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();
      const anchor = ast.sections[0].nodes[0].id;

      executeChart(
        {
          action: 'chart',
          operation: 'insert',
          chartType: 'bar3D',
          anchor,
          params: {
            title: 'Nested 3D',
            categories: ['Stage 1'],
            series: [{ name: 'Latency', values: [50] }],
          },
        },
        ast,
        parser
      );

      const node = ast.sections[0].nodes.find((n) => n.type === 'chart3d')!;
      expect(node).toBeDefined();
      expect(node.metadata.title).toBe('Nested 3D');
      expect(node.metadata.categories).toEqual(['Stage 1']);

      const xml = chartXmlFromBuffer(parser.rezip(), 'chart1.xml');
      expect(xml).toContain('<c:bar3DChart>');
      expect(xml).toContain('Nested 3D');
      expect(xml).toContain('<c:v>Latency</c:v>');
    });

    it('flat and params-nested inserts with the same data converge (flat path unchanged)', async () => {
      async function insertDoc(op: Record<string, unknown>) {
        const parser = DocxParser.createNewDocument('blank');
        const ast = await parser.getAST();
        executeChart(
          {
            action: 'chart',
            operation: 'insert',
            chartType: 'column',
            anchor: ast.sections[0].nodes[0].id,
            id: 'converge_node',
            ...op,
          } as never,
          ast,
          parser
        );
        const node = ast.sections[0].nodes.find((n) => n.type === 'chart')!;
        return { node, xml: chartXmlFromBuffer(parser.rezip(), 'chart1.xml') };
      }

      const data = {
        title: 'Converge',
        categories: ['A', 'B'],
        series: [{ name: 'S', values: [1, 2] }],
      };
      const flat = await insertDoc(data);
      const nested = await insertDoc({ params: data });

      expect(nested.node.metadata).toEqual(flat.node.metadata);
      expect(nested.node.metadata.title).toBe('Converge');
      for (const doc of [flat, nested]) {
        expect(doc.xml).toContain('Converge');
        expect(doc.xml).toContain('<c:v>S</c:v>');
        expect(doc.xml).not.toContain('Time Saved by SDLC Stage (Hours)');
      }
    });
  });

  describe('OperationEngine routing', () => {
    it('routes chart update through OperationEngine.execute', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();
      const engine = new OperationEngine(ast, parser);

      await engine.execute({
        action: 'chart',
        chartType: 'column',
        title: 'Initial Via Engine',
        categories: ['Q1', 'Q2'],
        series: [{ name: 'Val', values: [10, 20] }],
        id: 'engine_chart_node',
      });

      await engine.execute({
        action: 'chart',
        operation: 'update',
        target: 'engine_chart_node',
        title: 'Updated Via Engine',
        series: [{ name: 'Val', values: [50, 100] }],
      });

      const node = ast.sections[0].nodes.find((n) => n.id === 'engine_chart_node')!;
      expect(node.metadata.title).toBe('Updated Via Engine');
      expect(node.version).toBe(2);

      const buf = parser.rezip();
      expect(listChartEntriesFromBuffer(buf)).toEqual(['word/charts/chart1.xml']);
      const xml = chartXmlFromBuffer(buf, 'chart1.xml');
      expect(xml).toContain('Updated Via Engine');
      expect(xml).toContain('<c:v>100</c:v>');
    });
  });

  describe('Fail-Closed Validation Tests (E_VALIDATION & Actionable Remedy)', () => {
    it('fails closed when target node is not found', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();

      executeChart(
        {
          action: 'chart',
          chartType: 'column',
          title: 'Existing Chart',
          id: 'existing_chart_node',
        },
        ast,
        parser
      );

      try {
        executeChart(
          {
            action: 'chart',
            operation: 'update',
            target: 'non_existent_node_id',
            title: 'Ghost',
          },
          ast,
          parser
        );
        expect.unreachable('Should have thrown E_VALIDATION');
      } catch (err: unknown) {
        const error = err as Error & { code?: string; remedy?: string };
        expect(error.code).toBe('E_VALIDATION');
        expect(error.message).toContain("target target 'non_existent_node_id' not found");
        expect(error.remedy).toContain('inspect_document');
        expect(error.remedy).toContain('existing_chart_node');
      }
    });

    it('fails closed when target node is not a chart node (targets paragraph)', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();

      // Find the existing paragraph in the blank doc
      const paraNode = ast.sections[0].nodes.find((n) => n.type === 'paragraph')!;
      expect(paraNode).toBeDefined();

      executeChart(
        {
          action: 'chart',
          chartType: 'column',
          title: 'Existing Chart',
          id: 'chart_node_real',
        },
        ast,
        parser
      );

      try {
        executeChart(
          {
            action: 'chart',
            operation: 'update',
            target: paraNode.id,
            title: 'Bad Target Node',
          },
          ast,
          parser
        );
        expect.unreachable('Should have thrown E_VALIDATION');
      } catch (err: unknown) {
        const error = err as Error & { code?: string; remedy?: string };
        expect(error.code).toBe('E_VALIDATION');
        expect(error.message).toContain(`is a 'paragraph' node, not a chart`);
        expect(error.remedy).toContain('chart_node_real');
      }
    });

    it('fails closed when target is empty string or non-string', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();

      try {
        executeChart(
          {
            action: 'chart',
            operation: 'update',
            target: '',
            title: 'Empty Target',
          },
          ast,
          parser
        );
        expect.unreachable('Should have thrown E_VALIDATION');
      } catch (err: unknown) {
        const error = err as Error & { code?: string; remedy?: string };
        expect(error.code).toBe('E_VALIDATION');
        expect(error.message).toContain('chart update target must be a non-empty chart node id string');
        expect(error.remedy).toContain('inspect_document');
      }

      try {
        executeChart(
          {
            action: 'chart',
            operation: 'update',
            target: 12345 as unknown as string,
            title: 'Numeric Target',
          },
          ast,
          parser
        );
        expect.unreachable('Should have thrown E_VALIDATION');
      } catch (err: unknown) {
        const error = err as Error & { code?: string; remedy?: string };
        expect(error.code).toBe('E_VALIDATION');
        expect(error.message).toContain('chart update target must be a non-empty chart node id string');
      }
    });

    it('fails closed when document has zero charts and no target is given', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();

      try {
        executeChart(
          {
            action: 'chart',
            operation: 'update',
            title: 'No Charts Exist',
          },
          ast,
          parser
        );
        expect.unreachable('Should have thrown E_VALIDATION');
      } catch (err: unknown) {
        const error = err as Error & { code?: string; remedy?: string };
        expect(error.code).toBe('E_VALIDATION');
        expect(error.message).toContain('no charts exist in this document');
        expect(error.remedy).toContain("action:'chart', operation:'insert'");
      }
    });

    it('fails closed when multiple charts exist and update target is ambiguous', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();

      executeChart(
        {
          action: 'chart',
          chartType: 'column',
          title: 'Chart One',
          id: 'chart_1_node',
        },
        ast,
        parser
      );

      executeChart(
        {
          action: 'chart',
          chartType: 'line',
          title: 'Chart Two',
          id: 'chart_2_node',
        },
        ast,
        parser
      );

      try {
        executeChart(
          {
            action: 'chart',
            operation: 'update',
            title: 'Ambiguous Target',
          },
          ast,
          parser
        );
        expect.unreachable('Should have thrown E_VALIDATION');
      } catch (err: unknown) {
        const error = err as Error & { code?: string; remedy?: string };
        expect(error.code).toBe('E_VALIDATION');
        expect(error.message).toContain('chart update is ambiguous: 2 charts exist and no target was given');
        expect(error.remedy).toContain('chart_1_node');
      }
    });

    it('fails closed on non-numeric series values (NaN, Infinity, string)', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();

      executeChart(
        {
          action: 'chart',
          chartType: 'column',
          title: 'Valid Base Chart',
          id: 'c_val_node',
        },
        ast,
        parser
      );

      for (const badValue of ['text', NaN, Infinity, -Infinity, null, undefined]) {
        try {
          executeChart(
            {
              action: 'chart',
              operation: 'update',
              target: 'c_val_node',
              series: [{ name: 'S1', values: [10, badValue as unknown as number] }],
            },
            ast,
            parser
          );
          expect.unreachable(`Should have thrown for value: ${String(badValue)}`);
        } catch (err: unknown) {
          const error = err as Error & { code?: string; remedy?: string };
          expect(error.code).toBe('E_VALIDATION');
          expect(error.message).toContain('has a non-numeric value');
          expect(error.remedy).toContain('finite numbers only');
        }
      }
    });

    it('fails closed on empty series or series with empty values array', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();

      executeChart(
        {
          action: 'chart',
          chartType: 'column',
          title: 'Valid Base Chart',
          id: 'c_empty_node',
        },
        ast,
        parser
      );

      try {
        executeChart(
          {
            action: 'chart',
            operation: 'update',
            target: 'c_empty_node',
            series: [],
          },
          ast,
          parser
        );
        expect.unreachable('Should have thrown on empty series');
      } catch (err: unknown) {
        const error = err as Error & { code?: string; remedy?: string };
        expect(error.code).toBe('E_VALIDATION');
        expect(error.message).toContain('series must be a non-empty array');
      }

      try {
        executeChart(
          {
            action: 'chart',
            operation: 'update',
            target: 'c_empty_node',
            series: [{ name: 'EmptyValues', values: [] }],
          },
          ast,
          parser
        );
        expect.unreachable('Should have thrown on series with empty values');
      } catch (err: unknown) {
        const error = err as Error & { code?: string; remedy?: string };
        expect(error.code).toBe('E_VALIDATION');
        expect(error.message).toContain('needs non-empty numeric values');
      }
    });

    it('fails closed on invalid series name or invalid categories', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();

      executeChart(
        {
          action: 'chart',
          chartType: 'column',
          title: 'Valid Base Chart',
          id: 'c_inv_node',
        },
        ast,
        parser
      );

      // Non-string series name
      try {
        executeChart(
          {
            action: 'chart',
            operation: 'update',
            target: 'c_inv_node',
            series: [{ name: 12345 as unknown as string, values: [10, 20] }],
          },
          ast,
          parser
        );
        expect.unreachable('Should have thrown on non-string series name');
      } catch (err: unknown) {
        const error = err as Error & { code?: string; remedy?: string };
        expect(error.code).toBe('E_VALIDATION');
        expect(error.message).toContain('series entries need a string name');
      }

      // Empty categories array
      try {
        executeChart(
          {
            action: 'chart',
            operation: 'update',
            target: 'c_inv_node',
            categories: [],
          },
          ast,
          parser
        );
        expect.unreachable('Should have thrown on empty categories');
      } catch (err: unknown) {
        const error = err as Error & { code?: string; remedy?: string };
        expect(error.code).toBe('E_VALIDATION');
        expect(error.message).toContain('categories must be a non-empty array');
      }

      // Non-string category entry
      try {
        executeChart(
          {
            action: 'chart',
            operation: 'update',
            target: 'c_inv_node',
            categories: ['Q1', 42 as unknown as string],
          },
          ast,
          parser
        );
        expect.unreachable('Should have thrown on non-string category entry');
      } catch (err: unknown) {
        const error = err as Error & { code?: string; remedy?: string };
        expect(error.code).toBe('E_VALIDATION');
        expect(error.message).toContain('categories must be strings');
      }
    });

    it('fails closed on unsupported chartType or invalid 3D bar shape', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();

      executeChart(
        {
          action: 'chart',
          chartType: 'column',
          title: 'Valid Base Chart',
          id: 'c_shape_node',
        },
        ast,
        parser
      );

      try {
        executeChart(
          {
            action: 'chart',
            operation: 'update',
            target: 'c_shape_node',
            chartType: 'unsupportedCustomChartType' as unknown as 'column',
          },
          ast,
          parser
        );
        expect.unreachable('Should have thrown on unsupported chartType');
      } catch (err: unknown) {
        const error = err as Error & { code?: string; remedy?: string };
        expect(error.code).toBe('E_VALIDATION');
        expect(error.message).toContain('unsupported chart type');
        expect(error.remedy).toContain('Use one of the supported chart types');
      }

      try {
        executeChart(
          {
            action: 'chart',
            operation: 'update',
            target: 'c_shape_node',
            shape: 'torus' as unknown as 'box',
          },
          ast,
          parser
        );
        expect.unreachable('Should have thrown on unsupported shape');
      } catch (err: unknown) {
        const error = err as Error & { code?: string; remedy?: string };
        expect(error.code).toBe('E_VALIDATION');
        expect(error.message).toContain('unsupported 3D bar shape "torus"');
        expect(error.remedy).toContain('box | cone | cylinder | pyramid');
      }
    });

    it('fails closed when parser is missing', () => {
      const dummyAst: DocumentAST = {
        metadata: {},
        sections: [{ id: 's1', heading: null, nodes: [], pageSetup: null, version: 1, checksum: 'c' }],
        styles: {},
        numbering: { abstractNum: {}, concreteNum: {} },
        relationships: {},
        headers: {},
        footers: {},
        mediaFiles: {},
        settings: { trackChanges: false, defaultAuthor: 'ASTra' },
        nextChangeId: 1,
      };

      try {
        executeChartUpdate(
          { action: 'chart', operation: 'update', target: 'c1' },
          dummyAst,
          undefined
        );
        expect.unreachable('Should have thrown without parser');
      } catch (err: unknown) {
        const error = err as Error & { code?: string; remedy?: string };
        expect(error.code).toBe('E_VALIDATION');
        expect(error.message).toContain('chart update requires a loaded document package');
      }
    });

    it('enforces optimistic concurrency version check when expectedVersion is supplied', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();

      executeChart(
        {
          action: 'chart',
          chartType: 'column',
          title: 'Versioned Chart',
          id: 'versioned_node',
        },
        ast,
        parser
      );

      const node = ast.sections[0].nodes.find((n) => n.id === 'versioned_node')!;
      expect(node.version).toBe(1);

      // Version mismatch: expect version 999 instead of 1
      expect(() => {
        executeChart(
          {
            action: 'chart',
            operation: 'update',
            target: 'versioned_node',
            title: 'Mismatched Update',
            expectedVersion: 999,
          },
          ast,
          parser
        );
      }).toThrow(/Version mismatch on node versioned_node. Expected 999, got 1/);

      // Version match: expect version 1 succeeds
      executeChart(
        {
          action: 'chart',
          operation: 'update',
          target: 'versioned_node',
          title: 'Matched Version Update',
          expectedVersion: 1,
        },
        ast,
        parser
      );
      expect(node.version).toBe(2);
      expect(node.metadata.title).toBe('Matched Version Update');
    });

    it('version mismatch carries E_VALIDATION code and remedy, with no partial apply', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();

      executeChart(
        {
          action: 'chart',
          chartType: 'column',
          title: 'Version Contract Chart',
          categories: ['Jan', 'Feb'],
          series: [{ name: 'Sales', values: [10, 20], color: '0066CC' }],
          id: 'version_contract_node',
        },
        ast,
        parser
      );

      const node = ast.sections[0].nodes.find((n) => n.id === 'version_contract_node')!;
      expect(node.version).toBe(1);
      const checksumBefore = node.checksum;
      const xmlBefore = chartXmlFromBuffer(parser.rezip(), 'chart1.xml');

      try {
        executeChart(
          {
            action: 'chart',
            operation: 'update',
            target: 'version_contract_node',
            title: 'Stale Write Attempt',
            expectedVersion: 999,
          },
          ast,
          parser
        );
        expect.unreachable('Should have thrown version mismatch');
      } catch (err: unknown) {
        const error = err as Error & { code?: string; retryable?: boolean; remedy?: string };
        expect(error.code).toBe('E_VALIDATION');
        expect(error.retryable).toBe(false);
        expect(error.message).toContain('E_VALIDATION');
        expect(error.message).toContain('Version mismatch on node version_contract_node. Expected 999, got 1');
        expect(error.remedy).toBeDefined();
        expect(error.remedy).toContain('inspect_document');
        expect(error.remedy).toContain('version_contract_node');
      }

      expect(node.version).toBe(1);
      expect(node.checksum).toBe(checksumBefore);
      expect(node.metadata.title).toBe('Version Contract Chart');
      const xmlAfter = chartXmlFromBuffer(parser.rezip(), 'chart1.xml');
      expect(xmlAfter).toBe(xmlBefore);
      expect(xmlAfter).not.toContain('Stale Write Attempt');
    });

    it('rejects empty update with no mutable fields without version bump or zip write', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();

      executeChart(
        {
          action: 'chart',
          chartType: 'column',
          title: 'Empty Guard Base',
          categories: ['Jan', 'Feb'],
          series: [{ name: 'Sales', values: [10, 20], color: '0066CC' }],
          id: 'empty_guard_node',
        },
        ast,
        parser
      );

      const node = ast.sections[0].nodes.find((n) => n.id === 'empty_guard_node')!;
      expect(node.version).toBe(1);
      const checksumBefore = node.checksum;
      const xmlBefore = chartXmlFromBuffer(parser.rezip(), 'chart1.xml');

      const emptyAttempts: Array<Record<string, unknown>> = [
        { action: 'chart', operation: 'update', target: 'empty_guard_node' },
        { action: 'chart', operation: 'update', target: 'empty_guard_node', expectedVersion: 1 },
        {
          action: 'chart',
          operation: 'update',
          params: { target: 'empty_guard_node' },
        },
      ];
      for (const attempt of emptyAttempts) {
        try {
          executeChart(attempt as never, ast, parser);
          expect.unreachable('Empty update should have thrown E_VALIDATION');
        } catch (err: unknown) {
          const error = err as Error & { code?: string; remedy?: string };
          expect(error.code).toBe('E_VALIDATION');
          expect(error.message).toContain('no mutable fields');
          expect(error.remedy).toContain('inspect_document');
        }
        expect(node.version).toBe(1);
        expect(node.checksum).toBe(checksumBefore);
      }

      const xmlAfter = chartXmlFromBuffer(parser.rezip(), 'chart1.xml');
      expect(xmlAfter).toBe(xmlBefore);
      expect(listChartEntriesFromBuffer(parser.rezip())).toEqual(['word/charts/chart1.xml']);
    });

    it('rejects implicit update when zip part exists but AST has no chart node', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();

      executeChart(
        {
          action: 'chart',
          chartType: 'column',
          title: 'Skew Base Title',
          categories: ['Jan', 'Feb'],
          series: [{ name: 'Sales', values: [10, 20], color: '0066CC' }],
          id: 'skew_node',
        },
        ast,
        parser
      );

      const xmlBefore = chartXmlFromBuffer(parser.rezip(), 'chart1.xml');
      expect(xmlBefore).toContain('Skew Base Title');

      ast.sections[0].nodes = ast.sections[0].nodes.filter((n) => n.id !== 'skew_node');
      expect(ast.sections[0].nodes.find((n) => n.id === 'skew_node')).toBeUndefined();
      expect(parser.getZipAdapter().hasFile('word/charts/chart1.xml')).toBe(true);

      try {
        executeChart(
          {
            action: 'chart',
            operation: 'update',
            title: 'Skew Sneak Title',
            series: [{ name: 'Sales', values: [99, 99], color: '0066CC' }],
          },
          ast,
          parser
        );
        expect.unreachable('Skewed update should have thrown E_VALIDATION');
      } catch (err: unknown) {
        const error = err as Error & { code?: string; remedy?: string };
        expect(error.code).toBe('E_VALIDATION');
        expect(error.message).toContain('no chart AST node was found');
        expect(error.remedy).toContain('inspect_document');
      }

      const xmlAfter = chartXmlFromBuffer(parser.rezip(), 'chart1.xml');
      expect(xmlAfter).toBe(xmlBefore);
      expect(xmlAfter).toContain('Skew Base Title');
      expect(xmlAfter).not.toContain('Skew Sneak Title');
      expect(ast.sections[0].nodes.find((n) => n.id === 'skew_node')).toBeUndefined();
      expect(listChartEntriesFromBuffer(parser.rezip())).toEqual(['word/charts/chart1.xml']);
    });
  });
});

import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import AdmZip from '../adm-zip-safe';
import { clampZipEntryTimes } from '../zip-hardening';
import { DocumentAST } from '../ast';
import { DocxParser, DocxZipAdapter } from '../parser';
import { readChartStyle, type ChartStyle, type ChartFloatingStyle } from '../operations/chart-style';
import { chartTitleElement } from '../parser/factory';
import { buildChartDrawingContainer } from './chart-drawing';
import { RelationshipsManipulator } from './relationships';
import { createChartElement, createRelationshipsElement, createContentTypesElement, allocateNextDrawingDocPrId } from '../ooxml/ns';
import { assertNoXxe } from '../ooxml/xxe-guard';
import { expandShorthandHex, isShorthandHex } from '../style-guide/parser';

/**
 * Requestable 3D chart types: exactly the four the writer produces (identical
 * to SUPPORTED_CHART_TYPES). Read-only surfacing (`summarize3DChartXml`) may
 * still report 'area3D'/'surface3D' for foreign parts it detects, but those
 * are refused by the update/insert gates and are never produced here.
 */
export type Chart3DType = 'bar3D' | 'pie3D' | 'line3D' | 'column3D';
export const SUPPORTED_CHART_TYPES = ['bar3D', 'line3D', 'pie3D', 'column3D'] as const;
export type Bar3DShape = 'box' | 'cone' | 'cylinder' | 'pyramid';

export interface View3DProperties {
  rotX?: number; // Range [-90, 90]
  rotY?: number; // Range [0, 360]
  perspective?: number; // Range [0, 240]
  rAngAx?: boolean; // Right Angle Axes
  hPercent?: number; // Height percent [5, 500]
  depthPercent?: number; // Depth percent [20, 2000]
}

export interface Chart3DSeries {
  name: string;
  values: number[];
  color?: string; // HEX color without '#' e.g. '0066CC'
}

export interface CreateChart3DOptions {
  chartType?: Chart3DType;
  title?: string;
  view3D?: View3DProperties;
  shape?: Bar3DShape;
  categories?: string[];
  series?: Chart3DSeries[];
  /** Internal: transform generated chart XML before it is written. */
  transformChartXml?: (xml: string) => string;
  /** Floating placement; omitted keeps the legacy inline drawing. */
  floating?: ChartFloatingStyle;
}

export interface InsertChart3DOptions extends CreateChart3DOptions {
  anchorNodeId?: string;
  id?: string;
  widthEmu?: number; // Default 5486400 (6 inches)
  heightEmu?: number; // Default 3200400 (3.5 inches)
}

/**
 * Architectural Note on Chart Synchronization Boundary:
 * ASTra updates the OOXML DrawingML chart series cache (<c:strCache>, <c:numCache>),
 * which Word renders natively on open, while embedded Excel binary workbooks
 * (word/embeddings/*.xlsx) remain decoupled. When charts are mutated or updated,
 * Word relies primarily on the cached DrawingML series data (<c:strCache>, <c:numCache>)
 * for rendering and visual layout fidelity, while the embedded workbook acts as an
 * optional data snapshot for external Excel-based editing.
 *
 * Generates a valid minimal ISO/IEC 29500-1 (ECMA-376) SpreadsheetML (.xlsx) workbook
 * containing the chart's series and categorical data.
 * When embedded in word/embeddings/*.xlsx, Microsoft Word enables "Edit Data in Excel"
 * and Google Docs reliably renders chart labels and legends without dropping metadata.
 */
export function generateEmbeddedWorkbook(
  seriesList: Chart3DSeries[],
  categories: string[],
  scatterSeriesXValues?: ReadonlyArray<ReadonlyArray<number>>
): Buffer {
  const zip = new AdmZip();

  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;
  zip.addFile('[Content_Types].xml', Buffer.from(contentTypesXml, 'utf-8'));

  const rootRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
  zip.addFile('_rels/.rels', Buffer.from(rootRelsXml, 'utf-8'));

  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;
  zip.addFile('xl/workbook.xml', Buffer.from(workbookXml, 'utf-8'));

  const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
  zip.addFile('xl/_rels/workbook.xml.rels', Buffer.from(workbookRelsXml, 'utf-8'));

  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>`;
  zip.addFile('xl/styles.xml', Buffer.from(stylesXml, 'utf-8'));

  function getColLetter(colIdx: number): string {
    let letter = '';
    let temp = colIdx;
    while (temp >= 0) {
      letter = String.fromCharCode((temp % 26) + 65) + letter;
      temp = Math.floor(temp / 26) - 1;
    }
    return letter;
  }

  function escapeXml(str: string): string {
    return String(str || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  let rowsXml = '';
  if (scatterSeriesXValues !== undefined) {
    // Scatter keeps x per series, so the canonical workbook pairs each
    // series' x and y columns (A=x1, B=y1, C=x2, D=y2, ...). Chart formulas
    // written by the merge path point at these same columns.
    let headerCells = '';
    seriesList.forEach((ser, sIdx) => {
      const xCol = getColLetter(sIdx * 2);
      const yCol = getColLetter(sIdx * 2 + 1);
      headerCells += `<c r="${xCol}1" t="inlineStr"><is><t>${escapeXml(ser.name)} X</t></is></c>`;
      headerCells += `<c r="${yCol}1" t="inlineStr"><is><t>${escapeXml(ser.name)}</t></is></c>`;
    });
    rowsXml += `<row r="1">${headerCells}</row>`;

    const rowCount = seriesList.reduce((max, ser, sIdx) => {
      const xLen = scatterSeriesXValues[sIdx]?.length ?? 0;
      return Math.max(max, xLen, ser.values?.length ?? 0);
    }, 0);
    for (let i = 0; i < rowCount; i++) {
      const rowNum = i + 2;
      let cells = '';
      seriesList.forEach((ser, sIdx) => {
        const xCol = getColLetter(sIdx * 2);
        const yCol = getColLetter(sIdx * 2 + 1);
        const xValue = scatterSeriesXValues[sIdx]?.[i];
        if (xValue !== undefined && Number.isFinite(xValue)) {
          cells += `<c r="${xCol}${rowNum}"><v>${xValue}</v></c>`;
        }
        const val = ser.values?.[i];
        if (val !== undefined && Number.isFinite(val)) {
          cells += `<c r="${yCol}${rowNum}"><v>${val}</v></c>`;
        }
      });
      rowsXml += `<row r="${rowNum}">${cells}</row>`;
    }
  } else {
    // Row 1: Headers (Col A: Categories header, Col B..: Series names)
    let headerCells = `<c r="A1" t="inlineStr"><is><t>Categories</t></is></c>`;
    seriesList.forEach((ser, sIdx) => {
      const col = getColLetter(sIdx + 1);
      headerCells += `<c r="${col}1" t="inlineStr"><is><t>${escapeXml(ser.name)}</t></is></c>`;
    });
    rowsXml += `<row r="1">${headerCells}</row>`;

    // Rows 2..N: Category and Values
    categories.forEach((cat, cIdx) => {
      const rowNum = cIdx + 2;
      let cells = `<c r="A${rowNum}" t="inlineStr"><is><t>${escapeXml(cat)}</t></is></c>`;
      seriesList.forEach((ser, sIdx) => {
        const col = getColLetter(sIdx + 1);
        const val = ser.values && ser.values[cIdx] !== undefined ? ser.values[cIdx] : 0;
        cells += `<c r="${col}${rowNum}"><v>${val}</v></c>`;
      });
      rowsXml += `<row r="${rowNum}">${cells}</row>`;
    });
  }

  const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    ${rowsXml}
  </sheetData>
</worksheet>`;
  zip.addFile('xl/worksheets/sheet1.xml', Buffer.from(sheetXml, 'utf-8'));

  clampZipEntryTimes(zip);
  return zip.toBuffer();
}

/** Default series palette: per-point pie fills AND the per-series default for
 * category charts (column/bar/line/area). One scheme for both, so charts
 * without explicit series colors still render distinct series. Entry 0 is
 * the historic single-series blue; explicit series colors always win. */
export const DEFAULT_SERIES_PALETTE = ['0066CC', '10B981', 'F59E0B', 'EF4444', '8B5CF6', '06B6D4', 'EC4899', '84CC16'];

/** Explicit series color, else the palette entry for this series index.
 * Sole funnel from every series color to every `srgbClr val` writer (2D +
 * 3D, insert + update): shorthand `#RGB` is expanded here so emitted XML
 * always carries a valid 6-digit `ST_HexColor`. Already-6-hex input
 * returns verbatim (strip `#` only) — legacy bytes unchanged. */
export function defaultSeriesColor(color: string | undefined, seriesIdx: number): string {
  if (color) {
    const raw = color.replace(/^#/, '');
    if (isShorthandHex(raw)) return expandShorthandHex(raw);
    return raw;
  }
  return DEFAULT_SERIES_PALETTE[seriesIdx % DEFAULT_SERIES_PALETTE.length];
}

function chartColLetter(colIdx: number): string {
  let letter = '';
  let temp = colIdx;
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

/**
 * B-5: numeric values as a referenced cache (c:numCache) so renderers that
 * only read the cache (Word/LO) show data. The formula points at the embedded
 * workbook sheet written by insert3DChart (Sheet1, column B+ per series).
 */
function buildNumCacheXml(values: number[], seriesIdx: number): string {
  const col = chartColLetter(seriesIdx + 1);
  const lastRow = 1 + Math.max(values.length, 1);
  const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('');
  return `<c:numRef><c:f>Sheet1!$${col}$2:$${col}$${lastRow}</c:f>` +
    `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${pts}</c:numCache></c:numRef>`;
}

function buildSolidFillSpPr(color: string | undefined): string {
  const fill = (color || '0066CC').replace(/^#/, '');
  return `<c:spPr><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill></c:spPr>`;
}

/** B-5: visible line stroke for line3D series (default 1pt = 25400 EMU). */
function buildLineSpPr(color: string | undefined, widthEmu = 25400): string {
  const stroke = (color || '0066CC').replace(/^#/, '');
  return `<c:spPr><a:ln w="${widthEmu}"><a:solidFill><a:srgbClr val="${stroke}"/></a:solidFill></a:ln></c:spPr>`;
}

export class Chart3DManipulator {
  private ast: DocumentAST;
  private parser?: DocxParser;
  private domParser: DOMParser;
  private serializer: XMLSerializer;

  constructor(ast: DocumentAST, parser?: DocxParser) {
    this.ast = ast;
    this.parser = parser;
    this.domParser = new DOMParser();
    this.serializer = new XMLSerializer();
  }

  /**
   * Generates ISO/IEC 29500-1 Clause 21.2 compliant <c:chartSpace> XML string for 3D charts.
   */
  public generate3DChartXml(options: CreateChart3DOptions, externalDataRId?: string): string {
    // A part with no title must stay titleless: only an absent option gets
    // the historical default, an explicit empty title omits <c:title>.
    const rawTitle = options.title === undefined ? '3D Chart Analysis' : options.title;
    const title = rawTitle
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
    const titleXml = title.length > 0
      ? `
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr b="1" sz="1400"/>
              <a:t>${title}</a:t>
            </a:r>
          </a:p>
        </c:rich>
      </c:tx>
    </c:title>`
      : '';

    const rawChartType = options.chartType || 'bar3D';
    const supportedChartTypes = SUPPORTED_CHART_TYPES as readonly string[];
    if (!supportedChartTypes.includes(rawChartType)) {
      throw Object.assign(
        new Error(`Unsupported chart type "${rawChartType}". Supported types: ${SUPPORTED_CHART_TYPES.join(', ')}`),
        {
          code: 'E_VALIDATION',
          retryable: false,
          remedy: `Use one of the supported chart types: ${SUPPORTED_CHART_TYPES.join(', ')}`,
        }
      );
    }
    const chartType = rawChartType;

    const view3D = {
      rotX: Math.max(-90, Math.min(90, options.view3D?.rotX ?? 30)),
      rotY: Math.max(0, Math.min(360, options.view3D?.rotY ?? 20)),
      perspective: Math.max(0, Math.min(240, options.view3D?.perspective ?? 30)),
      rAngAx: options.view3D?.rAngAx ?? true,
      hPercent: options.view3D?.hPercent,
      depthPercent: options.view3D?.depthPercent
    };

    const categories = options.categories && options.categories.length > 0
      ? options.categories
      : ['Q1', 'Q2', 'Q3', 'Q4'];

    const seriesList: Chart3DSeries[] = options.series && options.series.length > 0
      ? options.series
      : [{ name: 'Performance', values: [12.5, 14.8, 18.2, 22.0], color: '0066CC' }];

    const shapeTag = options.shape ? `<c:shape val="${options.shape}"/>` : '';

    const escapeXml = (str: string): string => {
      if (!str) return '';
      return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
    };

    let view3DXml = `<c:view3D>
      <c:rotX val="${view3D.rotX}"/>`;
    if (view3D.hPercent !== undefined) view3DXml += `\n      <c:hPercent val="${view3D.hPercent}"/>`;
    view3DXml += `\n      <c:rotY val="${view3D.rotY}"/>`;
    if (view3D.depthPercent !== undefined) view3DXml += `\n      <c:depthPercent val="${view3D.depthPercent}"/>`;
    view3DXml += `\n      <c:rAngAx val="${view3D.rAngAx ? '1' : '0'}"/>`;
    view3DXml += `\n      <c:perspective val="${view3D.perspective}"/>`;
    view3DXml += `\n    </c:view3D>`;

    let chartTypeXml = '';

    if (chartType === 'pie3D') {
      let seriesXml = '';
      seriesList.forEach((ser, sIdx) => {
        let catPts = '';
        categories.forEach((cat, cIdx) => {
          catPts += `<c:pt idx="${cIdx}"><c:v>${escapeXml(cat)}</c:v></c:pt>`;
        });
        // B-5: per-point fills so slices actually differ (varyColors=1 alone
        // renders every slice the series color in several renderers).
        const baseFill = defaultSeriesColor(ser.color, sIdx);
        let dPts = '';
        ser.values.forEach((_v, vIdx) => {
          const fill = vIdx === 0 ? baseFill : DEFAULT_SERIES_PALETTE[vIdx % DEFAULT_SERIES_PALETTE.length];
          dPts += `<c:dPt><c:idx val="${vIdx}"/>${buildSolidFillSpPr(fill)}</c:dPt>`;
        });

        seriesXml += `<c:ser>
          <c:idx val="${sIdx}"/>
          <c:order val="${sIdx}"/>
          <c:tx><c:v>${escapeXml(ser.name)}</c:v></c:tx>
          ${buildSolidFillSpPr(baseFill)}
          ${dPts}
          <c:cat><c:strLit><c:ptCount val="${categories.length}"/>${catPts}</c:strLit></c:cat>
          <c:val>${buildNumCacheXml(ser.values, sIdx)}</c:val>
        </c:ser>`;
      });

      chartTypeXml = `<c:pie3DChart>
        <c:varyColors val="1"/>
        ${seriesXml}
      </c:pie3DChart>`;
    } else if (chartType === 'line3D') {
      let seriesXml = '';
      seriesList.forEach((ser, sIdx) => {
        let catPts = '';
        categories.forEach((cat, cIdx) => {
          catPts += `<c:pt idx="${cIdx}"><c:v>${escapeXml(cat)}</c:v></c:pt>`;
        });

        seriesXml += `<c:ser>
          <c:idx val="${sIdx}"/>
          <c:order val="${sIdx}"/>
          <c:tx><c:v>${escapeXml(ser.name)}</c:v></c:tx>
          ${buildLineSpPr(defaultSeriesColor(ser.color, sIdx))}
          <c:cat><c:strLit><c:ptCount val="${categories.length}"/>${catPts}</c:strLit></c:cat>
          <c:val>${buildNumCacheXml(ser.values, sIdx)}</c:val>
        </c:ser>`;
      });

      // ECMA-376 line3DChart requires 3 axes: catAx, valAx, and serAx
      chartTypeXml = `<c:line3DChart>
        <c:grouping val="standard"/>
        ${seriesXml}
        <c:axId val="1001"/>
        <c:axId val="1002"/>
        <c:axId val="1003"/>
      </c:line3DChart>`;
    } else {
      // Default bar3DChart
      let seriesXml = '';
      seriesList.forEach((ser, sIdx) => {
        let catPts = '';
        categories.forEach((cat, cIdx) => {
          catPts += `<c:pt idx="${cIdx}"><c:v>${escapeXml(cat)}</c:v></c:pt>`;
        });
        const fillColor = defaultSeriesColor(ser.color, sIdx);

        seriesXml += `<c:ser>
          <c:idx val="${sIdx}"/>
          <c:order val="${sIdx}"/>
          <c:tx><c:v>${escapeXml(ser.name)}</c:v></c:tx>
          <c:spPr>
            <a:solidFill>
              <a:srgbClr val="${fillColor}"/>
            </a:solidFill>
          </c:spPr>
          <c:cat><c:strLit><c:ptCount val="${categories.length}"/>${catPts}</c:strLit></c:cat>
          <c:val>${buildNumCacheXml(ser.values, sIdx)}</c:val>
        </c:ser>`;
      });

      const barDir = chartType === 'bar3D' ? 'bar' : 'col';
      chartTypeXml = `<c:bar3DChart>
        <c:barDir val="${barDir}"/>
        <c:grouping val="standard"/>
        ${seriesXml}
        <c:gapWidth val="150"/>
        <c:gapDepth val="150"/>
        ${shapeTag}
        <c:axId val="1001"/>
        <c:axId val="1002"/>
      </c:bar3DChart>`;
    }

    const categoryAxisXml = chartType === 'pie3D' ? '' : `<c:catAx>
        <c:axId val="1001"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="1002"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
        <c:noMultiLvlLbl val="0"/>
      </c:catAx>`;

    const valueAxisXml = chartType === 'pie3D' ? '' : `<c:valAx>
        <c:axId val="1002"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:majorGridlines/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="1001"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>`;

    const seriesAxisXml = chartType === 'line3D' ? `<c:serAx>
        <c:axId val="1003"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="1002"/>
        <c:crosses val="autoZero"/>
      </c:serAx>` : '';

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:date1904 val="0"/>
  <c:lang val="en-US"/>
  <c:chart>${titleXml}
    ${view3DXml}
    <c:plotArea>
      <c:layout/>
      ${chartTypeXml}
      ${categoryAxisXml}
      ${valueAxisXml}
      ${seriesAxisXml}
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:overlay val="0"/>
    </c:legend>
  </c:chart>
  ${externalDataRId ? `<c:externalData r:id="${externalDataRId}"><c:autoUpdate val="0"/></c:externalData>` : ''}
</c:chartSpace>`;
  }

  /**
   * Generates DrawingML inline wrapper for embedding chart in document.xml
   */
  public generateDrawingEmbedXml(rId: string, docPrId: number = 1, name: string = 'Chart 1', widthEmu: number = 5486400, heightEmu: number = 3200400, floating?: ChartFloatingStyle): string {
    const safeName = (name || 'Chart')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');

    return `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
     xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:pPr><w:spacing w:before="140" w:after="140" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>
  <w:r>
    <w:drawing>
      ${buildChartDrawingContainer({ rId, docPrId, name: safeName, widthEmu, heightEmu, floating })}
    </w:drawing>
  </w:r>
</w:p>`;
  }

  /**
   * Inserts a complete 3D chart into the document package and AST.
   */
  public insert3DChart(options: InsertChart3DOptions): { rId: string; chartPath: string; drawingNodeId?: string; workbookPath?: string } {
    const rawChartType = options.chartType || 'bar3D';
    const supportedChartTypes = SUPPORTED_CHART_TYPES as readonly string[];
    if (!supportedChartTypes.includes(rawChartType)) {
      throw Object.assign(
        new Error(`Unsupported chart type "${rawChartType}". Supported types: ${supportedChartTypes.join(', ')}`),
        {
          code: 'E_VALIDATION',
          retryable: false,
          remedy: `Use one of the supported chart types: ${supportedChartTypes.join(', ')}`,
        }
      );
    }

    const relsManipulator = new RelationshipsManipulator(this.ast);

    // Calculate next chart part number
    let maxNum = 1;
    for (const rel of Object.values(this.ast.relationships)) {
      if (rel.type.includes('chart')) {
        const match = rel.target.match(/chart(\d+)\.xml/);
        if (match) {
          const num = parseInt(match[1], 10);
          if (num >= maxNum) maxNum = num + 1;
        }
      }
    }

    const chartPartName = `charts/chart${maxNum}.xml`;
    const relType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
    const rId = relsManipulator.addRelationship(relType, chartPartName);

    const categories = options.categories && options.categories.length > 0
      ? options.categories
      : ['Q1', 'Q2', 'Q3', 'Q4'];

    const seriesList: Chart3DSeries[] = options.series && options.series.length > 0
      ? options.series
      : [{ name: 'Performance', values: [12.5, 14.8, 18.2, 22.0], color: '0066CC' }];

    const widthEmu = options.widthEmu || 5486400;
    const heightEmu = options.heightEmu || 3200400;

    // A-2: globally unique docPr id via shared DOM-scanning allocator
    // (images + charts occupy the same id space; scans wp:docPr + pic:cNvPr).
    let parserDom: Document | null = null;
    try {
      parserDom = this.parser ? this.parser.getDocumentDom() : null;
    } catch { parserDom = null; }
    const docPrId = allocateNextDrawingDocPrId(parserDom);

    const drawingXml = this.generateDrawingEmbedXml(rId, docPrId, options.title || `Chart ${maxNum}`, widthEmu, heightEmu, options.floating);

    if (this.parser) {
      const workbookPartName = `embeddings/Microsoft_Excel_Worksheet${maxNum}.xlsx`;
      const workbookBuf = generateEmbeddedWorkbook(seriesList, categories);
      const zipAdapter = this.parser.getZipAdapter();

      // 1. Embed real Excel workbook binary
      zipAdapter.writeBuffer(`word/${workbookPartName}`, workbookBuf);

      // 2. Add chart relationship pointing to embedded Excel workbook
      const chartRelsPath = `word/charts/_rels/chart${maxNum}.xml.rels`;
      const chartRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet${maxNum}.xlsx"/>
</Relationships>`;
      zipAdapter.writeText(chartRelsPath, chartRelsXml);

      // 3. Generate chart.xml with c:externalData referencing rId1
      const generatedXml = this.generate3DChartXml(options, 'rId1');
      const chartXmlStr = options.transformChartXml ? options.transformChartXml(generatedXml) : generatedXml;
      zipAdapter.writeText(`word/${chartPartName}`, chartXmlStr);

      // 4. Ensure relationship in word/_rels/document.xml.rels
      this.ensureDocumentRelationship(zipAdapter, rId, relType, chartPartName);

      // 5. Ensure content type overrides in [Content_Types].xml
      const chartContentType = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
      this.ensureContentTypeOverride(zipAdapter, `/word/${chartPartName}`, chartContentType);

      const xlsxContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      this.ensureContentTypeOverride(zipAdapter, `/word/${workbookPartName}`, xlsxContentType);
      this.ensureContentTypeDefault(zipAdapter, 'xlsx', xlsxContentType);

      let drawingNodeId: string | undefined;
      if (options.anchorNodeId) {
        drawingNodeId = this.parser.insertRootNode(drawingXml, options.anchorNodeId, 'after', options.id);
      } else {
        drawingNodeId = this.parser.insertRootNode(drawingXml, undefined, 'append', options.id);
      }
      return { rId, chartPath: `word/${chartPartName}`, drawingNodeId, workbookPath: `word/${workbookPartName}` };
    }

    const chartXmlStr = this.generate3DChartXml(options);
    void chartXmlStr;
    return { rId, chartPath: `word/${chartPartName}` };
  }

  /**
   * Updates 3D view perspective controls in an existing chart XML string.
   */
  public updateView3D(chartXmlStr: string, view3DProps: View3DProperties): string {
    const doc = this.domParser.parseFromString(chartXmlStr, 'text/xml');
    const chartNode = doc.getElementsByTagName('c:chart')[0];
    if (!chartNode) throw new Error("Invalid ChartML: missing <c:chart>");

    let view3DNode = chartNode.getElementsByTagName('c:view3D')[0];
    if (!view3DNode) {
      view3DNode = createChartElement(doc, 'c:view3D');
      const plotArea = chartNode.getElementsByTagName('c:plotArea')[0];
      if (plotArea) {
        chartNode.insertBefore(view3DNode, plotArea);
      } else {
        chartNode.appendChild(view3DNode);
      }
    }

    const setOrCreateChildVal = (tagName: string, val: number | string) => {
      let child = view3DNode.getElementsByTagName(tagName)[0];
      if (!child) {
        child = createChartElement(doc, tagName);
        view3DNode.appendChild(child);
      }
      child.setAttribute('val', val.toString());
    };

    if (view3DProps.rotX !== undefined) {
      const clampedRotX = Math.max(-90, Math.min(90, view3DProps.rotX));
      setOrCreateChildVal('c:rotX', clampedRotX);
    }
    if (view3DProps.hPercent !== undefined) {
      setOrCreateChildVal('c:hPercent', view3DProps.hPercent);
    }
    if (view3DProps.rotY !== undefined) {
      const clampedRotY = Math.max(0, Math.min(360, view3DProps.rotY));
      setOrCreateChildVal('c:rotY', clampedRotY);
    }
    if (view3DProps.depthPercent !== undefined) {
      setOrCreateChildVal('c:depthPercent', view3DProps.depthPercent);
    }
    if (view3DProps.rAngAx !== undefined) {
      setOrCreateChildVal('c:rAngAx', view3DProps.rAngAx ? '1' : '0');
    }
    if (view3DProps.perspective !== undefined) {
      const clampedPersp = Math.max(0, Math.min(240, view3DProps.perspective));
      setOrCreateChildVal('c:perspective', clampedPersp);
    }

    // Enforce ECMA-376 CT_View3D sequence: rotX,hPercent,rotY,depthPercent,rAngAx,perspective.
    // appendChild moves existing nodes, so re-append in schema order fixes both
    // newly created (appendChild) and pre-existing out-of-order children.
    const view3DOrder = ['c:rotX', 'c:hPercent', 'c:rotY', 'c:depthPercent', 'c:rAngAx', 'c:perspective'];
    for (const tag of view3DOrder) {
      const direct = Array.from(view3DNode.childNodes).find(
        (n): n is Element => n.nodeType === 1 && (n as Element).tagName === tag
      );
      if (direct) view3DNode.appendChild(direct);
    }

    return this.serializer.serializeToString(doc);
  }

  private ensureDocumentRelationship(zipAdapter: DocxZipAdapter, rId: string, type: string, target: string) {
    const relsPath = 'word/_rels/document.xml.rels';
    let relsXml = zipAdapter.readTextOrNull(relsPath);
    if (relsXml) assertNoXxe(relsXml, relsPath);

    if (!relsXml) {
      relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="${rId}" Type="${type}" Target="${target}"/>
</Relationships>`;
    } else {
      const doc = this.domParser.parseFromString(relsXml, 'text/xml');
      const relsRoot = doc.documentElement;
      let found = false;

      const existingRels = Array.from(relsRoot.getElementsByTagName('Relationship'));
      for (const rel of existingRels) {
        if (rel.getAttribute('Id') === rId) {
          rel.setAttribute('Type', type);
          rel.setAttribute('Target', target);
          found = true;
          break;
        }
      }

      if (!found) {
        const newRel = createRelationshipsElement(doc, 'Relationship');
        newRel.setAttribute('Id', rId);
        newRel.setAttribute('Type', type);
        newRel.setAttribute('Target', target);
        relsRoot.appendChild(newRel);
      }
      relsXml = this.serializer.serializeToString(doc);
    }

    zipAdapter.writeText(relsPath, relsXml);
  }

  private ensureContentTypeOverride(zipAdapter: DocxZipAdapter, partName: string, contentType: string) {
    const typesPath = '[Content_Types].xml';
    let typesXml = zipAdapter.readTextOrNull(typesPath);
    if (typesXml) assertNoXxe(typesXml, typesPath);

    if (!typesXml) {
      typesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="${partName}" ContentType="${contentType}"/>
</Types>`;
    } else {
      const doc = this.domParser.parseFromString(typesXml, 'text/xml');
      const typesRoot = doc.documentElement;
      let found = false;

      const overrides = Array.from(typesRoot.getElementsByTagName('Override'));
      for (const ov of overrides) {
        if (ov.getAttribute('PartName') === partName) {
          ov.setAttribute('ContentType', contentType);
          found = true;
          break;
        }
      }

      if (!found) {
        const newOv = createContentTypesElement(doc, 'Override');
        newOv.setAttribute('PartName', partName);
        newOv.setAttribute('ContentType', contentType);
        typesRoot.appendChild(newOv);
      }
      typesXml = this.serializer.serializeToString(doc);
    }

    zipAdapter.writeText(typesPath, typesXml);
  }

  private ensureContentTypeDefault(zipAdapter: DocxZipAdapter, extension: string, contentType: string) {
    const typesPath = '[Content_Types].xml';
    const typesXml = zipAdapter.readTextOrNull(typesPath);
    if (!typesXml) return;
    assertNoXxe(typesXml, typesPath);

    const doc = this.domParser.parseFromString(typesXml, 'text/xml');
    const typesRoot = doc.documentElement;
    const defaults = Array.from(typesRoot.getElementsByTagName('Default'));
    let found = false;
    for (const d of defaults) {
      if (d.getAttribute('Extension')?.toLowerCase() === extension.toLowerCase()) {
        d.setAttribute('ContentType', contentType);
        found = true;
        break;
      }
    }
    if (!found) {
      const newDef = createContentTypesElement(doc, 'Default');
      newDef.setAttribute('Extension', extension);
      newDef.setAttribute('ContentType', contentType);
      typesRoot.insertBefore(newDef, typesRoot.firstChild);
    }
    zipAdapter.writeText(typesPath, this.serializer.serializeToString(doc));
  }
}

/**
 * Read-only surfacing view of an existing chart part (parsed, never mutated).
 * `title`/`chartType` are undefined when the part carries none; categories and
 * series are [] when the part carries no cached points (no placeholders here —
 * callers decide the fallback).
 */
export interface ChartXmlSummary {
  title: string | undefined;
  categories: string[];
  series: Array<{ name: string; values: number[]; xValues?: number[]; unpairedX?: number; unpairedY?: number; chartType?: 'column' | 'line'; axis?: 'primary' | 'secondary' }>;
  chartType: string | undefined;
  style?: ChartStyle;
}

/** Read-only text helper for chart-part surfacing (no mutation). */
function chart3DXmlTextOf(el: Element | null | undefined): string {
  try {
    const t = el?.textContent;
    return typeof t === 'string' ? t : '';
  } catch {
    return '';
  }
}

/**
 * Parses an existing DrawingML 3D chart part into a summary (read path only —
 * no write path in this file is touched or affected). Best-effort: malformed
 * XML yields an empty summary, never a throw.
 */
export function summarize3DChartXml(chartXml: string): ChartXmlSummary {
  const empty: ChartXmlSummary = { title: undefined, categories: [], series: [], chartType: undefined };
  if (typeof chartXml !== 'string' || chartXml.length === 0) return empty;
  try {
    const doc = new DOMParser().parseFromString(chartXml, 'text/xml');
    if (!doc || !doc.documentElement) return empty;
    const root = doc.documentElement;

    let title: string | undefined;
    try {
      const titleEl = chartTitleElement(root);
      if (titleEl) {
        const tEls = titleEl.getElementsByTagName('a:t');
        const parts: string[] = [];
        for (let i = 0; i < tEls.length; i++) parts.push(chart3DXmlTextOf(tEls[i] as Element));
        const joined = parts.join('').trim();
        title = joined ? joined : undefined;
      }
    } catch {
      title = undefined;
    }

    const categories: string[] = [];
    try {
      const catEls = root.getElementsByTagName('c:cat');
      if (catEls.length > 0) {
        const ptEls = (catEls[0] as Element).getElementsByTagName('c:pt');
        for (let i = 0; i < ptEls.length; i++) {
          const vEls = (ptEls[i] as unknown as Element).getElementsByTagName('c:v');
          if (vEls.length > 0) categories.push(chart3DXmlTextOf(vEls[0] as Element));
        }
      }
    } catch {
      /* categories stay as collected */
    }

    const series: Array<{ name: string; values: number[] }> = [];
    try {
      const serEls = root.getElementsByTagName('c:ser');
      for (let s = 0; s < serEls.length; s++) {
        const ser = serEls[s] as Element;
        let name = 'Series';
        try {
          const txEls = ser.getElementsByTagName('c:tx');
          if (txEls.length > 0) {
            const vEls = (txEls[0] as Element).getElementsByTagName('c:v');
            const nm = vEls.length > 0 ? chart3DXmlTextOf(vEls[0] as Element).trim() : '';
            if (nm) name = nm;
          }
        } catch {
          /* name stays default */
        }
        const values: number[] = [];
        try {
          const valEls = ser.getElementsByTagName('c:val');
          if (valEls.length > 0) {
            const ptEls = (valEls[0] as Element).getElementsByTagName('c:pt');
            for (let i = 0; i < ptEls.length; i++) {
              const vEls = (ptEls[i] as unknown as Element).getElementsByTagName('c:v');
              if (vEls.length > 0) values.push(parseFloat(chart3DXmlTextOf(vEls[0] as Element)) || 0);
            }
          }
        } catch {
          /* values stay as collected */
        }
        series.push({ name, values });
      }
    } catch {
      /* series stay as collected */
    }

    let chartType: string | undefined;
    try {
      if (root.getElementsByTagName('c:pie3DChart').length > 0) chartType = 'pie3D';
      else if (root.getElementsByTagName('c:line3DChart').length > 0) chartType = 'line3D';
      else if (root.getElementsByTagName('c:bar3DChart').length > 0) {
        const bars = root.getElementsByTagName('c:bar3DChart');
        let dir = 'col';
        try {
          const dirEl = (bars[0] as Element).getElementsByTagName('c:barDir')[0] as Element | undefined;
          const v = dirEl?.getAttribute('val');
          if (v === 'bar') dir = 'bar';
        } catch {
          /* dir stays col */
        }
        chartType = dir === 'bar' ? 'bar3D' : 'column3D';
      } else if (root.getElementsByTagName('c:area3DChart').length > 0) chartType = 'area3D';
      else if (root.getElementsByTagName('c:surface3DChart').length > 0) chartType = 'surface3D';
    } catch {
      chartType = undefined;
    }
    const style = readChartStyle(chartXml);
    return { title, categories, series, chartType, ...(style !== undefined ? { style } : {}) };
  } catch {
    return empty;
  }
}

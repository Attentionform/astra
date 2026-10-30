/**
 * ASTra OOXML Guideline Bank & RAG Context Engine
 * 
 * Provides structured XML code snippets, OOXML schema guidelines, schema rules,
 * and prompt context helpers for injecting authoritative OOXML templates into LLM system prompts.
 */

export interface OOXMLSnippetItem {
  id: string;
  name: string;
  category: 'headersFooters' | 'pageNumbers' | 'tables' | 'charts' | 'paragraphsRuns' | 'sections' | 'metadataTracked';
  description: string;
  xml: string;
}

export interface SchemaRule {
  id: string;
  title: string;
  rule: string;
  exampleGood: string;
  exampleBad?: string;
}

export const OOXML_NAMESPACES = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  v: 'urn:schemas-microsoft-com:vml',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  w10: 'urn:schemas-microsoft-com:office:word',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
  w15: 'http://schemas.microsoft.com/office/word/2012/wordml',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
} as const;

export const OOXML_SNIPPETS: Record<string, Record<string, OOXMLSnippetItem>> = {
  headersFooters: {
    documentRels: {
      id: 'hdr_ftr_rels',
      name: 'Header/Footer Document Relationships',
      category: 'headersFooters',
      description: 'Relationship entries in word/_rels/document.xml.rels linking header1.xml and footer1.xml',
      xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rIdF1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>`,
    },
    contentTypesOverride: {
      id: 'hdr_ftr_content_types',
      name: 'Header/Footer Content Types',
      category: 'headersFooters',
      description: 'Override tags in [Content_Types].xml for header and footer parts',
      xml: `<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`,
    },
    headerPart: {
      id: 'hdr_part_template',
      name: 'Header Part Template (word/header1.xml)',
      category: 'headersFooters',
      description: 'Complete word/header1.xml structure with document title and bottom accent line',
      xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Header"/>
      <w:jc w:val="right"/>
      <w:pBdr>
        <w:bottom w:val="single" w:sz="6" w:space="4" w:color="003366"/>
      </w:pBdr>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
        <w:sz w:val="18"/>
        <w:color w:val="666666"/>
      </w:rPr>
      <w:t>ASTra Executive Confidential Report</w:t>
    </w:r>
  </w:p>
</w:hdr>`,
    },
    footerPart: {
      id: 'ftr_part_template',
      name: 'Footer Part Template (word/footer1.xml)',
      category: 'headersFooters',
      description: 'Complete word/footer1.xml structure with dynamic Page X of Y numbering',
      xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Footer"/>
      <w:jc w:val="center"/>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:sz w:val="18"/>
        <w:color w:val="888888"/>
      </w:rPr>
      <w:t xml:space="preserve">Page </w:t>
    </w:r>
    <w:fldSimple w:instr="PAGE"/>
    <w:r>
      <w:rPr>
        <w:sz w:val="18"/>
        <w:color w:val="888888"/>
      </w:rPr>
      <w:t xml:space="preserve"> of </w:t>
    </w:r>
    <w:fldSimple w:instr="NUMPAGES"/>
  </w:p>
</w:ftr>`,
    },
    firstPageHeaderPart: {
      id: 'first_hdr_part_template',
      name: 'First Page Header Template (word/header2.xml)',
      category: 'headersFooters',
      description: 'Distinct header for first page when w:titlePg is set in w:sectPr',
      xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Header"/>
      <w:jc w:val="left"/>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:b/>
        <w:sz w:val="16"/>
        <w:color w:val="003366"/>
      </w:rPr>
      <w:t>CONFIDENTIAL - FIRST PAGE</w:t>
    </w:r>
  </w:p>
</w:hdr>`,
    },
    sectPrReferences: {
      id: 'hdr_ftr_sectpr_refs',
      name: 'Section Properties Header/Footer References',
      category: 'headersFooters',
      description: 'w:headerReference and w:footerReference tags placed inside w:sectPr',
      xml: `<w:sectPr>
  <w:headerReference w:type="default" r:id="rIdH1"/>
  <w:footerReference w:type="default" r:id="rIdF1"/>
  <w:pgSz w:w="12240" w:h="15840"/>
  <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
</w:sectPr>`,
    },
    sectPrTitlePgReferences: {
      id: 'hdr_ftr_sectpr_titlepg_refs',
      name: 'Section Properties with First Page Title Header',
      category: 'headersFooters',
      description: 'w:sectPr containing w:titlePg alongside default and first header references',
      xml: `<w:sectPr>
  <w:headerReference w:type="first" r:id="rIdH2"/>
  <w:headerReference w:type="default" r:id="rIdH1"/>
  <w:footerReference w:type="default" r:id="rIdF1"/>
  <w:titlePg/>
  <w:pgSz w:w="12240" w:h="15840"/>
  <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
</w:sectPr>`,
    },
  },
  pageNumbers: {
    pageSimple: {
      id: 'page_simple',
      name: 'Simple PAGE Field',
      category: 'pageNumbers',
      description: 'Dynamic current page number instruction',
      xml: `<w:fldSimple w:instr="PAGE \\* MERGEFORMAT">
  <w:r>
    <w:t>1</w:t>
  </w:r>
</w:fldSimple>`,
    },
    numpagesSimple: {
      id: 'numpages_simple',
      name: 'Simple NUMPAGES Field',
      category: 'pageNumbers',
      description: 'Dynamic total document page count instruction',
      xml: `<w:fldSimple w:instr="NUMPAGES \\* MERGEFORMAT">
  <w:r>
    <w:t>1</w:t>
  </w:r>
</w:fldSimple>`,
    },
    pageXofYParagraph: {
      id: 'page_x_of_y_p',
      name: 'Page X of Y Paragraph Template',
      category: 'pageNumbers',
      description: 'Styled right-aligned paragraph with Page X of Y dynamic fields',
      xml: `<w:p>
  <w:pPr>
    <w:jc w:val="right"/>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
      <w:sz w:val="18"/>
      <w:color w:val="555555"/>
    </w:rPr>
    <w:t xml:space="preserve">Page </w:t>
  </w:r>
  <w:fldSimple w:instr="PAGE"/>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
      <w:sz w:val="18"/>
      <w:color w:val="555555"/>
    </w:rPr>
    <w:t xml:space="preserve"> of </w:t>
  </w:r>
  <w:fldSimple w:instr="NUMPAGES"/>
</w:p>`,
    },
  },
  tables: {
    executiveTable: {
      id: 'executive_table',
      name: 'Executive Styled Table with Zebra Striping',
      category: 'tables',
      description: 'Complete w:tbl structure with header row repeat (w:tblHeader), cantSplit, cell shading, and custom margins',
      xml: `<w:tbl>
  <w:tblPr>
    <w:tblW w:w="5000" w:type="pct"/>
    <w:jc w:val="center"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="12" w:space="0" w:color="003366"/>
      <w:bottom w:val="single" w:sz="12" w:space="0" w:color="003366"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>
      <w:left w:val="none"/>
      <w:right w:val="none"/>
      <w:insideV w:val="none"/>
    </w:tblBorders>
    <w:tblCellMar>
      <w:top w:w="120" w:type="dxa"/>
      <w:bottom w:w="120" w:type="dxa"/>
      <w:left w:w="160" w:type="dxa"/>
      <w:right w:w="160" w:type="dxa"/>
    </w:tblCellMar>
  </w:tblPr>
  <w:tblGrid>
    <w:gridCol w:w="3000"/>
    <w:gridCol w:w="3000"/>
    <w:gridCol w:w="3000"/>
  </w:tblGrid>
  <w:tr>
    <w:trPr>
      <w:tblHeader/>
      <w:cantSplit/>
    </w:trPr>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="3000" w:type="dxa"/>
        <w:shd w:val="clear" w:color="auto" w:fill="003366"/>
      </w:tcPr>
      <w:p>
        <w:pPr><w:jc w:val="left"/></w:pPr>
        <w:r>
          <w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="20"/></w:rPr>
          <w:t>Quarter</w:t>
        </w:r>
      </w:p>
    </w:tc>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="3000" w:type="dxa"/>
        <w:shd w:val="clear" w:color="auto" w:fill="003366"/>
      </w:tcPr>
      <w:p>
        <w:pPr><w:jc w:val="right"/></w:pPr>
        <w:r>
          <w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="20"/></w:rPr>
          <w:t>Revenue ($M)</w:t>
        </w:r>
      </w:p>
    </w:tc>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="3000" w:type="dxa"/>
        <w:shd w:val="clear" w:color="auto" w:fill="003366"/>
      </w:tcPr>
      <w:p>
        <w:pPr><w:jc w:val="right"/></w:pPr>
        <w:r>
          <w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="20"/></w:rPr>
          <w:t>Growth (%)</w:t>
        </w:r>
      </w:p>
    </w:tc>
  </w:tr>
  <w:tr>
    <w:trPr><w:cantSplit/></w:trPr>
    <w:tc>
      <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="F7F9FC"/></w:tcPr>
      <w:p><w:r><w:t>Q1 2026</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="F7F9FC"/></w:tcPr>
      <w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>$12.5</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="F7F9FC"/></w:tcPr>
      <w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>+14.2%</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>`,
    },
  },
  charts: {
    drawingEmbedding: {
      id: 'chart_drawing_embed',
      name: 'Chart Drawing Embedding Snippet',
      category: 'charts',
      description: 'w:drawing element embedding word/charts/chart1.xml via relationship ID',
      xml: `<w:p>
  <w:r>
    <w:drawing>
      <wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="5486400" cy="3200400"/>
        <wp:docPr id="1" name="Chart 1"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                     r:id="rIdC1"/>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
</w:p>`,
    },
    barChart3DSpace: {
      id: 'bar_chart_3d_space',
      name: '3D Bar Chart Space (word/charts/chart1.xml)',
      category: 'charts',
      description: 'Complete c:chartSpace for 3D Bar Chart with c:view3D perspective controls',
      xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:date1904 val="0"/>
  <c:lang val="en-US"/>
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr b="1" sz="1400"/>
              <a:t>Quarterly Revenue Performance</a:t>
            </a:r>
          </a:p>
        </c:rich>
      </c:tx>
    </c:title>
    <c:view3D>
      <c:rotX val="30"/>
      <c:rotY val="20"/>
      <c:rAngAx val="1"/>
      <c:perspective val="30"/>
    </c:view3D>
    <c:plotArea>
      <c:layout/>
      <c:bar3DChart>
        <c:barDir val="col"/>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:v>2026 Sales</c:v></c:tx>
          <c:spPr>
            <a:solidFill>
              <a:srgbClr val="0066CC"/>
            </a:solidFill>
          </c:spPr>
          <c:cat>
            <c:strLit>
              <c:ptCount val="4"/>
              <c:pt idx="0"><c:v>Q1</c:v></c:pt>
              <c:pt idx="1"><c:v>Q2</c:v></c:pt>
              <c:pt idx="2"><c:v>Q3</c:v></c:pt>
              <c:pt idx="3"><c:v>Q4</c:v></c:pt>
            </c:strLit>
          </c:cat>
          <c:val>
            <c:numLit>
              <c:ptCount val="4"/>
              <c:pt idx="0"><c:v>12.5</c:v></c:pt>
              <c:pt idx="1"><c:v>14.8</c:v></c:pt>
              <c:pt idx="2"><c:v>18.2</c:v></c:pt>
              <c:pt idx="3"><c:v>22.0</c:v></c:pt>
            </c:numLit>
          </c:val>
        </c:ser>
        <c:axId val="1001"/>
        <c:axId val="1002"/>
      </c:bar3DChart>
      <c:catAx>
        <c:axId val="1001"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:axPos val="b"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="1002"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="1002"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:axPos val="l"/>
        <c:majorGridlines/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="1001"/>
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:overlay val="0"/>
    </c:legend>
  </c:chart>
</c:chartSpace>`,
    },
    pieChart3DSpace: {
      id: 'pie_chart_3d_space',
      name: '3D Pie Chart Space (word/charts/chart1.xml)',
      category: 'charts',
      description: 'Complete c:chartSpace for 3D Pie Chart with c:pie3DChart and perspective controls',
      xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:date1904 val="0"/>
  <c:lang val="en-US"/>
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr b="1" sz="1400"/>
              <a:t>Market Share Breakdown</a:t>
            </a:r>
          </a:p>
        </c:rich>
      </c:tx>
    </c:title>
    <c:view3D>
      <c:rotX val="30"/>
      <c:rotY val="0"/>
      <c:perspective val="30"/>
    </c:view3D>
    <c:plotArea>
      <c:layout/>
      <c:pie3DChart>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:v>Share</c:v></c:tx>
          <c:cat>
            <c:strLit>
              <c:ptCount val="3"/>
              <c:pt idx="0"><c:v>Product A</c:v></c:pt>
              <c:pt idx="1"><c:v>Product B</c:v></c:pt>
              <c:pt idx="2"><c:v>Product C</c:v></c:pt>
            </c:strLit>
          </c:cat>
          <c:val>
            <c:numLit>
              <c:ptCount val="3"/>
              <c:pt idx="0"><c:v>45</c:v></c:pt>
              <c:pt idx="1"><c:v>35</c:v></c:pt>
              <c:pt idx="2"><c:v>20</c:v></c:pt>
            </c:numLit>
          </c:val>
        </c:ser>
      </c:pie3DChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:overlay val="0"/>
    </c:legend>
  </c:chart>
</c:chartSpace>`,
    },
    lineChart3DSpace: {
      id: 'line_chart_3d_space',
      name: '3D Line Chart Space (word/charts/chart1.xml)',
      category: 'charts',
      description: 'Complete c:chartSpace for 3D Line Chart with c:line3DChart and perspective controls',
      xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:date1904 val="0"/>
  <c:lang val="en-US"/>
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr b="1" sz="1400"/>
              <a:t>Growth Trend Analysis</a:t>
            </a:r>
          </a:p>
        </c:rich>
      </c:tx>
    </c:title>
    <c:view3D>
      <c:rotX val="20"/>
      <c:rotY val="15"/>
      <c:rAngAx val="1"/>
      <c:perspective val="30"/>
    </c:view3D>
    <c:plotArea>
      <c:layout/>
      <c:line3DChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:v>Monthly Active Users</c:v></c:tx>
          <c:cat>
            <c:strLit>
              <c:ptCount val="3"/>
              <c:pt idx="0"><c:v>Jan</c:v></c:pt>
              <c:pt idx="1"><c:v>Feb</c:v></c:pt>
              <c:pt idx="2"><c:v>Mar</c:v></c:pt>
            </c:strLit>
          </c:cat>
          <c:val>
            <c:numLit>
              <c:ptCount val="3"/>
              <c:pt idx="0"><c:v>1000</c:v></c:pt>
              <c:pt idx="1"><c:v>1500</c:v></c:pt>
              <c:pt idx="2"><c:v>2200</c:v></c:pt>
            </c:numLit>
          </c:val>
        </c:ser>
        <c:axId val="1001"/>
        <c:axId val="1002"/>
      </c:line3DChart>
      <c:catAx>
        <c:axId val="1001"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:axPos val="b"/>
        <c:crossAx val="1002"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="1002"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:axPos val="l"/>
        <c:crossAx val="1001"/>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`,
    },
    view3DControls: {
      id: 'view_3d_controls',
      name: '3D View Perspective Controls',
      category: 'charts',
      description: 'c:view3D element detailing rotX, rotY, rAngAx, perspective, hPercent, and depthPercent',
      xml: `<c:view3D>
  <c:rotX val="30"/>
  <c:rotY val="20"/>
  <c:rAngAx val="1"/>
  <c:perspective val="30"/>
  <c:hPercent val="100"/>
  <c:depthPercent val="100"/>
</c:view3D>`,
    },
  },
  paragraphsRuns: {
    heading1: {
      id: 'heading_1_styled',
      name: 'Styled Heading 1 Paragraph',
      category: 'paragraphsRuns',
      description: 'Heading paragraph with font family, size (32 = 16pt), line spacing, and brand color',
      xml: `<w:p>
  <w:pPr>
    <w:pStyle w:val="Heading1"/>
    <w:jc w:val="left"/>
    <w:spacing w:before="240" w:after="120" w:line="288" w:lineRule="auto"/>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii="Helvetica" w:hAnsi="Helvetica"/>
      <w:b/>
      <w:sz w:val="32"/>
      <w:color w:val="003366"/>
    </w:rPr>
    <w:t>Executive Summary</w:t>
  </w:r>
</w:p>`,
    },
    calloutBox: {
      id: 'callout_alert_box',
      name: 'Executive Callout / Alert Box',
      category: 'paragraphsRuns',
      description: 'Border-accented callout box with background shading (w:shd) and left thick border',
      xml: `<w:p>
  <w:pPr>
    <w:pBdr>
      <w:left w:val="single" w:sz="36" w:space="18" w:color="0066CC"/>
    </w:pBdr>
    <w:shd w:val="clear" w:color="auto" w:fill="F0F4F8"/>
    <w:spacing w:before="180" w:after="180" w:line="288" w:lineRule="auto"/>
    <w:ind w:left="240" w:right="240"/>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:b/>
      <w:color w:val="003366"/>
      <w:sz w:val="22"/>
    </w:rPr>
    <w:t>IMPORTANT NOTE: </w:t>
  </w:r>
  <w:r>
    <w:rPr>
      <w:color w:val="333333"/>
      <w:sz w:val="20"/>
    </w:rPr>
    <w:t>All financial figures presented in this report have been audited according to GAAP standards for Q2 2026.</w:t>
  </w:r>
</w:p>`,
    },
  },
  sections: {
    portraitSection: {
      id: 'sectpr_portrait',
      name: 'Portrait Section Properties',
      category: 'sections',
      description: 'w:sectPr for standard Letter portrait page layout (12240 x 15840 dxa)',
      xml: `<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>
  <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
</w:sectPr>`,
    },
    landscapeSection: {
      id: 'sectpr_landscape',
      name: 'Landscape Section Properties',
      category: 'sections',
      description: 'w:sectPr for Letter landscape orientation (15840 x 12240 dxa)',
      xml: `<w:pPr>
  <w:sectPr>
    <w:type w:val="nextPage"/>
    <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
  </w:sectPr>
</w:pPr>`,
    },
  },
  metadataTracked: {
    trackedChanges: {
      id: 'tracked_changes_ins_del',
      name: 'Tracked Insertions and Deletions',
      category: 'metadataTracked',
      description: 'w:ins and w:del inline tracked modification elements',
      xml: `<w:p>
  <w:r>
    <w:t xml:space="preserve">The project deadline is </w:t>
  </w:r>
  <w:del w:id="1" w:author="Sarah Jenkins" w:date="2026-07-25T00:00:00Z">
    <w:r>
      <w:rPr><w:strike/></w:rPr>
      <w:delText>August 1st</w:delText>
    </w:r>
  </w:del>
  <w:ins w:id="2" w:author="Sarah Jenkins" w:date="2026-07-25T00:00:00Z">
    <w:r>
      <w:rPr><w:u w:val="single"/><w:color w:val="008000"/></w:rPr>
      <w:t xml:space="preserve"> September 15th</w:t>
    </w:r>
  </w:ins>
</w:p>`,
    },
    watermarkHeader: {
      id: 'watermark_vml_header',
      name: 'VML Watermark Shape in Header',
      category: 'metadataTracked',
      description: 'Header part containing floating VML shape watermark text',
      xml: `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:v="urn:schemas-microsoft-com:vml"
       xmlns:o="urn:schemas-microsoft-com:office:office">
  <w:p>
    <w:r>
      <w:pict>
        <v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" path="m0@0l@1@2@3@4@5@6@7@8@9@10l21600@0e">
          <v:path textpathok="t"/>
          <v:textpath on="t" fitshape="t"/>
        </v:shapetype>
        <v:shape id="WatermarkShape" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:415.35pt;height:207.65pt;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" fillcolor="#e0e0e0" stroked="f">
          <v:fill opacity="0.5"/>
          <v:textpath string="DRAFT CONFIDENTIAL" style="font-family:&quot;Calibri&quot;;font-weight:bold"/>
        </v:shape>
      </w:pict>
    </w:r>
  </w:p>
</w:hdr>`,
    },
  },
};

export const SCHEMA_RULES: SchemaRule[] = [
  {
    id: 'rule_element_ordering',
    title: 'Strict Element Child Ordering',
    rule: 'Properties elements MUST appear before child run or block elements. In <w:p>, <w:pPr> must be first. In <w:r>, <w:rPr> must be first. In <w:tbl>, <w:tblPr> then <w:tblGrid> must precede <w:tr>. In <w:tc>, <w:tcPr> must precede <w:p>.',
    exampleGood: `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p>`,
    exampleBad: `<w:p><w:r><w:t>Text</w:t></w:r><w:pPr><w:jc w:val="center"/></w:pPr></w:p>`,
  },
  {
    id: 'rule_units_dxa_halfpoints',
    title: 'Measurement Units (DXA, Half-Points & EMUs)',
    rule: 'Font size (<w:sz>, <w:szCs>) is measured in half-points (24 = 12pt, 32 = 16pt). Page sizes, margins, indents, and spacing are measured in DXA (1 inch = 72pt = 1440 dxa). Drawing dimensions (<wp:extent cx="" cy=""/>) are measured in EMUs (1 pt = 12700 EMUs, 1 inch = 914400 EMUs).',
    exampleGood: `<w:rPr><w:sz w:val="28"/></w:rPr> <!-- 14pt -->`,
  },
  {
    id: 'rule_whitespace_preserve',
    title: 'Preserving Text Whitespace',
    rule: 'When text in <w:t> has leading or trailing space characters, xml:space="preserve" MUST be specified on <w:t>.',
    exampleGood: `<w:t xml:space="preserve"> Page </w:t>`,
  },
  {
    id: 'rule_field_structure',
    title: 'Dynamic Field Instruction Format',
    rule: 'Simple dynamic fields use <w:fldSimple w:instr="PAGE"/> or <w:fldSimple w:instr="NUMPAGES"/>. Avoid manually nesting broken run text without proper instructions.',
    exampleGood: `<w:fldSimple w:instr="PAGE \\* MERGEFORMAT"/>`,
  },
  {
    id: 'rule_table_header_cant_split',
    title: 'Table Pagination Controls',
    rule: 'Header rows repeat across pages using <w:tblHeader/> inside <w:trPr>. Prevent row splitting across page breaks using <w:cantSplit/> inside <w:trPr>.',
    exampleGood: `<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>...</w:tr>`,
  },
  {
    id: 'rule_3d_chart_structure',
    title: '3D ChartML (ISO/IEC 29500-1 Clause 21.2) Compliance',
    rule: '3D Chart spaces (<c:chartSpace>) MUST contain <c:chart>. Perspective controls (<c:view3D>) MUST be positioned inside <c:chart> directly before <c:plotArea>. Valid ranges: <c:rotX> val between -90 and 90, <c:rotY> val between 0 and 360, <c:perspective> val between 0 and 240. 3D Bar charts (<c:bar3DChart>) support 3D bar shapes (<c:shape val="box"|"cone"|"cylinder"|"pyramid">).',
    exampleGood: `<c:view3D><c:rotX val="30"/><c:rotY val="20"/><c:rAngAx val="1"/><c:perspective val="30"/></c:view3D>`,
  },
  {
    id: 'rule_2d_chart_default',
    title: '2D-Default Chart Types',
    rule: 'The chart action defaults to the 2D chartType \'column\' (vertical clustered columns, <c:barChart barDir="col">; packages/engine/src/operations/chart.ts:268). Standard 2D types are column | bar | line | pie | doughnut | area | scatter (<c:barChart>, <c:lineChart>, <c:pieChart>, <c:doughnutChart>, <c:areaChart>, <c:scatterChart>; packages/engine/src/manipulators/chart2d.ts:28). Doughnut is a pie with a hole: same series and percent labels inside <c:doughnutChart>, plus <c:holeSize val="50"/> (default hole half the chart diameter). Scatter uses categories as the shared numeric x list, or per-series xValues for one or more series (categories are optional when every series has xValues), and series values as y inside <c:scatterChart> with <c:scatterStyle val="lineMarker"/> and a <c:valAx> pair (packages/engine/src/manipulators/chart2d.ts:343,445). Column/bar variants are columnStacked | barStacked | columnPercentStacked | barPercentStacked (same <c:barChart>, <c:grouping val="stacked"|"percentStacked">, <c:overlap val="100"/>). \'column\' is vertical, \'bar\' is horizontal — both share <c:barChart> and differ only in barDir (packages/engine/src/manipulators/chart2d.ts:367), with axes transposed for horizontal bars (packages/engine/src/manipulators/chart2d.ts:407-408). Prefer 2D \'pie\' (percent labels, per-slice fills) when proportions must read accurately: 3D defaults to rotX=30/perspective=30 (packages/engine/src/manipulators/chart3d.ts:227-230), which foreshortens slices. 3D output is opt-in only via column3D | bar3D | line3D | pie3D (<c:bar3DChart>, <c:line3DChart>, <c:pie3DChart> with <c:view3D>). Multi-series 3D spacing is automatic — the emitter always writes <c:gapWidth val="150"/> + <c:gapDepth val="150"/> (packages/engine/src/manipulators/chart3d.ts:353-354), so never hand-author gap/overlap elements. Series colors are 6-digit hex without \'#\' (packages/engine/src/manipulators/chart3d.ts:166).',
    exampleGood: `{"action":"chart","chartType":"column","title":"Revenue","categories":["Q1","Q2"],"series":[{"name":"ARR","values":[10,25],"color":"0066CC"}]}`,
  },
  {
    id: 'rule_chart_page_budget',
    title: 'Report Page Budget (Font Sizes, One Chart Per Page)',
    rule: 'Style fontSize/sz bare numbers are half-points (w:sz units, ECMA-376 file truth) on every path — no points heuristic, no doubling (numbers pass straight through parseFontSize in packages/engine/src/operation-schema.ts; packages/engine/src/operations/style.ts). Points only via explicit "Npt" strings (e.g. fontSize "11pt" becomes w:sz 22 = 11pt; the bank Heading1 snippet uses w:sz 32 = 16pt). Keep headings at 16pt or below (fontSize "16pt", w:sz <= 32) and body text around 11pt (fontSize "11pt"). Letter portrait with 1in margins leaves a 9in usable budget (15840 - 1440 - 1440 dxa; packages/engine/src/ooxml/guideline-bank.ts:647-650; renderer baseline 864px = 9in, packages/engine/src/parser.ts:113-115): heading + body + table + chart must fit inside it or Word inserts soft breaks. Default charts are ~3in tall (2D: 6 x 3.06in, packages/engine/src/manipulators/chart2d.ts:27-28; 3D: 6 x 3.5in, packages/engine/src/manipulators/chart3d.ts:39-40), so use one chart per page: heading, then chart, then an explicit pageBreak ({action:"insert",type:"pageBreak",anchor:"<id>",position:"after"}; packages/engine/src/operation-schema.ts:61) before the next section.',
    exampleGood: `{"action":"style","target":"p_001","properties":{"fontSize":"11pt","color":"0F172A"}}`,
  },
  {
    id: 'rule_header_footer_relationships',
    title: 'Header/Footer OPC & Section Property Rules',
    rule: 'Header/footer XML parts (word/header1.xml, word/footer1.xml) MUST be registered in word/_rels/document.xml.rels with target relationship IDs and in [Content_Types].xml as overrides. In document.xml, they MUST be referenced in section properties (<w:sectPr>) via <w:headerReference> or <w:footerReference>. When w:type="first" is specified, <w:titlePg/> MUST be included in <w:sectPr>.',
    exampleGood: `<w:sectPr><w:headerReference w:type="first" r:id="rIdH2"/><w:headerReference w:type="default" r:id="rIdH1"/><w:titlePg/></w:sectPr>`,
  },
  {
    id: 'rule_drawing_emubounds',
    title: 'DrawingML Bounds & Extent Bounds',
    rule: 'DrawingML wrappers (<w:drawing><wp:inline>...</wp:inline></w:drawing>) MUST contain <wp:extent cx="..." cy="..."/> measured in EMUs and <wp:docPr id="..." name="..."/>. The <a:graphicData uri="..."> attribute for charts MUST be "http://schemas.openxmlformats.org/drawingml/2006/chart".',
    exampleGood: `<wp:extent cx="5486400" cy="3200400"/>`,
  },
];

/**
 * Get a specific OOXML snippet by category and key name
 */
export function getOOXMLSnippet(category: string, name: string): string | null {
  const catObj = OOXML_SNIPPETS[category];
  if (!catObj) return null;
  const item = catObj[name];
  return item ? item.xml : null;
}

/**
 * Get OOXML schema guidelines formatted for RAG prompt context injection
 */
export function getOOXMLGuidelinesContext(topic?: string): string {
  let context = `=== ASTRA AUTHORITATIVE OOXML SCHEMA GUIDELINES ===\n\n`;

  context += `1. MANDATORY NAMESPACES:\n`;
  context += `   - WordprocessingML: ${OOXML_NAMESPACES.w}\n`;
  context += `   - Relationships: ${OOXML_NAMESPACES.r}\n`;
  context += `   - DrawingML Main: ${OOXML_NAMESPACES.a}\n`;
  context += `   - DrawingML Chart: ${OOXML_NAMESPACES.c}\n`;
  context += `   - Word processing Drawing: ${OOXML_NAMESPACES.wp}\n\n`;

  context += `2. CRITICAL SCHEMA RULES:\n`;
  for (const r of SCHEMA_RULES) {
    context += `   * [${r.title}]: ${r.rule}\n`;
  }
  context += `\n`;

  if (topic && OOXML_SNIPPETS[topic]) {
    context += `=== RELEVANT TEMPLATES FOR CATEGORY: ${topic.toUpperCase()} ===\n\n`;
    const cat = OOXML_SNIPPETS[topic];
    for (const [key, snippet] of Object.entries(cat)) {
      context += `--- Template: ${snippet.name} (${key}) ---\n`;
      context += `${snippet.description}\n`;
      context += `${snippet.xml}\n\n`;
    }
  } else {
    context += `=== KEY TEMPLATES SUMMARY ===\n`;
    context += `- Headers/Footers: Use word/header1.xml and word/footer1.xml linked via _rels/document.xml.rels and referenced in w:sectPr (<w:headerReference>, <w:footerReference>). Use <w:titlePg/> for first page headers.\n`;
    context += `- Page Numbers: Use <w:fldSimple w:instr="PAGE"/> and <w:fldSimple w:instr="NUMPAGES"/>.\n`;
    context += `- Tables: Use <w:tblHeader/> and <w:cantSplit/> with <w:shd w:fill="HEX"/> for zebra striping.\n`;
    context += `- Charts (2D default): the chart action defaults to chartType 'column' (2D clustered vertical columns, <c:barChart barDir="col">). 2D types: column (vertical) | bar (horizontal, axes transposed) | columnStacked | barStacked | columnPercentStacked | barPercentStacked | line | pie | doughnut (pie with <c:holeSize val="50"/>) | area | scatter (shared numeric x categories or per-series xValues + y values in <c:scatterChart>, two <c:valAx> axes; a categories-only update moves series that share the first x list). Prefer 2D pie when proportions must read accurately (3D perspective foreshortens slices). Opt-in 3D: column3D | bar3D | line3D | pie3D, which emit <c:bar3DChart>/<c:line3DChart>/<c:pie3DChart> with <c:view3D> (<c:rotX>, <c:rotY>, <c:rAngAx>, <c:perspective>) before <c:plotArea> (multi-series 3D gapWidth/gapDepth spacing is automatic — never hand-author it). Series colors: 6-digit hex without '#'.\n`;
    context += `- Report page budget: style fontSize/sz bare numbers are half-points (w:sz units, ECMA-376 file truth — no doubling); points only via explicit "Npt" strings (e.g. "11pt" -> w:sz 22 = 11pt). Keep headings <= 16pt and body ~11pt. Letter portrait with 1in margins leaves 9in usable, and default charts are ~3in tall — use one chart per page (heading, chart, explicit pageBreak) or expect soft breaks.\n`;
    context += `- Callouts: Use <w:pBdr> with <w:left w:sz="36"> and <w:shd w:fill="HEX"/> background.\n`;
  }

  return context;
}

/**
 * Builds standard system prompt context block containing all schema rules & templates
 */
export function buildSystemPromptContext(): string {
  return getOOXMLGuidelinesContext();
}


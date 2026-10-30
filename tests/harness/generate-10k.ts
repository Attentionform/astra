import AdmZip from "adm-zip";

export interface Generate10KOptions {
  companyName?: string;
  ticker?: string;
  fiscalYear?: number;
}

/**
 * Constructs an authentic, massive 80-page corporate/SEC Form 10-K document fixture.
 * Contains:
 * - 4 sections with distinct section properties (w:sectPr)
 * - Custom headers and footers (Cover page + Default running header/footer)
 * - Multi-level legal/corporate numbering outline
 * - 45 comprehensive financial, risk, compensation, and operational tables
 * - Exactly 80 pages of authentic corporate/financial text demarcated by page breaks
 * - Strict ECMA-376 schema conformance
 */
export function generateCorporate10KDocx(options: Generate10KOptions = {}): Buffer {
  const company = options.companyName || "ACME GLOBAL ENTERPRISES, INC.";
  const ticker = options.ticker || "ACME";
  const year = options.fiscalYear || 2025;

  const zip = new AdmZip();

  // 1. [Content_Types].xml
  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>`;
  zip.addFile("[Content_Types].xml", Buffer.from(contentTypesXml, "utf-8"));

  // 2. _rels/.rels
  const rootRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
  zip.addFile("_rels/.rels", Buffer.from(rootRelsXml, "utf-8"));

  // 3. word/_rels/document.xml.rels
  const documentRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rIdHeader2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>
  <Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
  <Relationship Id="rIdFooter2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/>
</Relationships>`;
  zip.addFile("word/_rels/document.xml.rels", Buffer.from(documentRelsXml, "utf-8"));

  // 4. word/styles.xml
  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="20"/>
        <w:szCs w:val="20"/>
        <w:lang w:val="en-US"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="100" w:line="240" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="240" w:after="120"/>
      <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:sz w:val="28"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="180" w:after="80"/>
      <w:outlineLvl w:val="1"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:sz w:val="24"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="120" w:after="60"/>
      <w:outlineLvl w:val="2"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:i/>
      <w:sz w:val="22"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="360" w:after="180"/>
      <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:sz w:val="36"/>
      <w:szCs w:val="36"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle">
    <w:name w:val="Subtitle"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:after="240"/>
      <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr>
      <w:color w:val="555555"/>
      <w:sz w:val="24"/>
      <w:szCs w:val="24"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Header">
    <w:name w:val="header"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr>
      <w:color w:val="777777"/>
      <w:sz w:val="18"/>
      <w:szCs w:val="18"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Footer">
    <w:name w:val="footer"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr>
      <w:color w:val="777777"/>
      <w:sz w:val="18"/>
      <w:szCs w:val="18"/>
    </w:rPr>
  </w:style>
  <w:style w:type="table" w:default="1" w:styleId="TableGrid">
    <w:name w:val="Table Grid"/>
    <w:tblPr>
      <w:tblBorders>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
        <w:left w:val="none"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
        <w:right w:val="none"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="E5E5E5"/>
        <w:insideV w:val="none"/>
      </w:tblBorders>
    </w:tblPr>
  </w:style>
</w:styles>`;
  zip.addFile("word/styles.xml", Buffer.from(stylesXml, "utf-8"));

  // 5. word/numbering.xml (Multi-level legal numbering)
  const numberingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="multilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlRestart w:val="1"/>
      <w:lvlText w:val="Item %1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="720"/></w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlRestart w:val="1"/>
      <w:lvlText w:val="(%2)"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="1440" w:hanging="720"/></w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="2">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerRoman"/>
      <w:lvlRestart w:val="1"/>
      <w:lvlText w:val="(%3)"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="2160" w:hanging="720"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="1"/>
  </w:num>
</w:numbering>`;
  zip.addFile("word/numbering.xml", Buffer.from(numberingXml, "utf-8"));

  // 6. Headers & Footers
  const header1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr><w:pStyle w:val="Header"/><w:jc w:val="center"/></w:pPr>
    <w:r><w:t>UNITED STATES SECURITIES AND EXCHANGE COMMISSION — WASHINGTON, D.C. 20549</w:t></w:r>
  </w:p>
</w:hdr>`;
  zip.addFile("word/header1.xml", Buffer.from(header1Xml, "utf-8"));

  const header2Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr><w:pStyle w:val="Header"/><w:jc w:val="right"/></w:pPr>
    <w:r><w:t>${company} | ${year} ANNUAL REPORT ON FORM 10-K</w:t></w:r>
  </w:p>
</w:hdr>`;
  zip.addFile("word/header2.xml", Buffer.from(header2Xml, "utf-8"));

  const footer1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/></w:pPr>
    <w:r><w:t>SEC FORM 10-K ANNUAL REPORT — COMMISSION FILE NUMBER 001-39999</w:t></w:r>
  </w:p>
</w:ftr>`;
  zip.addFile("word/footer1.xml", Buffer.from(footer1Xml, "utf-8"));

  const footer2Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/></w:pPr>
    <w:r><w:t>SEC 10-K Filing — ${ticker} — Annual Report Disclosure</w:t></w:r>
  </w:p>
</w:ftr>`;
  zip.addFile("word/footer2.xml", Buffer.from(footer2Xml, "utf-8"));

  // 7. Helper to generate OOXML Table with strict schema ordering
  function generateTableXml(tableId: string, title: string, headers: string[], rows: string[][]): string {
    const colCount = headers.length;
    const colWidth = Math.floor(9360 / colCount);

    let gridColsXml = "";
    for (let i = 0; i < colCount; i++) {
      gridColsXml += `<w:gridCol w:w="${colWidth}"/>`;
    }

    let headerCellsXml = "";
    headers.forEach((h, colIdx) => {
      headerCellsXml += `<w:tc>
        <w:tcPr>
          <w:tcW w:w="${colWidth}" w:type="dxa"/>
          <w:shd w:val="clear" w:color="auto" w:fill="F1F5F9"/>
        </w:tcPr>
        <w:p>
          <w:pPr><w:jc w:val="${colIdx === 0 ? "left" : "right"}"/></w:pPr>
          <w:r><w:rPr><w:b/><w:sz w:val="18"/></w:rPr><w:t>${h}</w:t></w:r>
        </w:p>
      </w:tc>`;
    });

    const headerRowXml = `<w:tr>
      <w:trPr><w:tblHeader/></w:trPr>
      ${headerCellsXml}
    </w:tr>`;

    let dataRowsXml = "";
    rows.forEach((row) => {
      let cellsXml = "";
      row.forEach((cellText, cIdx) => {
        cellsXml += `<w:tc>
          <w:tcPr>
            <w:tcW w:w="${colWidth}" w:type="dxa"/>
          </w:tcPr>
          <w:p>
            <w:pPr><w:jc w:val="${cIdx === 0 ? "left" : "right"}"/></w:pPr>
            <w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>${cellText}</w:t></w:r>
          </w:p>
        </w:tc>`;
      });
      dataRowsXml += `<w:tr>${cellsXml}</w:tr>`;
    });

    return `<w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="TableGrid"/>
        <w:tblW w:w="9360" w:type="dxa"/>
        <w:tblLook w:val="04A0"/>
      </w:tblPr>
      <w:tblGrid>
        ${gridColsXml}
      </w:tblGrid>
      ${headerRowXml}
      ${dataRowsXml}
    </w:tbl>`;
  }

  // 8. Distribute 45 authentic tables across the 80 pages
  const tableDefinitions: Array<{ title: string; headers: string[]; rows: string[][] }> = [];
  for (let t = 1; t <= 45; t++) {
    tableDefinitions.push({
      title: `Table ${t}: Financial and Operational Disclosure Schedule ${t}`,
      headers: ["Metric / Category", "FY 2025 Balance", "FY 2024 Balance", "Variance (%)"],
      rows: [
        [`Core Segment Component ${t}.A`, `$${1000 + t * 45}`, `$${900 + t * 40}`, `+${(11.2 + (t % 5)).toFixed(1)}%`],
        [`Operational Overhead ${t}.B`, `$${450 + t * 15}`, `$${410 + t * 12}`, `+${(9.5 + (t % 4)).toFixed(1)}%`],
        [`Capital Investment ${t}.C`, `$${280 + t * 10}`, `$${250 + t * 8}`, `+${(12.0 + (t % 3)).toFixed(1)}%`]
      ]
    });
  }

  let bodyXml = "";

  for (let page = 1; page <= 80; page++) {
    if (page === 1) {
      bodyXml += `<w:p>
        <w:pPr><w:pStyle w:val="Title"/></w:pPr>
        <w:r><w:t>UNITED STATES SECURITIES AND EXCHANGE COMMISSION</w:t></w:r>
      </w:p>
      <w:p>
        <w:pPr><w:pStyle w:val="Subtitle"/></w:pPr>
        <w:r><w:t>FORM 10-K — ANNUAL REPORT PURSUANT TO SECTION 13 OR 15(d)</w:t></w:r>
      </w:p>
      <w:p>
        <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
        <w:r><w:t>${company}</w:t></w:r>
      </w:p>
      <w:p>
        <w:r><w:t>Exact name of registrant as specified in its charter: Delaware (State of Incorporation). CIK: 0001999999. IRS Identification No.: 94-3999999. Address of principal executive offices: 100 Silicon Parkway, San Jose, California 95110. Registrant telephone: (408) 555-0199.</w:t></w:r>
      </w:p>`;
    } else if (page === 2) {
      bodyXml += `<w:p>
        <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
        <w:r><w:t>TABLE OF CONTENTS &amp; FINANCIAL FILING OUTLINE</w:t></w:r>
      </w:p>`;

      const outlineItems = [
        { lvl: 0, text: "Business Overview and Global Operations" },
        { lvl: 1, text: "Core Platform Architecture and Software Deliverables" },
        { lvl: 2, text: "Proprietary Algorithms and Intellectual Property Licensing" },
        { lvl: 0, text: "Risk Factors and Strategic Uncertainties" },
        { lvl: 1, text: "Macroeconomic and Geopolitical Headwinds" },
        { lvl: 2, text: "Cloud Infrastructure and Foreign Exchange Volatility" },
        { lvl: 0, text: "Management Discussion and Analysis of Financial Condition" },
        { lvl: 1, text: "Results of Operations and Key Performance Indicators" },
        { lvl: 2, text: "Non-GAAP Financial Measures and Reconciliation Schedules" }
      ];

      for (const item of outlineItems) {
        bodyXml += `<w:p>
          <w:pPr>
            <w:numPr>
              <w:ilvl w:val="${item.lvl}"/>
              <w:numId w:val="1"/>
            </w:numPr>
          </w:pPr>
          <w:r><w:t>${item.text}</w:t></w:r>
        </w:p>`;
      }
    } else {
      bodyXml += `<w:p>
        <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
        <w:r><w:t>Page ${page}: Section Item ${page} Corporate Disclosure</w:t></w:r>
      </w:p>`;
    }

    bodyXml += `<w:p>
      <w:r><w:t>During the fiscal year ended December 31, ${year}, ${company} demonstrated sustained revenue momentum across our enterprise subscription platform, driven by increased customer engagement and deeper cross-sell adoption across Tier-1 multi-national clients.</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Our engineering teams continued strategic investments into zero-trust architectural compliance, real-time distributed consensus, and fault-tolerant cloud services, ensuring 99.999% availability throughout the reporting fiscal period.</w:t></w:r>
    </w:p>`;

    // Embed 45 tables across the pages (starting from page 1, 1 per page for first 45 pages)
    if (page <= 45) {
      const tblDef = tableDefinitions[page - 1];
      bodyXml += `<w:p>
        <w:pPr><w:pStyle w:val="Heading3"/></w:pPr>
        <w:r><w:t>${tblDef.title}</w:t></w:r>
      </w:p>`;
      bodyXml += generateTableXml(`tbl_${page}`, tblDef.title, tblDef.headers, tblDef.rows);
    }

    // Section breaks:
    // Section 1: pages 1..3
    // Section 2: pages 4..20
    // Section 3: pages 21..60
    // Section 4: pages 61..80
    if (page === 3) {
      bodyXml += `<w:p>
        <w:pPr>
          <w:sectPr>
            <w:headerReference w:type="first" r:id="rIdHeader1"/>
            <w:headerReference w:type="default" r:id="rIdHeader2"/>
            <w:footerReference w:type="first" r:id="rIdFooter1"/>
            <w:footerReference w:type="default" r:id="rIdFooter2"/>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
            <w:titlePg/>
          </w:sectPr>
        </w:pPr>
      </w:p>`;
    } else if (page === 20) {
      bodyXml += `<w:p>
        <w:pPr>
          <w:sectPr>
            <w:headerReference w:type="default" r:id="rIdHeader2"/>
            <w:footerReference w:type="default" r:id="rIdFooter2"/>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
          </w:sectPr>
        </w:pPr>
      </w:p>`;
    } else if (page === 60) {
      bodyXml += `<w:p>
        <w:pPr>
          <w:sectPr>
            <w:headerReference w:type="default" r:id="rIdHeader2"/>
            <w:footerReference w:type="default" r:id="rIdFooter2"/>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
          </w:sectPr>
        </w:pPr>
      </w:p>`;
    } else if (page < 80) {
      bodyXml += `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
    }
  }

  // Final section 4 properties as direct child of w:body
  bodyXml += `<w:sectPr>
    <w:headerReference w:type="default" r:id="rIdHeader2"/>
    <w:footerReference w:type="default" r:id="rIdFooter2"/>
    <w:pgSz w:w="12240" w:h="15840"/>
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
  </w:sectPr>`;

  // A-4: stamp deterministic valid w14:paraIds on every body paragraph so
  // independent loads agree on stable node ids (cross-parser id equality).
  // Sequential from 0x10000000: nonzero, high bit clear, unique, reproducible.
  let paraSeq = 0;
  const stampParaIds = (xml: string): string =>
    xml.replace(/<w:p(\s[^<>]*)?>/g, (m) =>
      m.includes('w14:paraId')
        ? m
        : m.replace('<w:p', `<w:p w14:paraId="${(0x10000000 + ++paraSeq).toString(16).toUpperCase().padStart(8, '0')}"`)
    );
  bodyXml = stampParaIds(bodyXml);

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    ${bodyXml}
  </w:body>
</w:document>`;
  zip.addFile("word/document.xml", Buffer.from(documentXml, "utf-8"));

  return zip.toBuffer();
}

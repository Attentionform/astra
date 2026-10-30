import { describe, it, expect, beforeEach, afterEach } from "vitest";
import {
  DocumentManager,
  DocxParser,
  validateDocxBuffer,
} from '@attentionform/astra-engine';

import { generateCorporate10KDocx } from "./generate-10k";
import { DOMParser } from "@xmldom/xmldom";
import AdmZip from "adm-zip";
import fs from "fs";
import os from "os";
import path from "path";

describe("HARNESS-10K: The 80-Page Corporate 10-K Roundtrip Test Suite", () => {
  let tmpDir: string;
  let fixturePath: string;
  let rawDocxBuf: Buffer;

  beforeEach(() => {
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "astra-10k-harness-"));
    fixturePath = path.join(tmpDir, "corporate_10k_filing.docx");
    rawDocxBuf = generateCorporate10KDocx({
      companyName: "ACME GLOBAL ENTERPRISES, INC.",
      ticker: "ACME",
      fiscalYear: 2025
    });
    fs.writeFileSync(fixturePath, rawDocxBuf);
  });

  afterEach(() => {
    if (tmpDir && fs.existsSync(tmpDir)) {
      try {
        fs.rmSync(tmpDir, { recursive: true, force: true });
      } catch { /* ignore cleanup failures */ }
    }
  });

  it("fixture verification: generates valid 80-page 10-K with 40+ tables, custom headers/footers, and numbering", async () => {
    // 1. Validate raw fixture with native .NET OOXML validator
    const rawVal = await validateDocxBuffer(rawDocxBuf);
    expect(rawVal.ok).toBe(true);
    expect(rawVal.errors).toEqual([]);

    // 2. Open with DocxParser and verify structural composition
    const parser = new DocxParser(rawDocxBuf);
    const summary = parser.getStructuralSummary();
    const ast = await parser.getAST();
    const dom = parser.getDocumentDom();

    // Verify 4 distinct sections configured via w:sectPr
    const sectPrs = dom.getElementsByTagName("w:sectPr");
    expect(sectPrs.length).toBe(4);

    // Verify 40+ tables exist in the document (our fixture generates 45 tables)
    const tables = summary.nodes.filter(n => n.type === "table");
    expect(tables.length).toBeGreaterThanOrEqual(40);
    expect(tables.length).toBe(45);

    // Verify multi-level numbering exists
    expect(ast.numbering).toBeDefined();
    expect(Object.keys(ast.numbering.abstractNum || {}).length).toBeGreaterThanOrEqual(1);
    expect(Object.keys(ast.numbering.concreteNum || {}).length).toBeGreaterThanOrEqual(1);

    // Verify custom headers and footers in the zip package
    const zip = new AdmZip(rawDocxBuf);
    expect(zip.getEntry("word/header1.xml")).not.toBeNull();
    expect(zip.getEntry("word/header2.xml")).not.toBeNull();
    expect(zip.getEntry("word/footer1.xml")).not.toBeNull();
    expect(zip.getEntry("word/footer2.xml")).not.toBeNull();

    const header1Xml = zip.getEntry("word/header1.xml")!.getData().toString("utf-8");
    expect(header1Xml).toContain("SECURITIES AND EXCHANGE COMMISSION");

    const header2Xml = zip.getEntry("word/header2.xml")!.getData().toString("utf-8");
    expect(header2Xml).toContain("ACME GLOBAL ENTERPRISES, INC.");
    expect(header2Xml).toContain("FORM 10-K");
  }, 120_000);

  it("roundtrip edit & fidelity: modifies 2 paragraphs and 1 table cell, saves with applyTypography: false", async () => {
    // Record baseline AST before any edits
    const baselineParser = new DocxParser(rawDocxBuf);
    const baselineAst = await baselineParser.getAST();

    // 1. Open with DocumentManager
    const dm = new DocumentManager();
    await dm.openDocument(fixturePath);
    const parser = dm.getParser();
    const summary = parser.getStructuralSummary();

    // Locate 2 narrative paragraphs in Section 2 (Part I)
    const qualifyingParas = summary.nodes.filter(
      n => n.type === "paragraph" && n.text && n.text.includes("revenue momentum across our enterprise subscription platform")
    );
    expect(qualifyingParas.length).toBeGreaterThanOrEqual(2);
    const targetP1 = qualifyingParas[0];
    const targetP2 = qualifyingParas[1];

    const updatedTextP1 = "MODIFIED PARAGRAPH 1: ACME reported record ARR acceleration of 38% across Cloud Platforms.";
    const updatedTextP2 = "MODIFIED PARAGRAPH 2: Operating margin expansion exceeded 420 basis points in the fiscal period.";

    parser.updateNode(targetP1.id, `<w:p><w:r><w:t>${updatedTextP1}</w:t></w:r></w:p>`);
    parser.updateNode(targetP2.id, `<w:p><w:r><w:t>${updatedTextP2}</w:t></w:r></w:p>`);

    // Locate 1 table cell in Section 2 (Table 1) and modify it
    const tableNodes = summary.nodes.filter(n => n.type === "table");
    expect(tableNodes.length).toBeGreaterThanOrEqual(40);
    const targetTable = tableNodes[0];

    const tblDomNode = parser.getNode(targetTable.id)!;
    const cells = tblDomNode.getElementsByTagName("w:tc");
    expect(cells.length).toBeGreaterThanOrEqual(4);

    const cellToModify = cells[1];
    const pInCell = cellToModify.getElementsByTagName("w:p")[0];
    const tInCell = pInCell.getElementsByTagName("w:t")[0];
    const updatedCellText = "$999,999,999 (AMENDED)";
    tInCell.textContent = updatedCellText;

    // 2. Save with applyTypography: false
    const savePath = path.join(tmpDir, "corporate_10k_no_typography.docx");
    const saveRes = await dm.saveAsDocument(savePath, { applyTypography: false });

    // Assert zero typography fixes were applied
    expect(saveRes.typographyFixes.length).toBe(0);

    // 3. Assert validateDocxBuffer returns ok: true
    const savedBuf = fs.readFileSync(savePath);
    const validation = await validateDocxBuffer(savedBuf);
    expect(validation.ok).toBe(true);
    expect(validation.errors).toEqual([]);

    // 4. Assert data-astra-id is stripped completely: unzipping word/document.xml must find ZERO occurrences
    const savedZip = new AdmZip(savedBuf);
    const documentXml = savedZip.getEntry("word/document.xml")!.getData().toString("utf-8");
    const occurrencesOfDataAstraId = (documentXml.match(/data-astra-id/g) || []).length;
    expect(occurrencesOfDataAstraId).toBe(0);
    expect(documentXml).not.toContain("data-astra-id");

    // 5. Assert untouched sections maintain identical structure
    const reloadedParser = new DocxParser(savedBuf);
    const reloadedAst = await reloadedParser.getAST();
    const reloadedDom = reloadedParser.getDocumentDom();

    // Verify all 4 sections are preserved in document XML
    const reloadedSectPrs = reloadedDom.getElementsByTagName("w:sectPr");
    expect(reloadedSectPrs.length).toBe(4);

    // Verify total node count is preserved
    expect(reloadedAst.sections[0].nodes.length).toBe(baselineAst.sections[0].nodes.length);

    // Compare all nodes: every untouched node must have identical type and text
    for (let i = 0; i < baselineAst.sections[0].nodes.length; i++) {
      const baseNode = baselineAst.sections[0].nodes[i];
      const relNode = reloadedAst.sections[0].nodes[i];
      expect(relNode.type).toBe(baseNode.type);

      if (baseNode.id === targetP1.id) {
        expect(relNode.metadata.text).toBe(updatedTextP1);
      } else if (baseNode.id === targetP2.id) {
        expect(relNode.metadata.text).toBe(updatedTextP2);
      } else if (baseNode.id === targetTable.id) {
        expect(relNode.metadata.text).toContain(updatedCellText);
      } else {
        expect(relNode.metadata.text).toBe(baseNode.metadata.text);
      }
    }

    // Verify the targeted modifications took effect
    expect(documentXml).toContain(updatedTextP1);
    expect(documentXml).toContain(updatedTextP2);
    expect(documentXml).toContain(updatedCellText);
  }, 120_000);

  it("typography guard: saves with applyTypography: true and enforces keepNext on headings and cantSplit on table rows", async () => {
    // 1. Open with DocumentManager
    const dm = new DocumentManager();
    await dm.openDocument(fixturePath);
    const parser = dm.getParser();
    const summary = parser.getStructuralSummary();

    // Modify 2 paragraphs and 1 table cell
    const qualifyingParas = summary.nodes.filter(
      n => n.type === "paragraph" && n.text && n.text.includes("revenue momentum across our enterprise subscription platform")
    );
    parser.updateNode(qualifyingParas[0].id, "<w:p><w:r><w:t>TYPOGRAPHY MODIFIED P1: Enhanced enterprise cloud architecture.</w:t></w:r></w:p>");
    parser.updateNode(qualifyingParas[1].id, "<w:p><w:r><w:t>TYPOGRAPHY MODIFIED P2: Strong positive operating leverage.</w:t></w:r></w:p>");

    const tableNodes = summary.nodes.filter(n => n.type === "table");
    const tblDomNode = parser.getNode(tableNodes[0].id)!;
    const cells = tblDomNode.getElementsByTagName("w:tc");
    cells[1].getElementsByTagName("w:p")[0].getElementsByTagName("w:t")[0].textContent = "$888,888 (TYPOGRAPHY REVISED)";

    // 2. Save with applyTypography: true
    const savePath = path.join(tmpDir, "corporate_10k_with_typography.docx");
    const saveRes = await dm.saveAsDocument(savePath, { applyTypography: true });

    // Assert typography fixes were applied (commercial publication rules)
    expect(saveRes.typographyFixes.length).toBeGreaterThan(0);

    // 3. Assert validateDocxBuffer returns ok: true
    const savedBuf = fs.readFileSync(savePath);
    const validation = await validateDocxBuffer(savedBuf);
    expect(validation.ok).toBe(true);
    expect(validation.errors).toEqual([]);

    // 4. Assert w:keepNext on headings and w:cantSplit on table rows
    const savedZip = new AdmZip(savedBuf);
    const documentXml = savedZip.getEntry("word/document.xml")!.getData().toString("utf-8");

    expect(documentXml).toContain("<w:keepNext/>");
    expect(documentXml).toContain("<w:cantSplit/>");

    // Parse the saved document with DOMParser
    const docDom = new DOMParser().parseFromString(documentXml, "text/xml");

    // Find all paragraphs with w:pStyle matching Heading1, Heading2, Heading3, Title, Subtitle and assert that their <w:pPr> contains <w:keepNext/>
    const pNodes = Array.from(docDom.getElementsByTagName("w:p"));
    const headingParas = pNodes.filter(p => {
      const pPr = p.getElementsByTagName("w:pPr")[0];
      if (!pPr) return false;
      const pStyle = pPr.getElementsByTagName("w:pStyle")[0];
      if (!pStyle) return false;
      const val = pStyle.getAttribute("w:val") || "";
      return ["Heading1", "Heading2", "Heading3", "Title", "Subtitle"].includes(val);
    });
    expect(headingParas.length).toBeGreaterThan(0);
    for (const p of headingParas) {
      const pPr = p.getElementsByTagName("w:pPr")[0];
      expect(pPr).toBeDefined();
      const keepNext = pPr.getElementsByTagName("w:keepNext");
      expect(keepNext.length).toBeGreaterThan(0);
    }

    // Find all tables and assert that their <w:trPr> elements contain <w:cantSplit/>
    const tblNodes = Array.from(docDom.getElementsByTagName("w:tbl"));
    expect(tblNodes.length).toBeGreaterThanOrEqual(40);
    for (const tbl of tblNodes) {
      const trNodes = Array.from(tbl.getElementsByTagName("w:tr"));
      expect(trNodes.length).toBeGreaterThan(0);
      for (const tr of trNodes) {
        const trPr = tr.getElementsByTagName("w:trPr")[0];
        expect(trPr).toBeDefined();
        const cantSplit = trPr.getElementsByTagName("w:cantSplit");
        expect(cantSplit.length).toBeGreaterThan(0);
      }
    }

    // 5. Assert data-astra-id is stripped completely
    expect(documentXml).not.toContain("data-astra-id");
    const occurrencesOfDataAstraId = (documentXml.match(/data-astra-id/g) || []).length;
    expect(occurrencesOfDataAstraId).toBe(0);

    // Verify content modifications persisted
    expect(documentXml).toContain("TYPOGRAPHY MODIFIED P1");
    expect(documentXml).toContain("TYPOGRAPHY MODIFIED P2");
    expect(documentXml).toContain("$888,888 (TYPOGRAPHY REVISED)");
  }, 120_000);
});
